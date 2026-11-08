"""Ordered Excel number queue; no QQ, settings, or network access."""
import re

from task_control import atomic_save, check_stop


def number_text(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip() if value is not None else ''


def write_row_markers(sheet, rows, marker, prefix, *, reason=None, mark_processed=True):
    for row in rows:
        if mark_processed:
            sheet.cell(row=row, column=3).value = marker
        if reason or mark_processed:
            cell = sheet.cell(row=row, column=4)
            existing = str(cell.value or '').strip()
            lines = [line for line in existing.splitlines() if not line.startswith(prefix)]
            if reason:
                lines.append(prefix + reason)
            cell.value = '\n'.join(lines) or None


class ExcelNumberQueue:
    def __init__(self, document, path, marker, reason_prefix):
        self.document = document
        self.path = path
        self.marker = marker
        self.reason_prefix = reason_prefix
        self.sheets = list(document.worksheets)
        if not self.sheets:
            raise ValueError('Excel 中没有数据工作表。')
        self.entries = {}
        self.locations = {}
        self.owners = {}
        self.completed = set()
        self.reasons = {}
        errors = {}
        for sheet in self.sheets:
            check_stop()
            entries = {}
            invalid = []
            for row, (cell, group, status) in enumerate(
                    sheet.iter_rows(min_col=1, max_col=3), start=1):
                check_stop()
                value = number_text(cell.value)
                if not value:
                    continue
                is_number = re.fullmatch(r'[0-9]{5,12}', value) is not None
                # Capture output starts each group with its name and group ID.
                if (cell.data_type != 'f' and isinstance(cell.value, str)
                        and not is_number and re.fullmatch(r'[0-9]{5,12}', number_text(group.value))):
                    continue
                if cell.data_type == 'f' or not is_number:
                    invalid.append((row, cell.value, cell.data_type == 'f'))
                    continue
                entries.setdefault(value, []).append(row)
                self.locations.setdefault(value, []).append((sheet.title, row))
                self.owners.setdefault(value, sheet.title)
                if str(status.value or '').strip() == marker:
                    self.completed.add(value)
                    reason = next((line[len(reason_prefix):] for line in
                                   str(sheet.cell(row, 4).value or '').splitlines()
                                   if line.startswith(reason_prefix)), None)
                    self.reasons.setdefault(value, reason)
            self.entries[sheet.title] = entries
            errors[sheet.title] = invalid
        for sheet in self.sheets:
            invalid = errors[sheet.title]
            if not invalid:
                continue
            # Preserve standalone text-only cover/notes sheets in multi-sheet files.
            # A numeric-looking value, formula, or any invalid row among QQ numbers
            # remains an error rather than silently losing an intended contact.
            text_only = all(isinstance(value, str) and not formula
                            and not re.search(r'\d', value) for _, value, formula in invalid)
            if not self.entries[sheet.title] and self.locations and text_only:
                continue
            row, _, formula = invalid[0]
            detail = ('是公式，请粘贴为实际 QQ 号码值。' if formula else
                      '不是有效 QQ 号，请删除表头或修正号码。')
            raise ValueError(f'工作表“{sheet.title}”第 {row} 行 A 列{detail}')

    @property
    def first_sheet(self):
        for number in self.locations:
            if number not in self.completed:
                return self.document[self.owners[number]]
        return self.sheets[0]

    def pending_numbers(self):
        return [(number, self.entries[self.owners[number]][number])
                for number in self.locations if number not in self.completed]

    def pending_in_sheet(self, title):
        return [(number, rows) for number, rows in self.entries[title].items()
                if number not in self.completed]

    def blocked_before(self, title):
        for sheet in self.sheets:
            if sheet.title == title:
                return None
            if self.pending_in_sheet(sheet.title):
                return sheet.title
        return None

    def mark(self, number, *, reason=None, mark_processed=True):
        # All occurrences are updated in one replacement, including later sheets.
        for title, row in self.locations[number]:
            write_row_markers(self.document[title], [row], self.marker, self.reason_prefix,
                              reason=reason, mark_processed=mark_processed)
        atomic_save(self.document, self.path)
        if mark_processed:
            self.completed.add(number)
            self.reasons[number] = reason

    def sync_completed_duplicates(self):
        changed = False
        for number in self.completed:
            for title, row in self.locations[number]:
                sheet = self.document[title]
                if str(sheet.cell(row, 3).value or '').strip() != self.marker:
                    write_row_markers(sheet, [row], self.marker, self.reason_prefix,
                                      reason=self.reasons.get(number))
                    changed = True
        if changed:
            atomic_save(self.document, self.path)

    def recover(self, journal, *, defer_uncertain=False):
        if journal.records:
            title = journal.sheet
            if title not in self.entries:
                raise RuntimeError(f'恢复记录所属工作表“{title}”已不存在，请先核对号码表。')
            pending = self.pending_in_sheet(title)
            row_numbers = {rows[0]: number for number, rows in pending}
            def mark(rows, reason=None):
                self.mark(row_numbers[rows[0]], reason=reason)
            options = {}
            if defer_uncertain:
                options['defer_uncertain'] = lambda rows, reason: self.mark(
                    row_numbers[rows[0]], reason=reason, mark_processed=False)
            journal.recover(pending, mark, all_numbers=set(self.entries[title]), **options)
        self.sync_completed_duplicates()
        deferred = {number for number, record in journal.records.items()
                    if defer_uncertain and record['state'] == 'sending'}
        return [(number, rows) for number, rows in self.pending_numbers() if number not in deferred]
