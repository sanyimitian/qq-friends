import unittest
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from openpyxl import Workbook

import zq_qq_001 as script


class Rectangle:
    def __init__(self, left, top, right, bottom):
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom

    def width(self):
        return self.right - self.left

    def height(self):
        return self.bottom - self.top


class Control:
    def __init__(self, name="", kind="Group", children=(), rect=(0, 0, 100, 40)):
        self.name = name
        self.element_info = SimpleNamespace(control_type=kind)
        self._children = list(children)
        self._parent = None
        for child in self._children:
            child._parent = self
        self._rect = Rectangle(*rect)

    def window_text(self):
        return self.name

    def descendants(self):
        result = []
        for child in self._children:
            result.append(child)
            result.extend(child.descendants())
        return result

    def children(self):
        return list(self._children)

    def is_visible(self):
        return True

    def rectangle(self):
        return self._rect


class DesktopWindow(Control):
    def __init__(self, name, kind, children=(), handle=1, process_id=100):
        super().__init__(name, kind, children)
        self.handle = handle
        self._process_id = process_id

    def process_id(self):
        return self._process_id


def profile_window(group_name="测试群", group_number="123456789", member_count="(460人)"):
    document = Control("资料页", "Document", [
        Control(group_name, "Text", rect=(100, 100, 160, 122)),
        Control(group_number, "Text", rect=(100, 129, 164, 145)),
        Control(member_count, "Text", rect=(163, 129, 210, 145)),
        Control("群介绍", "Text", rect=(100, 180, 140, 198)),
    ])
    return Control("QQ", "Window", [document])


class GroupNumberDetectionTests(unittest.TestCase):
    def test_reads_number_aligned_with_group_name_and_member_count(self):
        self.assertEqual(
            script.find_group_number_in_profile(profile_window(), "测试群"),
            "123456789",
        )

    def test_member_count_alone_is_not_a_group_number(self):
        window = profile_window(group_number="", member_count="(460人)")
        self.assertIsNone(script.find_group_number_in_profile(window, "测试群"))

    def test_unrelated_chat_number_is_not_mistaken_for_group_number(self):
        document = Control("资料页", "Document", [
            Control("测试群", "Text", rect=(100, 100, 160, 122)),
            Control("123456789", "Text", rect=(250, 200, 314, 216)),
            Control("(460人)", "Text", rect=(315, 200, 362, 216)),
            Control("群介绍", "Text", rect=(100, 250, 140, 268)),
        ])
        window = Control("QQ", "Window", [document])
        self.assertIsNone(script.find_group_number_in_profile(window, "测试群"))

    def test_chat_text_without_group_profile_marker_is_ignored(self):
        document = Control("聊天", "Document", [
            Control("测试群", "Text", rect=(100, 100, 160, 122)),
            Control("123456789", "Text", rect=(100, 129, 164, 145)),
            Control("(460人)", "Text", rect=(163, 129, 210, 145)),
        ])
        window = Control("QQ", "Window", [document])
        self.assertIsNone(script.find_group_number_in_profile(window, "测试群"))

    def test_distinct_group_number_candidates_fail_closed(self):
        first = profile_window(group_number="123456789").descendants()[0]
        second = profile_window(group_number="987654321").descendants()[0]
        window = Control("QQ", "Window", [first, second])
        with self.assertRaises(RuntimeError):
            script.find_group_number_in_profile(window, "测试群")

    def test_reads_group_number_when_header_text_is_nested_or_duplicated(self):
        cases = (
            ([], "header"),
            ([Control("测试群", "Text"), Control("测试群", "Text")], "header"),
            ([Control("测试群", "Text")], "name"),
        )
        for children, expected_target in cases:
            with self.subTest(expected_target=expected_target, child_count=len(children)):
                header = Control("测试群 (460)", "Button", children)
                header.click_input = Mock()
                for child in children:
                    child.click_input = Mock()
                main = Control("QQ", "Window")
                clock = SimpleNamespace(monotonic=lambda: 0.0, sleep=Mock())
                with patch.object(script, "find_group_number_in_profile",
                                  side_effect=[None, "123456789"]), \
                        patch.object(script, "get_group_header_button", return_value=header), \
                        patch.object(script, "time", clock):
                    self.assertEqual(
                        script.read_group_number_from_profile(main, "测试群"),
                        "123456789",
                    )
                if expected_target == "header":
                    header.click_input.assert_called_once_with()
                    for child in children:
                        child.click_input.assert_not_called()
                else:
                    children[0].click_input.assert_called_once_with()
                    header.click_input.assert_not_called()

    def test_reads_unique_number_from_new_top_level_profile_card(self):
        main = DesktopWindow("QQ", "Window")
        card = DesktopWindow(
            "资料卡", "Window", [Control("2152042795", "Text")], handle=2
        )
        header = Control("测试群 (460)", "Button")
        visible = {"card": False}
        header.click_input = Mock(side_effect=lambda: visible.__setitem__("card", True))
        desktop = SimpleNamespace(
            windows=lambda: [main, *([card] if visible["card"] else [])]
        )
        clock = SimpleNamespace(monotonic=lambda: 0.0, sleep=Mock())
        with patch.object(script, "find_group_number_in_profile", return_value=None), \
                patch.object(script, "get_group_header_button", return_value=header), \
                patch.object(script, "time", clock):
            self.assertEqual(
                script.read_group_number_from_profile(main, "测试群", desktop),
                "2152042795",
            )

    def test_does_not_reuse_an_unchanged_existing_profile_card(self):
        main = DesktopWindow("QQ", "Window")
        card = DesktopWindow(
            "资料卡", "Window", [Control("3805745874", "Text")], handle=2
        )
        desktop = SimpleNamespace(windows=lambda: [main, card])
        previous = script.profile_card_number_snapshot(main, desktop)
        self.assertEqual(
            script.find_new_group_number_in_profile_card(main, desktop, previous),
            None,
        )


class JoinedGroupOrderingTests(unittest.TestCase):
    def test_visible_joined_group_rows_are_returned_top_to_bottom(self):
        category = Control("我加入的群聊 3", "Button", rect=(10, 300, 250, 330))
        pane = Control("联系人列表", "Pane", [
            Control("第三个群", "Button", rect=(20, 210, 250, 255)),
            Control("我创建的群聊", "Button", rect=(10, 270, 250, 300)),
            Control("第一个群", "Button", rect=(20, 100, 250, 145)),
            Control("我加入的群聊 3", "Button", rect=(10, 300, 250, 330)),
            Control("第二个群", "Button", rect=(20, 155, 250, 200)),
        ], rect=(0, 0, 300, 800))
        window = Control("QQ", "Window", [pane])

        with patch.object(script, "get_contact_group_list", return_value=pane), \
                patch.object(script, "get_joined_groups_category", return_value=category):
            rows = script.collect_joined_group_rows(window, (0, 0, 1200, 800))

        self.assertEqual([name for name, _ in rows], ["第一个群", "第二个群", "第三个群"])


class WorkbookSectionTests(unittest.TestCase):
    def test_group_block_does_not_reuse_rows_with_other_existing_columns(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.cell(row=1, column=3, value="preserved")
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("unused.xlsx")), \
                patch.object(script, "atomic_save"), \
                patch.object(script, "append_journal_record"):
            section = script.append_group_info(
                workbook, worksheet, "群甲", "123456789"
            )
        self.assertEqual(section["group_row"], 4)
        self.assertEqual(worksheet.cell(row=1, column=3).value, "preserved")

    def test_group_blocks_leave_two_blank_rows_and_dedupe_globally(self):
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("unused.xlsx")), \
                patch.object(script, "atomic_save") as save, \
                patch.object(script, "append_journal_record") as journal, \
                patch.object(script, "load_journal_records", return_value=[]):
            workbook, worksheet, existing, existing_groups, records = script.open_number_workbook(
                Workbook, lambda _: None
            )
            self.assertEqual(existing_groups, set())
            self.assertEqual(records, [])
            first = script.append_group_info(workbook, worksheet, "群甲", "123456789")
            self.assertEqual(first["group_row"], 1)
            self.assertTrue(script.append_number(
                workbook, worksheet, existing, "111111111", first
            ))

            second = script.append_group_info(workbook, worksheet, "群乙", "987654321")
            self.assertEqual(second["group_row"], 5)
            self.assertIsNone(worksheet.cell(row=3, column=1).value)
            self.assertIsNone(worksheet.cell(row=4, column=1).value)
            self.assertFalse(script.append_number(
                workbook, worksheet, existing, "111111111", second
            ))
            self.assertTrue(script.append_number(
                workbook, worksheet, existing, "222222222", second
            ))

            self.assertEqual(worksheet.cell(row=1, column=1).value, "群甲")
            self.assertEqual(worksheet.cell(row=1, column=2).value, "123456789")
            self.assertEqual(worksheet.cell(row=2, column=1).value, "111111111")
            self.assertEqual(worksheet.cell(row=5, column=1).value, "群乙")
            self.assertEqual(worksheet.cell(row=5, column=2).value, "987654321")
            self.assertEqual(worksheet.cell(row=6, column=1).value, "222222222")
            self.assertEqual(existing, {"111111111", "222222222"})
            self.assertEqual(save.call_count, 3)
            self.assertEqual(journal.call_count, 4)

    def test_existing_group_metadata_is_not_added_to_member_dedupe_set(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.cell(row=1, column=1, value="群甲")
        worksheet.cell(row=1, column=2, value="123456789")
        worksheet.cell(row=2, column=1, value="987654321")
        worksheet.cell(row=2, column=3, value="preserved extra data")
        worksheet.cell(row=3, column=1, value="555555555")
        worksheet.cell(row=3, column=2, value="备注列不是群号")
        worksheet.cell(row=5, column=1, value="群乙")
        worksheet.cell(row=5, column=2, value=987654321)
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("existing.xlsx")), \
                patch.object(Path, "exists", return_value=True), \
                patch.object(script, "load_journal_records", return_value=[]):
            _, _, existing, existing_groups, records = script.open_number_workbook(
                Workbook, lambda _: workbook
            )
        self.assertEqual(existing, {"987654321"})
        self.assertEqual(existing_groups, {"123456789", "987654321"})
        self.assertEqual(records, [])

    def test_missing_workbook_starts_with_no_existing_groups(self):
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("missing.xlsx")), \
                patch.object(Path, "exists", return_value=False), \
                patch.object(script, "atomic_save"), \
                patch.object(script, "load_journal_records", return_value=[]):
            _, worksheet, existing_numbers, existing_groups, records = script.open_number_workbook(
                Workbook, lambda _: self.fail("missing workbook should not be loaded")
            )
        self.assertEqual(worksheet.title, "QQ号码")
        self.assertEqual(existing_numbers, set())
        self.assertEqual(existing_groups, set())
        self.assertEqual(records, [])

    def test_incomplete_journal_group_is_not_marked_as_skipped(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.cell(row=1, column=1, value="群甲")
        worksheet.cell(row=1, column=2, value="123456789")
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("existing.xlsx")), \
                patch.object(Path, "exists", return_value=True), \
                patch.object(script, "load_journal_records", return_value=[
                    {"type": "group", "group_name": "群甲", "group_number": "123456789"},
                ]):
            _, _, _, existing_groups, _ = script.open_number_workbook(
                Workbook, lambda _: workbook
            )
        self.assertEqual(existing_groups, set())

    def test_group_info_reuses_empty_existing_block(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.cell(row=1, column=1, value="群甲")
        worksheet.cell(row=1, column=2, value="123456789")
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("existing.xlsx")), \
                patch.object(script, "atomic_save") as save, \
                patch.object(script, "append_journal_record") as journal:
            section = script.append_group_info(
                workbook, worksheet, "群甲", "123456789"
            )
        self.assertEqual(section["group_row"], 1)
        self.assertEqual(section["next_member_row"], 2)
        save.assert_not_called()
        journal.assert_not_called()

    def test_journal_reconciliation_writes_members_in_one_workbook_save(self):
        workbook = Workbook()
        worksheet = workbook.active
        with patch.object(script, "NUMBER_WORKBOOK_PATH", Path("recovered.xlsx")), \
                patch.object(script, "atomic_save") as save:
            changed = script.reconcile_journal(
                workbook,
                worksheet,
                [
                    {"type": "group", "group_name": "群甲", "group_number": "123456789"},
                    {"type": "member", "group_number": "123456789", "number": "111111111"},
                    {"type": "member", "group_number": "123456789", "number": "222222222"},
                ],
            )
        self.assertTrue(changed)
        self.assertEqual(save.call_count, 1)
        self.assertEqual(worksheet.cell(row=1, column=1).value, "群甲")
        self.assertEqual(worksheet.cell(row=1, column=2).value, "123456789")
        self.assertEqual(worksheet.cell(row=2, column=1).value, "111111111")
        self.assertEqual(worksheet.cell(row=3, column=1).value, "222222222")

    def test_journal_fsync_uses_configured_batch_and_flushes_remainder(self):
        with tempfile.TemporaryDirectory() as directory:
            journal_path = Path(directory) / "numbers.jsonl"
            with patch.object(script, "NUMBER_JOURNAL_PATH", journal_path), \
                    patch.object(script, "JOURNAL_FSYNC_BATCH_SIZE", 2), \
                    patch.object(script.os, "fsync") as fsync:
                script._journal_pending_sync_count = 0
                script.append_journal_record({
                    "type": "member", "group_number": "123456789", "number": "111111111"
                })
                self.assertEqual(fsync.call_count, 0)
                script.append_journal_record({
                    "type": "member", "group_number": "123456789", "number": "222222222"
                })
                self.assertEqual(fsync.call_count, 1)
                script.append_journal_record({
                    "type": "member", "group_number": "123456789", "number": "333333333"
                })
                self.assertEqual(fsync.call_count, 1)
                script.sync_journal()
                self.assertEqual(fsync.call_count, 2)
                self.assertEqual(len(journal_path.read_text(encoding="utf-8").splitlines()), 3)


if __name__ == "__main__":
    unittest.main()
