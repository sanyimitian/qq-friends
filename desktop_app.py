"""Two configurable Windows apps sharing a Tk interface and worker controller."""
from pathlib import Path
from copy import deepcopy
import math
import json
import os
import re
import sys
import time

from gui_config import fields, defaults, validate, load_config, save_config, state_home, config_path

RUNNING_NOTICE = '正在操作 QQ，请暂时不要使用鼠标和键盘。'
REPEAT_INTERVAL_SECONDS = 24 * 60 * 60


class App:
    def __init__(self, root, mode):
        import tkinter as tk
        from tkinter import ttk
        from tkinter.scrolledtext import ScrolledText
        from gui_controller import TaskController
        self.root, self.mode = root, mode
        self.controller = TaskController(mode)
        self.inputs = {}
        self.edit_widgets = []
        self.closing = False
        self.test_mode = False
        self.smoke_report = None
        self.smoke_stopped = False
        self.smoke_seen = False
        self.repeat_config = None
        self.next_run_at = None
        self.repeat_round = 0
        self.wait_status = None
        from desktop_ui import build_interface
        build_interface(self)
        root.protocol('WM_DELETE_WINDOW', self.close)
        root.bind('<Escape>', lambda event: self.stop())
        root.after(120, self.poll)

    def collect(self, for_run=False):
        import tkinter as tk
        raw = {key: value.get('1.0', 'end-1c') if isinstance(value, tk.Text) else value.get() for key, value in self.inputs.items()}
        return validate(self.mode, raw, for_run=for_run)

    def append_log(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text)
        if int(self.log.index('end-1c').split('.')[0]) > 5000:
            self.log.delete('1.0', '1000.0')
        self.log.see('end')
        self.log.configure(state='disabled')

    def clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    def save(self):
        from tkinter import messagebox
        try:
            path = save_config(self.mode, self.collect())
            self.status.set('配置已保存')
            self.append_log(f'配置已保存：{path}\n')
        except Exception as error:
            messagebox.showerror('配置有误', str(error), parent=self.root)

    def reset(self):
        import tkinter as tk
        for key, value in defaults(self.mode).items():
            target = self.inputs[key]
            if isinstance(target, tk.Text):
                target.delete('1.0', 'end')
                target.insert('1.0', '\n'.join(value) if isinstance(value, list) else value)
            else:
                target.set(value)
        self.status.set('已恢复默认，点击保存或开始后生效')

    def browse(self, field):
        from tkinter import filedialog
        current = Path(self.inputs[field.key].get())
        options = {'parent': self.root, 'initialdir': str(current.parent if current.parent.exists() else Path.home())}
        if field.kind == 'save_xlsx':
            path = filedialog.asksaveasfilename(**options, defaultextension='.xlsx', filetypes=[('Excel 号码表', '*.xlsx')], initialfile=current.name)
        elif field.kind == 'open_numbers':
            path = filedialog.askopenfilename(**options, title='选择号码表', filetypes=[('Excel 号码表', '*.xlsx'), ('ODS 号码表（兼容旧版）', '*.ods')])
        else:
            path = filedialog.askopenfilename(**options, filetypes=[('Windows 程序', '*.exe')])
        if path:
            self.inputs[field.key].set(path)

    def choose_window(self):
        from tkinter import ttk, messagebox
        import tkinter as tk
        import win32gui
        import win32process
        rows = []
        title = self.inputs['WINDOW_TITLE'].get().strip()
        def callback(handle, unused):
            try:
                if win32gui.GetWindowText(handle).strip() == title and (win32gui.IsWindowVisible(handle) or win32gui.IsIconic(handle)):
                    pid = win32process.GetWindowThreadProcessId(handle)[1]
                    rows.append((pid, '最小化' if win32gui.IsIconic(handle) else '可见'))
            except Exception:
                # A QQ window may close while EnumWindows is visiting it.
                pass
        win32gui.EnumWindows(callback, None)
        if not rows:
            messagebox.showinfo('没有找到 QQ', '请先启动并登录 QQ，再选择窗口。', parent=self.root)
            return
        dialog = tk.Toplevel(self.root)
        dialog.title('选择 QQ 窗口')
        dialog.transient(self.root)
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)
        content = ttk.Frame(dialog, padding=12)
        content.grid(row=0, column=0, sticky='nsew')
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)
        description = (
            '选择窗口后会读取该窗口登录的 QQ 号并填入目标 QQ 号。'
            if self.mode == 'capture' else
            '单击选中后点击“使用所选窗口”；双击或回车也可确认。仅列出窗口，不读取账号。'
        )
        ttk.Label(content, text='单击选中后点击“使用所选窗口”；双击或回车也可确认。\n' + description,
                  wraplength=500).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 12))
        listing = tk.Listbox(content, height=7, exportselection=False, font=('Microsoft YaHei UI', 11))
        listing.grid(row=1, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(content, orient='vertical', command=listing.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        listing.configure(yscrollcommand=scrollbar.set)
        if self.mode == 'add':
            listing.insert('end', '0：全部未排除账号')
        unique = list(dict.fromkeys(rows))
        for pid, status in unique:
            listing.insert('end', f'{pid}：QQ（{status}）')
        if self.mode == 'add':
            try:
                current_pid = int(self.inputs['TARGET_PROCESS_ID'].get())
            except (TypeError, ValueError):
                current_pid = 0
            initial = next((index + 1 for index, (pid, _) in enumerate(unique) if pid == current_pid), 0)
        else:
            initial = 0
        listing.selection_set(initial)
        listing.activate(initial)
        listing.see(initial)
        def choose(event=None):
            selected = listing.curselection()
            if selected:
                if self.mode == 'capture':
                    pid = unique[selected[0]][0]
                    try:
                        import pyautogui
                        import win32con
                        from pywinauto import Desktop
                        import zq_qq_001 as capture

                        capture.WINDOW_TITLE = title
                        desktop = Desktop(backend='uia')
                        candidates = [window for window in desktop.windows()
                                      if window.process_id() == pid
                                      and window.window_text().strip() == title]
                        if len(candidates) != 1:
                            raise RuntimeError('未找到唯一的所选 QQ 主窗口。')
                        account = capture.read_qq_account_from_avatar(
                            candidates[0], pyautogui, win32con, win32gui, desktop
                        )
                        if not account or not re.fullmatch(r'[0-9]{5,12}', account):
                            raise RuntimeError('无法从所选窗口确认有效的 QQ 号。')
                    except Exception as error:
                        messagebox.showerror('读取 QQ 号失败', str(error), parent=dialog)
                        return 'break'
                    self.inputs['TARGET_QQ_NUMBER'].set(account)
                    target = f'目标 QQ 号 {account}'
                else:
                    pid = 0 if selected[0] == 0 else unique[selected[0] - 1][0]
                    self.inputs['TARGET_PROCESS_ID'].set(pid)
                    target = f'QQ 进程 {pid}' if pid else '全部未排除账号'
                self.status.set(f'已选择：{target}')
                self.append_log(f'已选择：{target}；点击“开始”后使用。\n')
                dialog.destroy()
            return 'break'
        def choose_double_click(event):
            index = listing.nearest(event.y)
            bounds = listing.bbox(index) if index >= 0 else None
            if bounds and bounds[1] <= event.y < bounds[1] + bounds[3]:
                listing.selection_clear(0, 'end')
                listing.selection_set(index)
                return choose()
            return 'break'
        buttons = ttk.Frame(content)
        buttons.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(12, 0))
        ttk.Button(buttons, text='使用所选窗口', command=choose).pack(side='right')
        ttk.Button(buttons, text='取消', command=dialog.destroy).pack(side='right', padx=(0, 8))
        listing.bind('<Double-Button-1>', choose_double_click)
        dialog.bind('<Return>', choose)
        dialog.bind('<KP_Enter>', choose)
        dialog.bind('<Escape>', lambda event: dialog.destroy())
        dialog.update_idletasks()
        # Size from the actual font/DPI metrics so the confirmation row stays visible.
        width = max(520, dialog.winfo_reqwidth())
        height = max(340, dialog.winfo_reqheight())
        dialog.minsize(width, height)
        dialog.geometry(f'{width}x{height}')
        dialog.grab_set()
        listing.focus_set()

    def open_output(self):
        from tkinter import messagebox
        try:
            key = 'NUMBER_WORKBOOK_PATH' if self.mode == 'capture' else 'NUMBER_FILE_PATH'
            folder = Path(self.inputs[key].get()).expanduser().resolve().parent
            folder.mkdir(parents=True, exist_ok=True)
            os.startfile(str(folder))
        except Exception as error:
            messagebox.showerror('打开失败', str(error), parent=self.root)

    def open_logs(self):
        directory = state_home() / 'runs'
        directory.mkdir(exist_ok=True)
        os.startfile(str(directory))

    def choose_excel(self):
        from tkinter import filedialog, messagebox
        current = Path(self.inputs['NUMBER_FILE_PATH'].get())
        source = filedialog.askopenfilename(parent=self.root, title='选择采集的 Excel 号码表（直接使用）',
                                          initialdir=str(current.parent if current.parent.exists() else Path.home()),
                                          filetypes=[('Excel 号码表', '*.xlsx')])
        if not source:
            return
        try:
            self.inputs['NUMBER_FILE_PATH'].set(source)
            config = self.collect()
            save_config(self.mode, config)
            self.append_log(f"已选择 Excel 号码表：{config['NUMBER_FILE_PATH']}\n直接读取 A 列，将“已处理”写回同一文件的 C 列，B 列忽略，无需转换。\n")
            self.status.set('已选择 Excel，配置后点击开始')
        except Exception as error:
            messagebox.showerror('选择失败', str(error), parent=self.root)

    def set_busy(self, busy):
        for widget in self.edit_widgets:
            widget.configure(state='disabled' if busy else 'normal')
        self.start_button.configure(state='disabled' if busy else 'normal')
        self.stop_button.configure(state='normal' if busy else 'disabled')
        if busy:
            self.running_notice.pack(fill='x', padx=16, pady=(0, 8), before=self.body)
            self.progress.start(15)
        else:
            self.running_notice.pack_forget()
            self.progress.stop()
            for widget in self.readonly_widgets:
                widget.configure(state='readonly')

    def start(self):
        from tkinter import messagebox
        if self.closing or self.controller.running or self.repeat_config is not None:
            return
        try:
            config = self.collect(for_run=not self.test_mode)
            save_config(self.mode, config)
            self.repeat_config = deepcopy(config)
            self.repeat_round = 0
            self.start_round()
        except Exception as error:
            self.append_log(f'启动失败：{error}\n')
            if self.schedule_next_round({'status': 'error'}):
                return
            self.cancel_repeat()
            self.set_busy(False)
            self.status.set('启动失败')
            if self.test_mode:
                raise
            messagebox.showerror('无法开始', str(error), parent=self.root)

    def cancel_repeat(self):
        self.repeat_config = None
        self.next_run_at = None
        self.wait_status = None

    def start_round(self):
        self.next_run_at = None
        self.wait_status = None
        self.set_busy(True)
        self.root.update_idletasks()
        self.repeat_round += 1
        self.controller.start(deepcopy(self.repeat_config), simulate=self.test_mode)
        repeating = self.mode == 'add' and self.repeat_config.get('REPEAT_EVERY_24_HOURS', False)
        label = f'第 {self.repeat_round} 轮运行中' if repeating else '运行中'
        self.status.set(label)
        self.append_log(f'\n—— 开始任务 · {time.strftime("%Y-%m-%d %H:%M:%S")} · {label} ——\n日志：{self.controller.log_path}\n')
        self.append_log(RUNNING_NOTICE + '\n')

    def schedule_next_round(self, result):
        config = self.repeat_config
        if (self.closing or self.mode != 'add' or not config
                or not config.get('REPEAT_EVERY_24_HOURS', False)
                or result.get('status') not in ('finished', 'error')):
            return False
        interval_hours = config.get('REPEAT_INTERVAL_HOURS', REPEAT_INTERVAL_SECONDS / 3600)
        self.next_run_at = time.time() + interval_hours * 3600
        # The worker has exited and released the QQ mutex; only the UI timer remains.
        self.running_notice.pack_forget()
        self.progress.stop()
        self.update_repeat_wait()
        due = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.next_run_at))
        outcome = '异常结束' if result.get('status') == 'error' or result.get('exit_code', 0) != 0 else '已结束'
        self.append_log(f'第 {self.repeat_round} 轮{outcome}，等待 {interval_hours:g} 小时；下次开始：{due}。\n'
                        '等待期间可正常使用电脑；请保持程序开启、电脑唤醒和 QQ 登录，点击“停止”可取消循环。\n')
        return True

    def update_repeat_wait(self):
        if self.next_run_at is None or self.closing:
            return
        remaining = max(0, math.ceil(self.next_run_at - time.time()))
        if remaining == 0:
            try:
                self.start_round()
            except Exception as error:
                self.append_log(f'第 {self.repeat_round} 轮启动失败：{error}\n')
                if not self.schedule_next_round({'status': 'error'}):
                    self.cancel_repeat()
                    self.set_busy(False)
                    self.status.set('启动失败，循环已停止')
            return
        hours, tail = divmod(remaining, 3600)
        minutes, seconds = divmod(tail, 60)
        due = time.strftime('%m-%d %H:%M:%S', time.localtime(self.next_run_at))
        label = f'等待第 {self.repeat_round + 1} 轮 · {hours:02}:{minutes:02}:{seconds:02}\n下次开始：{due}'
        if label != self.wait_status:
            self.status.set(label)
            self.wait_status = label

    def stop(self):
        waiting = self.next_run_at is not None
        self.cancel_repeat()
        if self.controller.running:
            self.controller.stop()
            self.stop_button.configure(state='disabled')
            self.status.set('正在停止…')
            self.append_log('已请求停止；后台若无响应，将在约 2 秒后强制结束。\n')
        elif waiting:
            self.set_busy(False)
            self.status.set('已停止，循环已取消')
            self.append_log('已取消等待和后续循环；再次点击“开始添加”才会重新运行。\n')

    def poll(self):
        text, result = self.controller.poll()
        if text:
            self.append_log(text)
        if self.test_mode and '模拟任务已启动' in text:
            self.smoke_seen = True
            self.root.after(200, self.stop_button.invoke)
        if result:
            self.append_log(result['message'] + '\n')
            if not self.schedule_next_round(result):
                self.cancel_repeat()
                self.set_busy(False)
                self.status.set({'finished': '流程已结束', 'stopped': '已停止', 'error': '运行失败'}.get(result['status'], '已结束'))
            if self.test_mode:
                self.smoke_stopped = result['status'] == 'stopped'
                self.smoke_report.write_text(json.dumps({'mode': self.mode, 'worker_started': self.smoke_seen, 'stopped': self.smoke_stopped, 'result': result, 'config_saved': config_path(self.mode).is_file()}, ensure_ascii=False, indent=2), encoding='utf-8')
                self.root.after(200, self.root.destroy)
                return
        self.update_repeat_wait()
        if self.closing and not self.controller.running:
            self.controller.cleanup()
            self.root.destroy()
            return
        self.root.after(120, self.poll)

    def close(self):
        self.closing = True
        self.cancel_repeat()
        if self.controller.running:
            self.stop()
            self.status.set('停止任务后关闭窗口…')
        else:
            self.controller.cleanup()
            self.root.destroy()


def launch(mode):
    if '--package-check' in sys.argv:
        from package_check import run
        return run(sys.argv[sys.argv.index('--package-check') + 1])
    if '--worker' in sys.argv:
        from gui_worker import run_worker
        i = sys.argv.index('--worker')
        return run_worker(mode, *sys.argv[i + 1:i + 6], simulate='--simulate' in sys.argv)
    import tkinter as tk
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    app = App(root, mode)
    if '--smoke-test' in sys.argv:
        app.test_mode = True
        app.smoke_report = Path(sys.argv[sys.argv.index('--smoke-test') + 1]).resolve()
        app.smoke_report.parent.mkdir(parents=True, exist_ok=True)
        root.after(300, app.start_button.invoke)
        def timeout():
            app.controller.cleanup()
            app.smoke_report.write_text(json.dumps({'error': 'UI smoke test timed out', 'mode': mode}), encoding='utf-8')
            root.destroy()
        root.after(25000, timeout)
    root.mainloop()
    app.controller.cleanup()
    return 0 if not app.test_mode or (app.smoke_seen and app.smoke_stopped) else 1
