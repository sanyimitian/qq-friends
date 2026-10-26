"""Two configurable Windows apps sharing a Tk interface and worker controller."""
from pathlib import Path
import json
import os
import sys
import time

from gui_config import fields, defaults, validate, load_config, save_config, state_home, config_path

RUNNING_NOTICE = '请勿操作鼠标，等待程序执行完成。'


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
        title = 'QQ 群成员采集' if mode == 'capture' else 'QQ 好友添加'
        root.title(title)
        root.geometry('1040x820')
        root.minsize(850, 650)
        root.configure(background='#f1f5f9')
        style = ttk.Style(root)
        style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI', 10))
        style.configure('TFrame', background='#f1f5f9')
        style.configure('TLabel', background='#f1f5f9', foreground='#172b4d')
        style.configure('Hint.TLabel', foreground='#64748b', font=('Microsoft YaHei UI', 9))
        style.configure('TButton', padding=(14, 8))
        style.configure('Start.TButton', background='#2563eb', foreground='white', font=('Microsoft YaHei UI', 11, 'bold'))
        style.map('Start.TButton', background=[('active', '#1d4ed8'), ('disabled', '#94a3b8')])
        style.configure('Stop.TButton', background='#dc2626', foreground='white')
        style.map('Stop.TButton', background=[('active', '#b91c1c'), ('disabled', '#94a3b8')])
        header = tk.Frame(root, bg='#172b4d', padx=22, pady=15)
        header.pack(fill='x')
        tk.Label(header, text=title, font=('Microsoft YaHei UI', 21, 'bold'), bg='#172b4d', fg='white').pack(anchor='w')
        subtitle = '设置采集范围与输出位置，开始后实时查看进度' if mode == 'capture' else '选择号码表与账号，按配置执行添加流程'
        tk.Label(header, text=subtitle, font=('Microsoft YaHei UI', 10), bg='#172b4d', fg='#cbd5e1').pack(anchor='w', pady=(5, 0))
        actions = ttk.Frame(root, padding=(16, 10))
        actions.pack(fill='x')
        self.start_button = ttk.Button(actions, text='▶  开始', style='Start.TButton', command=self.start)
        self.start_button.pack(side='left')
        self.stop_button = ttk.Button(actions, text='■  停止', style='Stop.TButton', command=self.stop, state='disabled')
        self.stop_button.pack(side='left', padx=8)
        self.status = tk.StringVar(value='就绪 · 配置后点击开始')
        ttk.Label(actions, textvariable=self.status).pack(side='left', padx=12)
        self.progress = ttk.Progressbar(actions, mode='indeterminate', length=130)
        self.progress.pack(side='right')
        self.running_notice = tk.Label(root, text=RUNNING_NOTICE,
                                       font=('Microsoft YaHei UI', 12, 'bold'),
                                       background='#fff3cd', foreground='#92400e',
                                       anchor='w', padx=16, pady=12, wraplength=760)
        body = ttk.Panedwindow(root, orient='vertical')
        self.body = body
        body.pack(fill='both', expand=True, padx=16, pady=(0, 12))
        settings_frame = ttk.Frame(body)
        body.add(settings_frame, weight=3)
        toolbar = ttk.Frame(settings_frame)
        toolbar.pack(fill='x', pady=(0, 8))
        for label, callback in [('保存配置', self.save), ('恢复默认', self.reset), ('选择 QQ 窗口', self.choose_window), ('打开结果目录', self.open_output)]:
            button = ttk.Button(toolbar, text=label, command=callback)
            button.pack(side='left', padx=(0, 6))
            self.edit_widgets.append(button)
        if mode == 'add':
            button = ttk.Button(toolbar, text='选择 Excel', command=self.choose_excel)
            button.pack(side='left')
            self.edit_widgets.append(button)
        self.notebook = ttk.Notebook(settings_frame)
        self.notebook.pack(fill='both', expand=True)
        try:
            config = load_config(mode)
            config_error = ''
        except Exception as error:
            config = defaults(mode)
            config_error = f'配置读取失败，已显示默认值：{error}'
        for group in dict.fromkeys(field.group for field in fields(mode)):
            page = ttk.Frame(self.notebook)
            self.notebook.add(page, text=group)
            canvas = tk.Canvas(page, highlightthickness=0, background='#f1f5f9')
            scrollbar = ttk.Scrollbar(page, orient='vertical', command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side='right', fill='y')
            canvas.pack(side='left', fill='both', expand=True)
            inside = ttk.Frame(canvas, padding=(16, 12))
            window_id = canvas.create_window((0, 0), window=inside, anchor='nw')
            inside.columnconfigure(1, weight=1)
            inside.bind('<Configure>', lambda event, c=canvas: c.configure(scrollregion=c.bbox('all')))
            canvas.bind('<Configure>', lambda event, c=canvas, w=window_id: c.itemconfigure(w, width=event.width))
            canvas.bind('<Enter>', lambda event, c=canvas: c.bind_all('<MouseWheel>', lambda e: c.yview_scroll(int(-e.delta / 120), 'units')))
            canvas.bind('<Leave>', lambda event, c=canvas: c.unbind_all('<MouseWheel>'))
            row = 0
            for field in (f for f in fields(mode) if f.group == group):
                ttk.Label(inside, text=field.label).grid(row=row, column=0, sticky='nw', padx=(0, 18), pady=8)
                if field.kind in ('lines', 'multiline'):
                    height = min(max(len(config[field.key]), 2), 4) if field.kind == 'lines' else 3
                    content = '\n'.join(config[field.key]) if field.kind == 'lines' else config[field.key]
                    widget = tk.Text(inside, height=height, font=('Microsoft YaHei UI', 10), wrap='word', relief='solid', borderwidth=1)
                    widget.insert('1.0', content)
                    self.inputs[field.key] = widget
                elif field.kind == 'bool':
                    variable = tk.BooleanVar(value=config[field.key])
                    widget = ttk.Checkbutton(inside, text='启用', variable=variable)
                    self.inputs[field.key] = variable
                else:
                    variable = tk.StringVar(value=str(config[field.key]))
                    widget = ttk.Entry(inside, textvariable=variable)
                    self.inputs[field.key] = variable
                widget.grid(row=row, column=1, sticky='ew', pady=5)
                self.edit_widgets.append(widget)
                if field.kind in ('save_xlsx', 'open_numbers', 'exe'):
                    button = ttk.Button(inside, text='浏览…', command=lambda f=field: self.browse(f))
                    button.grid(row=row, column=2, padx=(8, 0))
                    self.edit_widgets.append(button)
                row += 1
                if field.hint:
                    ttk.Label(inside, text=field.hint, style='Hint.TLabel', wraplength=620).grid(row=row, column=1, columnspan=2, sticky='w', pady=(0, 5))
                    row += 1
        log_frame = ttk.Frame(body)
        body.add(log_frame, weight=2)
        log_bar = ttk.Frame(log_frame)
        log_bar.pack(fill='x', pady=(8, 4))
        ttk.Label(log_bar, text='运行日志', font=('Microsoft YaHei UI', 11, 'bold')).pack(side='left')
        ttk.Button(log_bar, text='打开日志目录', command=self.open_logs).pack(side='right')
        ttk.Button(log_bar, text='清空显示', command=self.clear_log).pack(side='right', padx=6)
        self.log = ScrolledText(log_frame, height=11, font=('Consolas', 10), background='#101c31', foreground='#dbeafe', insertbackground='white', relief='flat', state='disabled', wrap='word')
        self.log.pack(fill='both', expand=True)
        self.append_log('参数会自动保存；运行期间锁定配置。停止任务后可以再次开始。\n')
        if mode == 'add':
            self.append_log('“自动发送”开启会实际发送好友申请。“已处理”只表示已尝试，不代表添加成功。\n')
        if config_error:
            self.append_log(config_error + '\n')
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
        ttk.Label(content, text='单击选中后点击“使用所选窗口”；双击或回车也可确认。\n仅列出窗口，不打开资料卡或读取账号。',
                  wraplength=500).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 12))
        listing = tk.Listbox(content, height=7, exportselection=False, font=('Microsoft YaHei UI', 11))
        listing.grid(row=1, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(content, orient='vertical', command=listing.yview)
        scrollbar.grid(row=1, column=1, sticky='ns')
        listing.configure(yscrollcommand=scrollbar.set)
        listing.insert('end', '0：自动选择首个 QQ 窗口' if self.mode == 'capture' else '0：全部未排除账号')
        unique = list(dict.fromkeys(rows))
        for pid, status in unique:
            listing.insert('end', f'{pid}：QQ（{status}）')
        try:
            current_pid = int(self.inputs['TARGET_PROCESS_ID'].get())
        except (TypeError, ValueError):
            current_pid = 0
        initial = next((index + 1 for index, (pid, _) in enumerate(unique) if pid == current_pid), 0)
        listing.selection_set(initial)
        listing.activate(initial)
        listing.see(initial)
        def choose(event=None):
            selected = listing.curselection()
            if selected:
                pid = 0 if selected[0] == 0 else unique[selected[0] - 1][0]
                self.inputs['TARGET_PROCESS_ID'].set(pid)
                target = f'QQ 进程 {pid}' if pid else ('自动选择首个 QQ 窗口' if self.mode == 'capture' else '全部未排除账号')
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
            self.append_log(f"已选择 Excel 号码表：{config['NUMBER_FILE_PATH']}\n直接读取 A 列，将“已处理”写回同一文件的 B 列，无需转换。\n")
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

    def start(self):
        from tkinter import messagebox
        if self.controller.running:
            return
        try:
            config = self.collect(for_run=not self.test_mode)
            save_config(self.mode, config)
            self.set_busy(True)
            self.root.update_idletasks()
            self.controller.start(config, simulate=self.test_mode)
            self.status.set('运行中')
            self.append_log(f'\n—— 开始任务 · {time.strftime("%H:%M:%S")} ——\n日志：{self.controller.log_path}\n')
            self.append_log(RUNNING_NOTICE + '\n')
        except Exception as error:
            self.set_busy(False)
            self.status.set('启动失败')
            if self.test_mode:
                raise
            messagebox.showerror('无法开始', str(error), parent=self.root)

    def stop(self):
        if self.controller.running:
            self.controller.stop()
            self.stop_button.configure(state='disabled')
            self.status.set('正在停止…')
            self.append_log('已请求停止；后台若无响应，将在约 2 秒后强制结束。\n')

    def poll(self):
        text, result = self.controller.poll()
        if text:
            self.append_log(text)
        if self.test_mode and '模拟任务已启动' in text:
            self.smoke_seen = True
            self.root.after(200, self.stop_button.invoke)
        if result:
            self.set_busy(False)
            self.status.set({'finished': '流程已结束', 'stopped': '已停止', 'error': '运行失败'}.get(result['status'], '已结束'))
            self.append_log(result['message'] + '\n')
            if self.test_mode:
                self.smoke_stopped = result['status'] == 'stopped'
                self.smoke_report.write_text(json.dumps({'mode': self.mode, 'worker_started': self.smoke_seen, 'stopped': self.smoke_stopped, 'result': result, 'config_saved': config_path(self.mode).is_file()}, ensure_ascii=False, indent=2), encoding='utf-8')
                self.root.after(200, self.root.destroy)
                return
        if self.closing and not self.controller.running:
            self.controller.cleanup()
            self.root.destroy()
            return
        self.root.after(120, self.poll)

    def close(self):
        self.closing = True
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
