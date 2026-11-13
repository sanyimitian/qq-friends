"""Shared lightweight presentation for the two QQ desktop tools."""
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import ttkbootstrap as tb
from gui_config import fields as config_fields, defaults, load_config
from dataclasses import replace


def fields(mode):
    if mode != 'add':
        return config_fields(mode)
    presentation = {
        'NUMBER_FILE_PATH': dict(label='好友号码表', hint='Excel 按标签从左到右处理，当前表完成后才进入下一张。'),
        'TARGET_PROCESS_ID': dict(label='使用哪个 QQ', hint='选择已登录的 QQ 窗口；未指定时使用全部未排除账号。'),
        'NUMBERS_PER_ACCOUNT': dict(label='每个 QQ 处理几人', hint='每个账号本次的数量上限；Excel 剩余不足也处理，下次接着当前表。'),
        'VERIFICATION_MESSAGE': dict(hint='普通验证页面填写这段信息，不修改好友备注。'),
        'QUESTION_ANSWER': dict(hint='遇到验证问题时，每题都填写这段内容；留空则跳过。'),
        'AUTO_SEND': dict(label='自动发送申请', hint='开启后核对内容并发送；关闭时只填表，随后停止。'),
        'QQ_EXE': dict(group='高级设置'),
        'EXCLUDED_QQ_ACCOUNTS': dict(group='高级设置', label='不使用这些 QQ'),
        'NUMBER_INTERVAL_SECONDS': dict(group='高级设置', label='每个号码间隔 / 秒'),
        'AVATAR_AUTO_DETECT': dict(group='高级设置', label='自动找到我的头像'),
        'CLOSE_SECONDARY_WINDOWS': dict(group='高级设置', label='开始前关闭其他 QQ 页面'),
    }
    displayed = [replace(field, **presentation.get(field.key, {})) for field in config_fields(mode)]
    order = ['NUMBER_FILE_PATH', 'TARGET_PROCESS_ID', 'NUMBERS_PER_ACCOUNT', 'REPEAT_EVERY_24_HOURS', 'REPEAT_INTERVAL_HOURS',
             'VERIFICATION_MESSAGE', 'QUESTION_ANSWER', 'AUTO_SEND']
    return sorted(displayed, key=lambda field: order.index(field.key) if field.key in order else len(order))


BG, WHITE, TEXT = '#f3f6fb', '#ffffff', '#24324b'
MUTED, BLUE, BORDER = '#6b778b', '#2764e7', '#e3e9f2'
FONT = ('Microsoft YaHei UI', 10)


def apply_theme(root):
    if tb.Style.instance is not None and tb.Style.instance.tk is not root.tk:
        tb.Style.instance = None
    style = tb.Style(theme='bootstrap-light')
    root.option_add('*Font', FONT)
    root.configure(background=BG)
    style.configure('.', font=FONT)
    style.configure('TFrame', background=WHITE)
    style.configure('TLabel', background=WHITE, foreground=TEXT)
    style.configure('Hint.TLabel', foreground=MUTED, font=('Microsoft YaHei UI', 9))
    style.configure('TEntry', padding=(10, 7))
    style.configure('TButton', font=FONT, padding=(14, 8))
    style.configure('TNotebook', background=WHITE, borderwidth=0, tabmargins=(12, 8, 12, 0))
    style.configure('TNotebook.Tab', padding=(18, 10), font=FONT)
    style.map('TNotebook.Tab', foreground=[('selected', BLUE)],
              background=[('selected', WHITE), ('!selected', '#f7f9fc')])
    style.configure('TPanedwindow', background=BG, sashwidth=7)
    return style


def tooltip(app, widget, text):
    app.tooltips.append(tb.ToolTip(widget, text=text, delay=450, wraplength=330))


def button(app, parent, text, command, help_text, *, primary=False, editable=False):
    widget = tb.Button(parent, text=text, command=command,
                       bootstyle='primary' if primary else 'secondary-outline')
    tooltip(app, widget, help_text)
    if editable:
        app.edit_widgets.append(widget)
    return widget


def build_interface(app):
    root, mode = app.root, app.mode
    app.tooltips = []
    app.style = apply_theme(root)
    app.readonly_widgets = []
    title = 'QQ 好友添加' if mode == 'add' else 'QQ 群成员采集'
    root.title(title)
    root.geometry('980x820')
    root.minsize(820, 680)
    header = tk.Frame(root, background=BG, padx=24, pady=20)
    header.pack(fill='x')
    title_block = tk.Frame(header, background=BG)
    title_block.pack(side='left')
    tk.Label(title_block, text=title, background=BG, foreground=TEXT,
             font=('Microsoft YaHei UI', 20, 'bold')).pack(anchor='w')
    subtitle = ('选好号码表和 QQ，核对验证信息即可开始' if mode == 'add' else
                '选好 QQ 和采集范围，号码自动保存到 Excel')
    app.send_notice = tk.StringVar(value=subtitle)
    tk.Label(title_block, textvariable=app.send_notice, background=BG, foreground=MUTED,
             font=('Microsoft YaHei UI', 9)).pack(anchor='w', pady=(5, 0))
    header_buttons = tk.Frame(header, background=BG)
    header_buttons.pack(side='right')
    button(app, header_buttons, '保存设置', app.save, '保存当前设置，下次打开继续用。',
           editable=True).pack(side='left', padx=(0, 8))
    button(app, header_buttons, '恢复默认', app.reset, '恢复页面默认值，保存后生效。',
           editable=True).pack(side='left')

    footer = tk.Frame(root, background=WHITE, padx=24, pady=14,
                      highlightbackground=BORDER, highlightthickness=1)
    footer.pack(side='bottom', fill='x')
    actions = ttk.Frame(footer)
    actions.pack(fill='x')
    app.capture_hint = None
    if mode == 'capture' and hasattr(app, 'toggle_pause'):
        app.capture_hint = ttk.Label(footer, text='Ctrl+B 暂停或继续，进度会自动保存。',
                                     style='Hint.TLabel', anchor='w', justify='left')
        app.capture_hint.pack(fill='x', pady=(10, 0))
        footer.bind('<Configure>', lambda event: app.capture_hint.configure(
            wraplength=max(1, event.width - 50)))
    app.start_button = button(app, actions, '开始添加' if mode == 'add' else '开始采集',
                              app.start, '保存设置并开始，进度会显示在运行记录里。', primary=True)
    app.start_button.pack(side='left')
    app.stop_button = button(app, actions, '停止', app.stop,
                             '结束任务并取消后续循环，已保存的结果会保留。' if mode == 'add' else
                             '结束任务，已保存的结果会保留。')
    app.stop_button.configure(state='disabled')
    app.stop_button.pack(side='left', padx=(8, 0))
    app.pause_button = None
    app.pause_state = None
    if mode == 'capture' and hasattr(app, 'toggle_pause'):
        app.pause_button = button(app, actions, '暂停 Ctrl+B', app.toggle_pause,
                                  '暂停并保存进度，再点一次继续；也可按 Ctrl+B。')
        app.pause_button.configure(state='disabled')
        app.pause_button.pack(side='left', padx=(8, 0))
    app.status = tk.StringVar(value='就绪')
    app.status_label = ttk.Label(actions, textvariable=app.status, style='Hint.TLabel')
    app.status_label.configure(wraplength=420, justify='left')
    app.status_label.pack(side='left', padx=16)
    if mode == 'add':
        footer.bind('<Configure>', lambda event: app.status_label.configure(wraplength=max(180, event.width - 390)))
    app.progress = tb.Progressbar(actions, mode='indeterminate', length=90, bootstyle='primary')
    app.progress.pack(side='right')
    from desktop_app import RUNNING_NOTICE
    app.running_notice = tk.Label(root, text=RUNNING_NOTICE, background='#fff5dd',
                                  foreground='#885c15', font=FONT, anchor='w',
                                  padx=14, pady=8, wraplength=760)
    body = ttk.Panedwindow(root, orient='vertical')
    app.body = body
    body.pack(fill='both', expand=True, padx=24, pady=(0, 14))
    settings = tk.Frame(body, background=WHITE, highlightbackground=BORDER, highlightthickness=1)
    body.add(settings, weight=3)
    app.notebook = ttk.Notebook(settings)
    app.notebook.pack(fill='both', expand=True)
    try:
        config, error = load_config(mode), ''
    except Exception as exc:
        config, error = defaults(mode), f'设置读取失败，已使用默认值：{exc}'
    app.setting_canvases = []
    page_names = {'常用设置': '常用设置', '高级设置': '更多设置', '界面适配': '界面识别'}
    for group in ('常用设置', '高级设置', '界面适配'):
        page = ttk.Frame(app.notebook)
        app.notebook.add(page, text=page_names[group])
        canvas = tk.Canvas(page, background=WHITE, highlightthickness=0, height=280)
        bar = tb.Scrollbar(page, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=bar.set)
        bar.pack(side='right', fill='y', padx=(0, 5), pady=8)
        canvas.pack(side='left', fill='both', expand=True)
        inside = ttk.Frame(canvas, padding=(20, 16, 20, 12))
        window_id = canvas.create_window((0, 0), window=inside, anchor='nw')
        inside.columnconfigure(1, weight=1)
        inside.columnconfigure(0, minsize=164)
        hints = []
        inside.bind('<Configure>', lambda event, c=canvas: c.configure(scrollregion=c.bbox('all')))
        def resize(event, c=canvas, w=window_id, labels=hints):
            c.itemconfigure(w, width=event.width)
            for label in labels:
                label.configure(wraplength=max(230, event.width - 240))
        canvas.bind('<Configure>', resize)
        app.setting_canvases.append(canvas)
        row = 0
        if group != '常用设置':
            text = '一般保持默认即可；需要时再调整。' if group == '高级设置' else '只有 QQ 页面识别不准时才需要修改。'
            ttk.Label(inside, text=text, style='Hint.TLabel').grid(
                row=row, column=0, columnspan=2, sticky='w', pady=(0, 16))
            row += 1
        for field in (f for f in fields(mode) if f.group == group):
            label = ttk.Label(inside, text=field.label)
            label.grid(row=row, column=0, sticky='nw', pady=(9, 0), padx=(0, 16))
            editor = ttk.Frame(inside)
            editor.grid(row=row, column=1, sticky='ew', pady=(0, 2))
            editor.columnconfigure(0, weight=1)
            if field.kind in ('lines', 'multiline', 'optional_multiline'):
                content = chr(10).join(config[field.key]) if field.kind == 'lines' else config[field.key]
                height = 2
                widget = tk.Text(editor, height=height, width=28, font=FONT, wrap='word',
                                 background='#fafcff', foreground=TEXT, relief='flat',
                                 highlightthickness=1, highlightbackground=BORDER,
                                 highlightcolor=BLUE, padx=10, pady=8, undo=True)
                widget.insert('1.0', content)
                app.inputs[field.key] = widget
            else:
                variable = tk.BooleanVar(value=config[field.key]) if field.kind == 'bool' else tk.StringVar(value=str(config[field.key]))
                app.inputs[field.key] = variable
                if field.kind == 'bool':
                    widget = tb.Checkbutton(editor, text='开启', variable=variable, bootstyle='primary-round-toggle')
                    if field.key == 'AUTO_SEND':
                        def refresh_send(*_, value=variable):
                            app.send_notice.set('发送方式：核对内容后自动发送申请' if value.get() else
                                                '发送方式：只填写申请，随后停止，不发送')
                        variable.trace_add('write', refresh_send)
                        refresh_send()
                elif field.key == 'TARGET_PROCESS_ID':
                    display = tk.StringVar()
                    def refresh_account(*_, value=variable, display=display):
                        value_text = value.get().strip()
                        display.set('全部可用 QQ' if value_text in ('', '0') else f'已选 QQ 窗口 · 编号 {value_text}')
                    variable.trace_add('write', refresh_account)
                    refresh_account()
                    widget = ttk.Entry(editor, textvariable=display, state='readonly')
                    app.readonly_widgets.append(widget)
                    app.account_display = display
                else:
                    widget = ttk.Entry(editor, textvariable=variable, width=28)
            widget.grid(row=0, column=0, sticky='ew', ipady=3 if field.kind == 'bool' else 0)
            app.edit_widgets.append(widget)
            if field.kind in ('save_xlsx', 'open_numbers', 'exe'):
                caption = '保存位置' if field.kind == 'save_xlsx' else '选择文件'
                action = button(app, editor, caption, lambda f=field: app.browse(f),
                                '选择文件保存位置。' if field.kind == 'save_xlsx' else '选择要使用的文件。', editable=True)
                action.grid(row=0, column=1, padx=(10, 0))
            elif field.key in ('TARGET_PROCESS_ID', 'TARGET_QQ_NUMBER'):
                action = button(app, editor, '选择 QQ', app.choose_window, '从已登录的 QQ 窗口中选择。', editable=True)
                action.grid(row=0, column=1, padx=(10, 0))
            row += 1
            if field.hint:
                tooltip(app, label, field.hint)
            compact_hint = mode == 'add' and field.key in {'TARGET_PROCESS_ID', 'NUMBERS_PER_ACCOUNT'}
            if field.hint and not compact_hint:
                hint = ttk.Label(inside, text=field.hint, style='Hint.TLabel', wraplength=570)
                hint.grid(row=row, column=1, sticky='w', pady=(0, 9))
                hints.append(hint)
                row += 1
            else:
                editor.grid_configure(pady=(0, 14))
    def wheel(event):
        hovered = root.winfo_containing(event.x_root, event.y_root)
        while hovered is not None:
            if isinstance(hovered, tk.Text):
                return None
            if hovered in app.setting_canvases:
                top, bottom = hovered.yview()
                if bottom - top < .999:
                    hovered.yview_scroll(-int(event.delta / 120), 'units')
                return 'break'
            hovered = getattr(hovered, 'master', None)
        return None
    root.bind('<MouseWheel>', wheel, add='+')
    log_card = tk.Frame(body, background=WHITE, highlightbackground=BORDER, highlightthickness=1)
    body.add(log_card, weight=2)
    log_bar = ttk.Frame(log_card, padding=(14, 9))
    log_bar.pack(fill='x')
    ttk.Label(log_bar, text='运行记录', font=('Microsoft YaHei UI', 10, 'bold')).pack(side='left')
    for caption, callback, explanation in (
        ('查看文件', app.open_output, '打开号码表所在文件夹。'),
        ('日志目录', app.open_logs, '打开保存运行日志的文件夹。'),
        ('清空记录', app.clear_log, '仅清空这里的文字，不删除日志文件。'),
    ):
        action = tb.Button(log_bar, text=caption, command=callback, bootstyle='secondary-link')
        action.pack(side='right', padx=(10, 0))
        tooltip(app, action, explanation)
    app.log = ScrolledText(log_card, height=13, font=('Microsoft YaHei UI', 9),
                           background='#f8fafc', foreground='#526079', relief='flat',
                           borderwidth=0, padx=14, pady=8, state='disabled', wrap='word')
    app.log.pack(fill='both', expand=True, padx=1, pady=(0, 1))
    app.append_log('点击开始会保存设置。运行进度显示在这里。' + chr(10))
    if mode == 'add':
        app.append_log('请核对发送方式；已处理不代表对方已通过。' + chr(10))
    if error:
        app.append_log(error + chr(10))
