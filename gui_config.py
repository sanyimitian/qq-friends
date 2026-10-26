"""Typed page definitions and persistent settings for the desktop apps."""
from dataclasses import dataclass
from pathlib import Path
import json
import math
import os
import re
import sys


DEFAULT_VERIFICATION_MESSAGE = '您好，0.18稳定不降智中转免费送额度，欢迎来试，地址看头像'


@dataclass(frozen=True)
class Field:
    key: str
    label: str
    kind: str
    default: object
    group: str = '常用设置'
    hint: str = ''
    minimum: float = 0


def app_home():
    return Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent


def state_home():
    override = os.environ.get('QQ_ASSISTANT_STATE_DIR')
    root = Path(override) if override else Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'QQDesktopAssistant'
    root.mkdir(parents=True, exist_ok=True)
    return root


def data_home():
    return state_home() / 'data' if getattr(sys, 'frozen', False) else app_home()


def fields(mode):
    data = data_home()
    common = [Field('TARGET_PROCESS_ID', 'QQ 窗口进程 ID', 'int', 0, hint='0 为自动选择；抓取用首个窗口，加好友使用全部账号。指定一个窗口且未排除账号时，加好友无需读取资料卡。')]
    if mode == 'capture':
        return [
            Field('NUMBER_WORKBOOK_PATH', '号码输出文件', 'save_xlsx', str(data / 'qq_member_numbers.xlsx'), hint='自动创建 XLSX；新号码逐条保存、自动去重。'),
            *common,
            Field('MAX_CHATS', '最多检查会话数', 'int', 3, minimum=1, hint='私聊也占名额；排除的会话不占名额。'),
            Field('MAX_GROUPS', '最多处理群数', 'int', 2, minimum=1),
            Field('MAX_MEMBER_PAGE_TURNS', '每群向下翻页次数', 'int', 2, hint='0 只处理当前一页；2 表示当前页后最多再向下翻 2 次。到列表底部提前结束。'),
            Field('EXCLUDED_CHAT_NAMES', '排除群名', 'lines', ['蒂玖mod交流群', 'CodebotAPI客服2️⃣群', 'AI资源共享群', 'tokkenShop中转站'], hint='每行一个完整群名；清空表示不排除。'),
            Field('EXCLUDED_MEMBER_NAMES', '排除成员名', 'lines', ['Q群管家'], hint='每行一个名称；群主和管理员按原规则跳过。'),
            Field('USE_OCR', '启用 OCR 后备识别', 'bool', True),
            Field('WAIT_SECONDS', '界面操作等待 / 秒', 'float', 1.0, '高级设置', minimum=0.1),
            Field('MEMBER_CLICK_DELAY', '点击成员后等待 / 秒', 'float', 1.0, '高级设置', minimum=0.1),
            Field('PROFILE_READ_TIMEOUT', '资料读取超时 / 秒', 'float', 2.0, '高级设置', minimum=0.1),
            Field('PROFILE_POLL_INTERVAL', '资料检查间隔 / 秒', 'float', 0.25, '高级设置', minimum=0.05),
            Field('MAX_NUMBER_RECOGNITION_ATTEMPTS', '每人识别尝试次数', 'int', 3, '高级设置', minimum=1),
            Field('MEMBER_PAGE_SIZE', '翻页目标行数', 'int', 14, '高级设置', minimum=1, hint='按成员列表滚动接口推进，最多推进可见高度的 70%，保留重叠避免漏人。'),
            Field('MEMBER_TRIGGER_POSITION', '旧版翻页位置（已停用）', 'int', 13, '高级设置', minimum=1, hint='仅保留旧配置；现在少于 13 名可见成员也会继续翻页。'),
            Field('MAX_SCROLLS', '会话列表最大滚动次数', 'int', 500, '高级设置'),
            Field('NO_PROGRESS_LIMIT', '连续无新会话结束阈值', 'int', 3, '高级设置', minimum=1),
            Field('WINDOW_TITLE', 'QQ 主窗口标题', 'text', 'QQ', '界面适配'),
            Field('SESSION_LIST_TEXT', '会话列表名称', 'text', '会话列表', '界面适配'),
            Field('MEMBER_LIST_TEXT', '成员列表名称', 'text', '成员列表', '界面适配'),
            Field('GROUP_HEADER_PATTERN', '群标题识别表达式', 'regex', r'\(\s*\d+\s*\)\s*$', '界面适配'),
            Field('MEMBER_COUNT_PATTERN', '成员数量识别表达式', 'regex', r'群聊成员\s*(\d+)', '界面适配'),
            Field('NUMBER_TEXT_PATTERN', '号码识别表达式', 'regex', r'^\s*(\d{5,12})\s*$', '界面适配'),
        ]
    if mode != 'add':
        raise ValueError('未知程序类型')
    return [
        Field('NUMBER_FILE_PATH', '待处理号码表', 'open_numbers', str(data / 'qq_member_numbers.xlsx'), hint='直接选择采集的 Excel（.xlsx），兼容 ODS。A 列 QQ 号，无表头；B 列“已处理”自动跳过并写回原文件。'),
        Field('QQ_EXE', 'QQ 安装位置', 'exe', r'D:\01App\Tencent\QQNT\QQ.exe', hint='请先登录 QQ；程序通过已登录窗口操作。'),
        *common,
        Field('NUMBERS_PER_ACCOUNT', '每个账号本次处理数量', 'int', 3, minimum=1, hint='不足这个数量时结束；每个账号处理一个批次。'),
        Field('EXCLUDED_QQ_ACCOUNTS', '不使用的 QQ 账号', 'lines', [], hint='每行一个 QQ 号；清空表示不排除。'),
        Field('VERIFICATION_MESSAGE', '加好友验证信息', 'multiline', DEFAULT_VERIFICATION_MESSAGE, hint='填写申请表单的验证信息，不修改好友备注。关闭自动发送时，填好后停留在申请表单。'),
        Field('AUTO_SEND', '自动点击申请表单的“发送”', 'bool', True, hint='勾选会实际发送申请。关闭时，在首个申请表单处结束且不标记该号码。'),
        Field('NUMBER_INTERVAL_SECONDS', '号码之间等待 / 秒', 'float', 1.0),
        Field('AVATAR_AUTO_DETECT', '自动适配“某某的头像”按钮', 'bool', True),
        Field('CLOSE_SECONDARY_WINDOWS', '切换账号前关闭附属 QQ 窗口', 'bool', True),
        Field('STARTUP_TIMEOUT_SECONDS', 'QQ 启动等待 / 秒', 'float', 180, '高级设置', hint='供备用启动逻辑使用；正常流程要求先登录。', minimum=1),
        Field('PAGE_TIMEOUT_SECONDS', '搜索/申请页面打开超时 / 秒', 'float', 20, '高级设置', minimum=1),
        Field('SEARCH_TIMEOUT_SECONDS', '单次搜索超时 / 秒', 'float', 20, '高级设置', minimum=1),
        Field('MAX_SEARCH_ATTEMPTS', '搜索最多尝试次数', 'int', 3, '高级设置', minimum=1),
        Field('RESULT_STABLE_SECONDS', '结果稳定等待 / 秒', 'float', 1.0, '高级设置', minimum=0.1),
        Field('SEND_TIMEOUT_SECONDS', '发送后申请窗关闭等待 / 秒', 'float', 5.0, '高级设置', minimum=0.1, hint='发送后等待申请窗关闭，再继续下一号码；窗口未关闭时暂停，请核对 QQ 提示。'),
        Field('AVATAR_TIMEOUT_SECONDS', '读取登录账号超时 / 秒', 'float', 5.0, '高级设置', minimum=0.1),
        Field('PROCESSED_MARKER', '处理完成标记', 'text', '已处理', '高级设置', hint='更改后原标记可能失效。已处理不代表加好友成功。'),
        Field('WINDOW_TITLE', 'QQ 主窗口标题', 'text', 'QQ', '界面适配'),
        Field('PAGE_TITLE', '搜索窗口标题', 'text', '综合搜索', '界面适配', hint='自动兼容“综合搜索”和“全网搜索”，旧配置无需修改。'),
        Field('MORE_MENU_TEXT', '更多菜单名称', 'text', '更多', '界面适配'),
        Field('MENU_ITEM_TEXT', '加好友入口名称', 'text', '加好友/群', '界面适配'),
        Field('USER_TAB_TEXT', '用户分类名称', 'text', '用户', '界面适配'),
        Field('AVATAR_BUTTON_TEXT', '头像按钮备用名称', 'text', '昵称', '界面适配', hint='自动适配开启时也会匹配“某某的头像”。'),
        Field('SEND_BUTTON_TEXT', '发送按钮名称', 'text', '发送', '界面适配'),
        Field('RESULT_ACTIONS', '允许点击的添加按钮', 'lines', ['添加', '加入', '申请加入'], '界面适配'),
        Field('NO_RESULT_MARKERS', '无结果提示词', 'lines', ['没有找到','未找到','无搜索结果','没有搜索结果','未搜索到','暂无搜索结果','暂无相关结果','没有相关结果','没有相关用户','没有相关群聊','没有结果','暂无结果'], '界面适配'),
        Field('LOADING_MARKERS', '加载提示词', 'lines', ['搜索中','正在搜索','加载中','正在加载'], '界面适配'),
        Field('STATIC_PAGE_TEXTS', '固定导航文字', 'lines', ['全网搜索','搜索','全部','用户','群聊','小程序','机器人','推荐','游戏','二次元','学习','音乐','影视','运动'], '界面适配'),
    ]


def defaults(mode):
    return {f.key: f.default for f in fields(mode)}


def validate(mode, raw, for_run=False):
    output = {}
    for field in fields(mode):
        value = raw.get(field.key, field.default)
        try:
            if field.kind == 'int':
                value = int(str(value).strip())
                if value < field.minimum:
                    raise ValueError(f'应不小于 {field.minimum:g}')
            elif field.kind == 'float':
                value = float(value)
                if not math.isfinite(value) or value < field.minimum:
                    raise ValueError(f'应为不小于 {field.minimum:g} 的有限数值')
            elif field.kind == 'bool':
                if not isinstance(value, bool):
                    raise ValueError('应为开关值')
            elif field.kind == 'lines':
                value = value.splitlines() if isinstance(value, str) else value
                if not isinstance(value, list):
                    raise ValueError('每行填写一项')
                value = list(dict.fromkeys(str(item).strip() for item in value if str(item).strip()))
            else:
                value = str(value).strip()
                if not value:
                    raise ValueError('不能为空')
                if field.kind == 'regex':
                    pattern = re.compile(value)
                    if field.key in {'MEMBER_COUNT_PATTERN', 'NUMBER_TEXT_PATTERN'} and pattern.groups < 1:
                        raise ValueError('需要一个括号捕获组')
                if field.kind in {'save_xlsx', 'open_numbers', 'exe'}:
                    value = str(Path(os.path.expandvars(value)).expanduser().resolve())
                    suffixes = {'save_xlsx': ('.xlsx',), 'open_numbers': ('.xlsx', '.ods'), 'exe': ('.exe',)}[field.kind]
                    if Path(value).suffix.lower() not in suffixes:
                        raise ValueError(f"请选择 {' 或 '.join(suffixes)} 文件")
            output[field.key] = value
        except (ValueError, TypeError, re.error) as error:
            raise ValueError(f'{field.label}：{error}') from error
    if mode == 'add':
        for account in output['EXCLUDED_QQ_ACCOUNTS']:
            if not re.fullmatch(r'[0-9]{5,12}', account):
                raise ValueError('排除账号必须是 5～12 位纯数字 QQ 号')
        if for_run and not Path(output['NUMBER_FILE_PATH']).is_file():
            raise ValueError('号码表不存在，请点击“选择 Excel”或“浏览”选择已有号码表。')
    return output


def config_path(mode):
    return state_home() / f'{mode}_settings.json'


def load_config(mode):
    path = config_path(mode)
    return validate(mode, json.loads(path.read_text(encoding='utf-8'))) if path.is_file() else defaults(mode)


def save_config(mode, config):
    path = config_path(mode)
    temporary = path.with_suffix('.tmp')
    temporary.write_text(json.dumps(validate(mode, config), ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(temporary, path)
    return path


def apply_config(mode, module, config):
    for field in fields(mode):
        value = config[field.key]
        if field.kind == 'regex':
            value = re.compile(value)
        elif field.kind in {'save_xlsx', 'open_numbers', 'exe'}:
            value = Path(value)
        elif field.kind == 'lines':
            value = set(value)
        setattr(module, field.key, value)
