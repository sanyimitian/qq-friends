"""Separate worker process; cooperative cancellation with a parent-owned job."""
from pathlib import Path
import json
import os
import sys
import time
import traceback


def run_worker(mode, config_file, stop_name, ready_name, log_file, result_file, simulate=False):
    import win32event
    import win32api
    from task_control import StopRequested, check_stop, set_stop_event, InterruptibleTime
    stream = open(log_file, 'a', encoding='utf-8', buffering=1)
    sys.stdout = stream
    sys.stderr = stream
    stop = win32event.OpenEvent(0x00100000, False, stop_name)
    ready = win32event.OpenEvent(0x00100000, False, ready_name)
    mutex = None
    owned = False
    outcome = {'status': 'error', 'message': '任务未完成'}
    code = 1
    set_stop_event(stop)
    try:
        if win32event.WaitForSingleObject(ready, 20000) != win32event.WAIT_OBJECT_0:
            raise RuntimeError('启动协调超时，请重新开始。')
        check_stop()
        mutex = win32event.CreateMutex(None, False, r'Local\QQDesktopAssistant.Automation')
        result = win32event.WaitForSingleObject(mutex, 0)
        if result not in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
            raise RuntimeError('另一个 QQ 任务正在运行，请先停止它再开始。')
        owned = True
        if simulate:
            print('模拟任务已启动，不连接 QQ。', flush=True)
            while True:
                check_stop()
                InterruptibleTime().sleep(0.1)
        from gui_config import validate, apply_config
        config = validate(mode, json.loads(Path(config_file).read_text(encoding='utf-8')), for_run=True)
        if mode == 'capture':
            import zq_qq as module
        else:
            import open_qq_add_friend as module
        apply_config(mode, module, config)
        module.time = InterruptibleTime()
        if mode == 'capture':
            Path(config['NUMBER_WORKBOOK_PATH']).parent.mkdir(parents=True, exist_ok=True)
        else:
            module.ensure_number_file_writable(config['NUMBER_FILE_PATH'])
            from odf.opendocument import load
            from odf.table import Table, TableCell, TableRow
            from odf.teletype import extractText
            numbers, document, _, _ = module.read_number_rows(load, Table, TableRow, TableCell, extractText)
            if hasattr(document, 'close'):
                document.close()
            # 批次数量在 main 恢复上次未写回的记录后检查。
        target = os.path.normcase(str(Path(module.__file__).resolve()))

        def trace(frame, event, arg):
            if event in ('line', 'call') and os.path.normcase(frame.f_code.co_filename) == target:
                check_stop()
            return trace

        print('任务开始。点击窗口的“停止”可结束当前任务。', flush=True)
        sys.settrace(trace)
        module.main()
        sys.settrace(None)
        outcome = {'status': 'finished', 'message': '流程已结束，请查看日志和结果文件。'}
        code = 0
        print(outcome['message'], flush=True)
    except (StopRequested, KeyboardInterrupt):
        sys.settrace(None)
        outcome = {'status': 'stopped', 'message': '任务已停止，已保存的数据保留。'}
        code = 130
        print(outcome['message'], flush=True)
    except BaseException as error:
        sys.settrace(None)
        outcome = {'status': 'error', 'message': str(error) or type(error).__name__}
        print('任务失败：' + outcome['message'], flush=True)
        traceback.print_exc()
    finally:
        sys.settrace(None)
        if owned:
            win32event.ReleaseMutex(mutex)
        if mutex is not None:
            win32api.CloseHandle(mutex)
        set_stop_event(None)
        win32api.CloseHandle(stop)
        win32api.CloseHandle(ready)
        path = Path(result_file)
        temporary = path.with_suffix('.tmp')
        temporary.write_text(json.dumps(outcome, ensure_ascii=False), encoding='utf-8')
        os.replace(temporary, path)
        stream.flush()
    return code
