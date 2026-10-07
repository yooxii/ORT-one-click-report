# -*- coding: utf-8 -*-
"""顶层异常兜底与「启动期问题」的可见化。

XP 目标机上程序是窗口子系统（没有控制台，``sys.stderr`` 是 ``None``），错误如果不落盘、
不弹窗就等于**静默失败** —— 用户只看到「双击没反应」，运维也拿不到线索。所以：

- 任何未捕获异常都写 ``<程序目录>\\Logs\\fatal_YYYYmmdd_HHMMSS.log``（UTF-8）；
- 同时用 ``user32!MessageBoxW``（ctypes，零依赖）弹窗告诉用户日志在哪；
- 启动期的「可预期」问题（找不到数据库、路径含非法字符、tkinter 起不来）走
  :func:`report`：写进 ``Logs\\ort_xp.log`` 并弹窗给出可操作的提示；
- tkinter 回调里的异常也一并接住（见 ``ui/app.py`` 的 ``report_callback_exception``）。

无人值守场景（``--no-dialog``、``--selftest`` 等）用 :func:`set_dialogs_enabled`
关掉弹窗，避免脚本被消息框卡住。

弹窗失败、日志写失败都不再抛异常：兜底逻辑自己崩掉是最糟的结果。
"""

import datetime
import os
import sys
import traceback

from . import compat

MB_ICONERROR = 0x00000010
MB_TOPMOST = 0x00040000

#: 设成 1（或 true/yes/on）时一律不弹窗，只写日志（运维/脚本用，等价于 --no-dialog）
NO_DIALOG_ENV = "ORT_XP_NO_DIALOG"

_dialogs_enabled = [True]


def dialogs_enabled():
    """当前是否允许弹窗。"""
    if (os.environ.get(NO_DIALOG_ENV) or "").strip().lower() in ("1", "true", "yes", "on"):
        return False
    return _dialogs_enabled[0]


def set_dialogs_enabled(value):
    """开关弹窗（``--no-dialog`` 与无界面模式用）。"""
    _dialogs_enabled[0] = bool(value)


def log_dir(app_dir=None):
    return os.path.join(app_dir or compat.app_dir(), "Logs")


def format_exception(exc_type, exc_value, tb, app_dir=None):
    """拼出带环境信息的完整回溯文本。"""
    lines = [
        "时间：%s" % compat.format_datetime_text(datetime.datetime.now()),
        "程序目录：%s" % (app_dir or compat.app_dir()),
        "解释器：Python %s（%s 位）" % (compat.python_version_text(), compat.python_bits()),
        "操作系统：%s%s" % (compat.os_description(), "（Windows XP）" if compat.is_windows_xp() else ""),
        "打包运行：%s" % ("是" if getattr(sys, "frozen", False) else "否"),
        "",
    ]
    return "\n".join(lines) + "".join(traceback.format_exception(exc_type, exc_value, tb))


def write_log(text, app_dir=None):
    """写一份致命错误日志，返回文件路径。"""
    directory = log_dir(app_dir)
    compat.ensure_dir(directory)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(directory, "fatal_%s.log" % stamp)
    compat.write_text(path, text)
    return path


def show_message(title, text):
    """尽量弹原生消息框；弹不出来也不抛。"""
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, text, title, MB_ICONERROR | MB_TOPMOST)
        return True
    except Exception:
        return False


def _write_log_line(text, app_dir=None):
    """把一行 ERROR 直接追加进 ``Logs\\ort_xp.log``（格式与 logging_setup 一致）。

    只在 logger 还没建立时用得上：启动最早期的错误如果只依赖 logging，就可能一条都留不下。
    """
    try:
        directory = log_dir(app_dir)
        compat.ensure_dir(directory)
        path = os.path.join(directory, "ort_xp.log")
        # 新建文件时补 UTF-8 BOM，与 logging_setup 保持一致（XP 记事本靠它认编码）
        is_new = (not os.path.isfile(path)) or os.path.getsize(path) == 0
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S,%f")[:-3]
        stream = compat.open_text(path, "a")
        try:
            if is_new:
                stream.write(u"\ufeff")
            stream.write("%s [ERROR] ort_xp: %s%s" % (stamp, text, os.linesep))
        finally:
            stream.close()
        return True
    except Exception:
        return False


def report(title, message, app_dir=None, logger=None):
    """启动期问题的统一出口：先落日志，再（允许时）弹窗，返回建议的退出码 2。

    ``logger`` 传了就用它记 ERROR（含多行正文）；没传或没有处理器时直接写日志文件，
    保证窗口子系统下也一定留得下现场。
    """
    text = "%s：%s" % (title, message)
    logged = False
    if logger is not None and getattr(logger, "handlers", None):
        try:
            logger.error(text)
            logged = True
        except Exception:
            logged = False
    if not logged:
        _write_log_line(text, app_dir)
    if dialogs_enabled():
        hint = "%s\n\n日志：%s" % (message, os.path.join(log_dir(app_dir), "ort_xp.log"))
        show_message(title, hint[:2000])
    return 2


def handle(exc_type, exc_value, tb, app_dir=None, quiet=False):
    """统一兜底入口：写日志 + 弹窗，返回建议的退出码。"""
    try:
        text = format_exception(exc_type, exc_value, tb, app_dir)
    except Exception:
        text = "异常信息格式化失败\n"
    path = None
    try:
        path = write_log(text, app_dir)
    except Exception:
        pass
    if not quiet and dialogs_enabled():
        summary = "%s\n\n详细信息已写入：\n%s" % (exc_value, path or "(日志写入失败)")
        show_message("程序遇到错误", summary[:2000])
    return 1


def install(app_dir=None):
    """把兜底挂到 ``sys.excepthook``。"""

    def _hook(exc_type, exc_value, tb):
        handle(exc_type, exc_value, tb, app_dir)

    sys.excepthook = _hook
    return _hook
