# -*- coding: utf-8 -*-
"""顶层异常兜底：把回溯写进 ``Logs`` 目录，并弹一个原生消息框。

XP 目标机上程序是窗口子系统（没有控制台），异常如果不落盘就等于**静默失败** ——
用户只看到窗口闪一下，运维也拿不到线索。所以：

- 任何未捕获异常都写 ``<程序目录>\\Logs\\fatal_YYYYmmdd_HHMMSS.log``（UTF-8）；
- 同时用 ``user32!MessageBoxW``（ctypes，零依赖）弹窗告诉用户日志在哪；
- tkinter 回调里的异常也一并接住（见 ``ui/app.py`` 的 ``report_callback_exception``）。

弹窗失败、日志写失败都不再抛异常：兜底逻辑自己崩掉是最糟的结果。
"""

import datetime
import os
import sys
import traceback

from . import compat

MB_ICONERROR = 0x00000010
MB_TOPMOST = 0x00040000


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
    if not quiet:
        summary = "%s\n\n详细信息已写入：\n%s" % (exc_value, path or "(日志写入失败)")
        show_message("程序遇到错误", summary[:2000])
    return 1


def install(app_dir=None):
    """把兜底挂到 ``sys.excepthook``。"""

    def _hook(exc_type, exc_value, tb):
        handle(exc_type, exc_value, tb, app_dir)

    sys.excepthook = _hook
    return _hook
