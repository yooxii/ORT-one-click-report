# -*- coding: utf-8 -*-
"""Python 3.4 兼容层。

XP 目标机上运行的代码必须同时满足两件事：

1. 语法与 API 不超出 Python 3.4（XP 上可用的最后一个 CPython）；
2. 控制台代码页是 cp950（繁体中文），直接打印简体中文会抛 ``UnicodeEncodeError``。

本模块把这两件事收在一处：解释器检查、编码安全的输入输出、UTF-8 文件读写、
.NET 风格取值解析（``True`` / ``False`` 文本、两种精度的日期文本）。
"""

import codecs
import datetime
import os
import re
import sys

#: XP 目标机使用的解释器版本
MIN_PYTHON = (3, 4)

#: 开发机验证过的上限；高于它只提示，不阻止运行
MAX_TESTED = (3, 13)

_DATE_FORMATS = (
    "%Y-%m-%d %H:%M:%S.%f",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",
    "%Y/%m/%d",
)

_TRUE_TEXT = ("true", "1", "yes", "y", "on", "是")
_FALSE_TEXT = ("false", "0", "no", "n", "off", "")


def python_version_text():
    """返回 "3.4.4" 这样的版本字符串。"""
    v = sys.version_info
    return "%d.%d.%d" % (v[0], v[1], v[2])


def check_interpreter():
    """检查解释器是否可用，返回 ``(ok, message)``。

    低于 3.4 不可用（XP 上的语法下限）；高于已测版本只给提示，
    因为开发机用的是较新的解释器，发布包必须在 3.4 下构建。
    """
    v = sys.version_info[:2]
    if v < MIN_PYTHON:
        return False, "需要 Python %d.%d 及以上（XP 目标机请用 3.4.x 32 位），当前 %s" % (
            MIN_PYTHON[0],
            MIN_PYTHON[1],
            python_version_text(),
        )
    if v > MAX_TESTED:
        return True, "当前解释器 %s 高于已测上限 %d.%d：可用于开发，发布包必须在 Python 3.4 下构建" % (
            python_version_text(),
            MAX_TESTED[0],
            MAX_TESTED[1],
        )
    return True, "解释器 %s 可用" % python_version_text()


def is_windows():
    return os.name == "nt"


def is_windows_xp():
    """是否运行在 Windows XP（5.1）上。"""
    if not is_windows():
        return False
    try:
        info = sys.getwindowsversion()
        return info[0] == 5 and info[1] == 1
    except Exception:
        return False


def os_description():
    if not is_windows():
        return os.name
    try:
        info = sys.getwindowsversion()
        return "Windows %d.%d build %d (%s)" % (info[0], info[1], info[2], info[3])
    except Exception:
        return "Windows"


def filesystem_encoding():
    """文件系统编码：Windows 上是 ``mbcs``（系统 ANSI 代码页），其它平台为 utf-8。"""
    return sys.getfilesystemencoding() or "ascii"


def can_encode_path(path):
    """路径能否用文件系统编码表示。

    Windows 下 ``mbcs`` 就是 ANSI 代码页。本机是 cp950（繁体），程序目录里只要出现
    简体字（例如仓库名「ORT一键报告」），tkinter / PyInstaller 的模块加载器在按 ANSI
    编码路径时就会抛 ``UnicodeEncodeError``，程序直接起不来。所以启动前先检查并给出人话提示。
    """
    if not path:
        return True
    try:
        path.encode(filesystem_encoding())
        return True
    except (UnicodeEncodeError, LookupError):
        return False


def python_bits():
    """返回 '32' 或 '64'。"""
    return "64" if sys.maxsize > 2 ** 32 else "32"


# --------------------------------------------------------------------------- 输出


def _is_tty(stream):
    try:
        return bool(stream.isatty())
    except Exception:
        return True


def _write_safe(stream, text):
    if stream is None:
        return
    line = text + "\n"
    try:
        stream.write(line)
    except (UnicodeEncodeError, UnicodeDecodeError):
        # 重定向到文件/管道时直接写 UTF-8 字节，日志里保留中文（`--selftest` 落盘、
        # 窗口子系统的重定向输出都是这种情况）；交互式控制台（本机 cp950）则退化为转义，
        # 避免乱码或二次抛错。
        if not _is_tty(stream):
            buffer = getattr(stream, "buffer", None)
            if buffer is not None:
                try:
                    buffer.write(line.encode("utf-8"))
                    buffer.flush()
                    return
                except Exception:
                    pass
        try:
            encoding = getattr(stream, "encoding", None) or "ascii"
            stream.write(line.encode(encoding, "backslashreplace").decode(encoding, "replace"))
        except Exception:
            return
    except Exception:
        return
    try:
        stream.flush()
    except Exception:
        pass


def say(text=""):
    """控制台安全输出。

    cp950 控制台打印简体中文会抛 ``UnicodeEncodeError``，这里退化为反斜杠转义，
    宁可显示 ``\\u70ed`` 也不让程序崩在打印上。
    """
    _write_safe(sys.stdout, "" if text is None else str(text))


def say_err(text=""):
    _write_safe(sys.stderr, "" if text is None else str(text))


# --------------------------------------------------------------------------- 文件


def ensure_dir(path):
    """确保目录存在。"""
    if not path:
        return path
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except OSError:
            if not os.path.isdir(path):
                raise
    return path


def read_text(path):
    """按 UTF-8 读取文本；文件不存在返回 None。

    日志文件带 UTF-8 BOM（给 XP 记事本看的），读取时去掉，调用方拿到的就是正文。
    """
    if not os.path.isfile(path):
        return None
    stream = codecs.open(path, "r", "utf-8")
    try:
        text = stream.read()
    finally:
        stream.close()
    if text and text[0] == u"\ufeff":
        text = text[1:]
    return text


def write_text(path, text):
    """按 UTF-8 写文本。"""
    ensure_dir(os.path.dirname(path))
    stream = codecs.open(path, "w", "utf-8")
    try:
        stream.write(text)
    finally:
        stream.close()
    return path


def open_text(path, mode="r"):
    """返回 UTF-8 文本流（调用方负责 close）。"""
    if mode.startswith("w") or mode.startswith("a"):
        ensure_dir(os.path.dirname(path))
    return codecs.open(path, mode, "utf-8")


def walk_files(root, suffix=None):
    """递归列出文件（不用 ``os.scandir``，3.4 上没有）。"""
    result = []
    if not root or not os.path.isdir(root):
        return result
    for name in os.listdir(root):
        full = os.path.join(root, name)
        if os.path.isdir(full):
            result.extend(walk_files(full, suffix))
        elif suffix is None or name.lower().endswith(suffix.lower()):
            result.append(full)
    return result


def app_dir():
    """程序所在目录：打包后是 exe 目录，开发时是包目录的上一级。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def local_app_data_dir(folder_name):
    """当前 Windows 用户的 %LocalAppData%\\<folder_name>；非 Windows 退化为 ~/.local/share。"""
    if is_windows():
        base = os.environ.get("LOCALAPPDATA")
        if not base:
            base = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    else:
        base = os.path.join(os.path.expanduser("~"), ".local", "share")
    return os.path.join(base, folder_name)


# --------------------------------------------------------------------------- .NET 值解析


def to_bool(text, default=False):
    """解析 .NET 写进库里的布尔文本（'True' / 'False' / '1' / '0' …）。"""
    if text is None:
        return default
    if isinstance(text, bool):
        return text
    if isinstance(text, int):
        return text != 0
    value = str(text).strip().lower()
    if value in _TRUE_TEXT:
        return True
    if value in _FALSE_TEXT:
        return False
    return default


def to_int(text, default=0):
    if text is None or text == "":
        return default
    try:
        return int(str(text).strip())
    except (TypeError, ValueError):
        try:
            return int(float(str(text).strip()))
        except (TypeError, ValueError):
            return default


def to_float(text, default=0.0):
    if text is None or text == "":
        return default
    try:
        return float(str(text).strip())
    except (TypeError, ValueError):
        return default


def bool_text(value):
    """写出 .NET 风格的布尔文本。"""
    return "True" if value else "False"


def parse_datetime_text(text):
    """解析主程序写进库的日期文本，失败返回 None。

    实测两种形态：``2026-09-01 00:00:00`` 与 ``2026-09-29 11:29:17.806635``。
    """
    if text is None:
        return None
    if isinstance(text, datetime.datetime):
        return text
    value = str(text).strip()
    if not value:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(value, fmt)
        except ValueError:
            continue
    if len(value) > 19:
        try:
            return datetime.datetime.strptime(value[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    return None


def format_datetime_text(value):
    """把 datetime 写成主程序使用的日期文本（不带微秒）。"""
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return value.strftime("%Y-%m-%d %H:%M:%S")


#: .NET 日期格式符（长记号优先，避免 MM 被 M 抢先匹配）
_DOTNET_DATE_TOKENS = re.compile(r"(yyyy|yy|MM|dd|HH|mm|ss|M|d|H|m|s)")


def format_dotnet_date(value, pattern):
    """实现主程序邮件模板里的 .NET 日期格式符（``{{EndDate|yyyy/M/d}}``）。

    支持 yyyy yy MM M dd d HH H mm m ss s；其余字符原样保留。
    Python 3.4 的 strftime 没有 ``%-m``，所以这里自己拼，不用 strftime。
    """
    if value is None:
        return ""
    if isinstance(value, str):
        parsed = parse_datetime_text(value)
        if parsed is None:
            return value
        value = parsed

    def _replace(match):
        token = match.group(0)
        if token == "yyyy":
            return "%04d" % value.year
        if token == "yy":
            return "%02d" % (value.year % 100)
        if token == "MM":
            return "%02d" % value.month
        if token == "M":
            return str(value.month)
        if token == "dd":
            return "%02d" % value.day
        if token == "d":
            return str(value.day)
        if token == "HH":
            return "%02d" % value.hour
        if token == "H":
            return str(value.hour)
        if token == "mm":
            return "%02d" % value.minute
        if token == "m":
            return str(value.minute)
        if token == "ss":
            return "%02d" % value.second
        if token == "s":
            return str(value.second)
        return token

    return _DOTNET_DATE_TOKENS.sub(_replace, pattern)
