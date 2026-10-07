# -*- coding: utf-8 -*-
"""DPAPI（Windows 数据保护 API）封装：只用 ctypes，零第三方依赖。

作用域与主程序一致：``DataProtectionScope.CurrentUser``
（主程序 ``AppSettingsService`` / ``LoginCredentialStore`` 用的就是 CurrentUser），
因此同一台机器、同一个 Windows 用户下，两边可以互相解开对方加密的内容。
"""

import base64
import ctypes
from ctypes import wintypes

_CRYPTPROTECT_UI_FORBIDDEN = 0x01

AVAILABLE = False
_ERROR = ""

if hasattr(ctypes, "windll"):
    try:
        class _DataBlob(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char)),
            ]

        _crypt32 = ctypes.windll.crypt32
        _kernel32 = ctypes.windll.kernel32

        _crypt32.CryptProtectData.argtypes = [
            ctypes.POINTER(_DataBlob),
            ctypes.c_wchar_p,
            ctypes.POINTER(_DataBlob),
            ctypes.c_void_p,
            ctypes.c_void_p,
            wintypes.DWORD,
            ctypes.POINTER(_DataBlob),
        ]
        _crypt32.CryptProtectData.restype = wintypes.BOOL
        _crypt32.CryptUnprotectData.argtypes = [
            ctypes.POINTER(_DataBlob),
            ctypes.POINTER(ctypes.c_wchar_p),
            ctypes.POINTER(_DataBlob),
            ctypes.c_void_p,
            ctypes.c_void_p,
            wintypes.DWORD,
            ctypes.POINTER(_DataBlob),
        ]
        _crypt32.CryptUnprotectData.restype = wintypes.BOOL
        AVAILABLE = True
    except Exception as exc:  # pragma: no cover - 只有非 Windows 会走到
        _ERROR = str(exc)
else:  # pragma: no cover
    _ERROR = "非 Windows 平台没有 DPAPI"


def _make_blob(data):
    """构造 DATA_BLOB，同时返回缓冲区引用（必须持有，否则内存会被回收）。"""
    buffer_ = ctypes.create_string_buffer(data, len(data))
    blob = _DataBlob(len(data), ctypes.cast(buffer_, ctypes.POINTER(ctypes.c_char)))
    return blob, buffer_


def _take_result(blob):
    try:
        data = ctypes.string_at(blob.pbData, blob.cbData)
    finally:
        _kernel32.LocalFree(blob.pbData)
    return data


def protect(text):
    """加密文本，返回 base64 字符串；失败返回 None。"""
    if not AVAILABLE or text is None:
        return None
    blob_in, _keep = _make_blob(text.encode("utf-8"))
    blob_out = _DataBlob()
    ok = _crypt32.CryptProtectData(
        ctypes.byref(blob_in), None, None, None, None, _CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out)
    )
    if not ok:
        return None
    return base64.b64encode(_take_result(blob_out)).decode("ascii")


def unprotect(encoded):
    """解密 base64 文本；失败（换机器 / 换用户 / 内容不是 DPAPI 密文）返回 None。"""
    if not AVAILABLE or not encoded:
        return None
    try:
        data = base64.b64decode(encoded.encode("ascii"))
    except Exception:
        return None
    blob_in, _keep = _make_blob(data)
    blob_out = _DataBlob()
    description = ctypes.c_wchar_p()
    ok = _crypt32.CryptUnprotectData(
        ctypes.byref(blob_in),
        ctypes.byref(description),
        None,
        None,
        None,
        _CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(blob_out),
    )
    if not ok:
        return None
    try:
        return _take_result(blob_out).decode("utf-8")
    except UnicodeDecodeError:
        return None


def status_text():
    if AVAILABLE:
        return "DPAPI 可用（CurrentUser）"
    return "DPAPI 不可用：%s" % _ERROR
