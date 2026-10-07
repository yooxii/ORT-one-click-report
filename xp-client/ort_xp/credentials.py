# -*- coding: utf-8 -*-
"""本机凭据存取（当前 Windows 用户目录 + DPAPI 加密）。

**故意与主程序分文件**：主程序用 ``%LocalAppData%\\ORT实验室管理系统\\login.json``，
本客户端用 ``%LocalAppData%\\ORT实验室管理系统-XP\\login_xp.json``，
避免两边互相覆盖对方看不懂的字段。

存放内容：
- ``login_xp.json``：记住登录（用户名 + DPAPI 密文 + 到期时间）
- ``mail_credential.json``：本机 SMTP 口令（DPAPI 密文）
"""

import datetime
import json
import os

from . import compat, dpapi
from .version import APP_FOLDER_NAME

LOGIN_FILE = "login_xp.json"
MAIL_FILE = "mail_credential.json"


class LocalCredentialStore(object):
    def __init__(self, folder_name=APP_FOLDER_NAME):
        self.folder_name = folder_name
        self.base_dir = compat.local_app_data_dir(folder_name)

    # ---------------------------------------------------------------- 路径

    @property
    def login_path(self):
        return os.path.join(self.base_dir, LOGIN_FILE)

    @property
    def mail_password_path(self):
        return os.path.join(self.base_dir, MAIL_FILE)

    # ---------------------------------------------------------------- 内部

    def _read_json(self, path):
        text = compat.read_text(path)
        if not text:
            return None
        try:
            return json.loads(text)
        except ValueError:
            return None

    def _write_json(self, path, payload):
        compat.write_text(path, json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))

    def _delete(self, path):
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass

    # ---------------------------------------------------------------- 记住登录

    def save_login(self, username, password, days=7):
        """保存登录凭据（DPAPI 加密）。DPAPI 不可用时返回 False，调用方应提示用户。"""
        encrypted = dpapi.protect(password or "")
        if not encrypted:
            return False
        expiry = datetime.datetime.now() + datetime.timedelta(days=days)
        self._write_json(
            self.login_path,
            {
                "Username": username,
                "PasswordEnc": encrypted,
                "Expiry": compat.format_datetime_text(expiry),
            },
        )
        return True

    def load_login(self):
        """返回 ``(username, password)``；没有、过期或解不开都返回 None。"""
        payload = self._read_json(self.login_path)
        if not payload:
            return None
        username = payload.get("Username")
        encrypted = payload.get("PasswordEnc")
        if not username or not encrypted:
            return None
        expiry = compat.parse_datetime_text(payload.get("Expiry"))
        if expiry is not None and expiry < datetime.datetime.now():
            self.clear_login()
            return None
        password = dpapi.unprotect(encrypted)
        if password is None:
            # 换了机器或换了 Windows 用户 → 密文解不开，清掉避免每次启动都失败
            self.clear_login()
            return None
        return username, password

    def clear_login(self):
        self._delete(self.login_path)

    # ---------------------------------------------------------------- 本机 SMTP 口令

    def save_mail_password(self, password):
        if not password:
            self.clear_mail_password()
            return True
        encrypted = dpapi.protect(password)
        if not encrypted:
            return False
        self._write_json(self.mail_password_path, {"PasswordEnc": encrypted})
        return True

    def load_mail_password(self):
        payload = self._read_json(self.mail_password_path)
        if not payload:
            return None
        return dpapi.unprotect(payload.get("PasswordEnc"))

    def clear_mail_password(self):
        self._delete(self.mail_password_path)
