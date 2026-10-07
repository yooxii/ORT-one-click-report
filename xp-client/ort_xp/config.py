# -*- coding: utf-8 -*-
"""配置：数据目录解析、数据库路径、``app_settings`` 键值读写、邮件设置。

与主程序的对应关系见 ``docs/02-数据契约.md``。
"""

import collections
import json
import os

from . import compat, dpapi
from .credentials import LocalCredentialStore

#: 覆盖数据目录的环境变量（运维/排错用，优先级最高）
ENV_DATA_FOLDER = "ORT_XP_DATA_FOLDER"

#: 数据库文件名（与主程序一致）
DB_FILE_NAME = "ort_plans.db"

#: 本机设置文件名（与主程序同名同结构）
LOCAL_SETTINGS_FILE = "local_settings.json"


# --------------------------------------------------------------------------- 本机设置


class LocalSettings(object):
    """本机设置（``<程序目录>\\Data\\local_settings.json``），键名与主程序一致。"""

    def __init__(self, data_folder=None, ate_data_path=None, emi_data_path=None):
        self.data_folder = data_folder
        self.ate_data_path = ate_data_path
        self.emi_data_path = emi_data_path

    @staticmethod
    def load(path):
        settings = LocalSettings()
        text = compat.read_text(path)
        if not text:
            return settings
        try:
            payload = json.loads(text)
        except ValueError:
            return settings
        if not isinstance(payload, dict):
            return settings
        settings.data_folder = payload.get("DataFolder")
        settings.ate_data_path = payload.get("AteDataPath")
        settings.emi_data_path = payload.get("EmiDataPath")
        return settings

    def save(self, path):
        payload = collections.OrderedDict()
        payload["DataFolder"] = self.data_folder
        payload["AteDataPath"] = self.ate_data_path
        payload["EmiDataPath"] = self.emi_data_path
        compat.write_text(path, json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))
        return path


def local_settings_path(app_directory=None):
    return os.path.join(app_directory or compat.app_dir(), "Data", LOCAL_SETTINGS_FILE)


def resolve_data_dir(app_directory=None, environ=None):
    """解析数据目录，返回 ``(目录, 来源说明)``。

    顺序：环境变量 → 本机设置 DataFolder → ``<程序目录>\\Data``。
    """
    app_directory = app_directory or compat.app_dir()
    environ = os.environ if environ is None else environ
    env_value = (environ.get(ENV_DATA_FOLDER) or "").strip()
    if env_value:
        return env_value, "环境变量 %s" % ENV_DATA_FOLDER
    settings = LocalSettings.load(local_settings_path(app_directory))
    if settings.data_folder and settings.data_folder.strip():
        return settings.data_folder.strip(), "本机设置 DataFolder"
    return os.path.join(app_directory, "Data"), "默认（程序目录\\Data）"


def db_path(data_dir):
    return os.path.join(data_dir, DB_FILE_NAME)


def is_unc_path(path):
    if not path:
        return False
    return path.startswith("\\\\") or path.startswith("//")


# --------------------------------------------------------------------------- app_settings


class AppSettingsStore(object):
    """``app_settings`` 键值表（与主程序同格式：值一律文本，布尔写 True/False）。"""

    def __init__(self, database):
        self._db = database

    def available(self):
        return self._db.table_exists("app_settings")

    def all(self):
        if not self.available():
            return {}
        result = {}
        for row in self._db.query("SELECT Key, Value FROM app_settings"):
            result[row["Key"]] = row["Value"]
        return result

    def get(self, key, default=None):
        if not self.available():
            return default
        row = self._db.query_one("SELECT Value FROM app_settings WHERE Key = ?", (key,))
        if row is None or row["Value"] is None:
            return default
        return row["Value"]

    def get_bool(self, key, default=False):
        return compat.to_bool(self.get(key), default)

    def get_int(self, key, default=0):
        return compat.to_int(self.get(key), default)

    def get_str(self, key, default=""):
        value = self.get(key)
        if value is None:
            return default
        return value

    def set(self, key, value):
        """写入单个键（存在则更新，不存在则插入）；值的文本化与主程序一致。"""
        if isinstance(value, bool):
            text = compat.bool_text(value)
        elif value is None:
            text = None
        else:
            text = str(value)
        with self._db.transaction(immediate=True) as conn:
            cursor = conn.execute("UPDATE app_settings SET Value = ? WHERE Key = ?", (text, key))
            if cursor.rowcount == 0:
                conn.execute("INSERT INTO app_settings (Key, Value) VALUES (?, ?)", (key, text))
        return text

    def set_many(self, mapping):
        for key, value in mapping.items():
            self.set(key, value)


# --------------------------------------------------------------------------- 邮件设置


class MailSettings(object):
    """SMTP 与提醒设置，字段名对应 ``app_settings`` 里的 ``mail.*`` 键。"""

    def __init__(self):
        self.enabled = False
        self.notice_enabled = True
        self.warning_enabled = True
        self.host = ""
        self.port = 25
        self.security = "None"
        self.ignore_cert_errors = False
        self.use_default_credentials = False
        self.username = ""
        self.password = ""
        self.password_source = "无"
        self.from_address = ""
        self.from_name = ""
        self.cc_list = ""
        self.body_is_html = False
        self.timeout_seconds = 30
        self.warning_days_before = 3
        self.warning_include_overdue = True
        self.dedupe_days = 1
        self.cc_admins = {}
        self.templates = {}

    @staticmethod
    def from_store(store, credential_store=None):
        """从 ``app_settings`` 读取，并按需回落到本机口令文件。"""
        settings = MailSettings()
        if not store.available():
            return settings
        settings.enabled = store.get_bool("mail.enabled", False)
        settings.notice_enabled = store.get_bool("mail.noticeEnabled", True)
        settings.warning_enabled = store.get_bool("mail.warningEnabled", True)
        settings.host = store.get_str("mail.host")
        settings.port = store.get_int("mail.port", 25)
        settings.security = store.get_str("mail.security", "None") or "None"
        settings.ignore_cert_errors = store.get_bool("mail.ignoreCertErrors", False)
        settings.use_default_credentials = store.get_bool("mail.useDefaultCredentials", False)
        settings.username = store.get_str("mail.username")
        settings.from_address = store.get_str("mail.fromAddress")
        settings.from_name = store.get_str("mail.fromName")
        settings.cc_list = store.get_str("mail.ccList")
        settings.body_is_html = store.get_bool("mail.bodyIsHtml", False)
        settings.timeout_seconds = store.get_int("mail.timeoutSeconds", 30)
        settings.warning_days_before = store.get_int("mail.warningDaysBefore", 3)
        settings.warning_include_overdue = store.get_bool("mail.warningIncludeOverdue", True)
        settings.dedupe_days = store.get_int("mail.dedupeDays", 1)

        for row in store.all().items():
            key = row[0]
            value = row[1]
            if key.startswith("mail.ccAdmin."):
                settings.cc_admins[key[len("mail.ccAdmin."):]] = compat.to_bool(value, False)
            elif key.startswith("mail.template."):
                settings.templates[key[len("mail.template."):]] = value

        # 密码：先试共享库里的 DPAPI 密文（同机同用户才可能成功），再回落到本机口令文件
        encrypted = store.get_str("mail.passwordEnc")
        if encrypted:
            password = dpapi.unprotect(encrypted)
            if password:
                settings.password = password
                settings.password_source = "共享库（mail.passwordEnc）"
        if not settings.password and credential_store is not None:
            password = credential_store.load_mail_password()
            if password:
                settings.password = password
                settings.password_source = "本机口令文件"
        return settings

    def template(self, kind, subject):
        key = "%s.%s" % (kind, "subject" if subject else "body")
        return self.templates.get(key)

    def recipients_cc(self):
        """固定抄送列表（分号或逗号分隔）。"""
        text = self.cc_list or ""
        for separator in (";", "，", ","):
            text = text.replace(separator, ";")
        return [item.strip() for item in text.split(";") if item.strip()]

    def describe(self):
        """一行摘要（不含密码），用于日志与界面状态栏。"""
        return "SMTP %s:%d 安全=%s 发件=%s 密码来源=%s 启用=%s" % (
            self.host or "(未配置)",
            self.port,
            self.security,
            self.from_address or "(未配置)",
            self.password_source,
            "是" if self.enabled else "否",
        )
