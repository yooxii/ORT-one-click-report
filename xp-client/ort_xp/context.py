# -*- coding: utf-8 -*-
"""应用上下文：把配置、数据层、业务服务装配到一处，界面与自检共用。

界面（``ui/app.py``）、命令行自检（``__main__.py``）和 ``tools/check_env.py``
都只依赖这个类，避免各自重复解析路径与装配服务。
"""

import datetime
import os
import sqlite3

from . import compat, config, credentials, dpapi, logging_setup
from .db.connection import Database, DatabaseError
from .services.auth import AuthService
from .services.mail import MailService, PlanDeadlineReminder, is_valid_address
from .version import APP_NAME, VERSION, BUILD_STAGE, TARGET_OS, TARGET_PYTHON

#: 本范围必须存在的表（缺失说明数据文件夹不是主程序的，或库还没建好）
REQUIRED_TABLES = ("users", "user_roles", "plans", "requisitions", "app_settings")


class AppContext(object):
    def __init__(self, app_directory=None, data_folder=None, logger=None):
        self.app_directory = app_directory or compat.app_dir()
        if data_folder:
            self.data_dir = data_folder
            self.data_source = "命令行指定"
        else:
            self.data_dir, self.data_source = config.resolve_data_dir(self.app_directory)
        self.db_path = config.db_path(self.data_dir)
        self.logger = logger if logger is not None else logging_setup.setup_logging(self.app_directory)

        self.database = Database(self.db_path, logger=self.logger)
        self.credentials = credentials.LocalCredentialStore()
        self.settings_store = None
        self.mail_settings = None
        self.auth = None
        self.mail = None
        self.reminder = None

    # ---------------------------------------------------------------- 装配

    def open(self):
        """连接数据库并装配服务；失败抛出 ``DatabaseError``。"""
        self.database.connect()
        self.settings_store = config.AppSettingsStore(self.database)
        self.mail_settings = config.MailSettings.from_store(self.settings_store, self.credentials)
        self.auth = AuthService(self.database, self.credentials, self.logger)
        self.mail = MailService(
            self.database,
            self.mail_settings,
            self.credentials,
            self.logger,
            ca_bundle=self.ca_bundle(),
        )
        self.reminder = PlanDeadlineReminder(self.database, self.mail, self.mail_settings, self.logger)
        if self.logger is not None:
            self.logger.info("已连接数据库：%s（来源：%s）" % (self.db_path, self.data_source))
        return self

    def close(self):
        self.database.close()

    def ca_bundle(self):
        """随程序分发的 CA 根证书包（XP 系统证书库已过期，必须自带）。"""
        candidates = [
            os.path.join(self.app_directory, "cacert.pem"),
            os.path.join(self.app_directory, "ca", "cacert.pem"),
        ]
        for path in candidates:
            if os.path.isfile(path):
                return path
        return None

    @property
    def local_settings_path(self):
        return config.local_settings_path(self.app_directory)

    def save_data_folder(self, folder):
        """把数据文件夹写进本机设置（``Data\\local_settings.json``，与主程序同结构）。"""
        settings = config.LocalSettings.load(self.local_settings_path)
        settings.data_folder = folder
        settings.save(self.local_settings_path)
        self.data_dir = folder
        self.db_path = config.db_path(folder)
        self.database.close()
        self.database = Database(self.db_path, logger=self.logger)
        return self.local_settings_path

    # ---------------------------------------------------------------- 自检

    def selftest(self):
        """无界面自检，返回 ``[(项目, 是否通过, 说明)]``。"""
        checks = []

        ok, message = compat.check_interpreter()
        checks.append(("解释器版本", ok, message))

        checks.append(("解释器位数", True, "%s 位" % compat.python_bits()))
        checks.append(("操作系统", True, "%s%s" % (compat.os_description(), "（XP）" if compat.is_windows_xp() else "")))

        missing = []
        for name in ("sqlite3", "ssl", "smtplib", "tkinter", "ctypes", "email", "hashlib", "json", "uuid"):
            try:
                __import__(name)
            except ImportError:
                missing.append(name)
        checks.append(("标准库依赖", not missing, "缺少：%s" % ", ".join(missing) if missing else "全部可用"))

        try:
            import ssl as _ssl

            checks.append(
                (
                    "TLS 能力",
                    True,
                    "%s；协议常量 PROTOCOL_TLSv1_2 可用（XP 上靠它绕开 Schannel 只支持 TLS 1.0 的限制）"
                    % getattr(_ssl, "OPENSSL_VERSION", "OpenSSL 版本未知"),
                )
            )
        except Exception as exc:
            checks.append(("TLS 能力", False, str(exc)))

        checks.append(("DPAPI", dpapi.AVAILABLE, dpapi.status_text()))

        checks.append(("数据目录", os.path.isdir(self.data_dir), "%s（来源：%s）" % (self.data_dir, self.data_source)))

        health = self.database.health()
        db_ok = health["error"] is None
        checks.append(
            (
                "数据库连接",
                db_ok,
                health["error"] or "%d 张表，日志模式 %s，%s，SQLite %s"
                % (
                    len(health["tables"]),
                    health["journal_mode"],
                    "网络共享" if health["network"] else "本地磁盘",
                    health["sqlite_version"],
                ),
            )
        )

        if db_ok:
            if health["network"] and health["journal_mode"] == "wal":
                checks.append(("WAL + 共享目录", False, "网络共享上 WAL 不可用，请让主程序改回 TRUNCATE 模式"))
            names = health["tables"]
            absent = [name for name in REQUIRED_TABLES if name not in names]
            checks.append(("关键表", not absent, "缺少：%s" % ", ".join(absent) if absent else ", ".join(REQUIRED_TABLES)))
            if "users" in names:
                total = self.database.scalar("SELECT COUNT(*) FROM users", default=0)
                active = self.database.scalar("SELECT COUNT(*) FROM users WHERE IsActive = 1", default=0)
                checks.append(("账号可读", total > 0, "共 %d 个账号，启用 %d 个" % (total, active)))
            if self.settings_store is not None and self.settings_store.available():
                checks.append(("设置可读", True, "app_settings 共 %d 个键" % len(self.settings_store.all())))
            if self.mail_settings is not None:
                ready, reason = self.mail.is_ready()
                checks.append(
                    (
                        "邮件配置",
                        True,
                        "可发信" if ready else "未就绪：%s（%s）" % (reason, self.mail_settings.describe()),
                    )
                )

        if dpapi.AVAILABLE:
            probe = "ort-xp-selftest"
            encrypted = dpapi.protect(probe)
            decrypted = dpapi.unprotect(encrypted) if encrypted else None
            checks.append(("DPAPI 往返", decrypted == probe, "加解密%s" % ("一致" if decrypted == probe else "失败")))
        else:
            checks.append(("DPAPI 往返", False, "DPAPI 不可用，跳过"))

        checks.append(("构建阶段", True, "%s / 目标：%s + Python %s" % (BUILD_STAGE, TARGET_OS, TARGET_PYTHON)))
        return checks

    def selftest_text(self):
        lines = ["%s v%s" % (APP_NAME, VERSION), ""]
        for name, ok, detail in self.selftest():
            lines.append("[%s] %s：%s" % ("通过" if ok else "失败", name, detail))
        return "\n".join(lines)

    def selftest_ok(self):
        for _name, ok, _detail in self.selftest():
            if not ok:
                return False
        return True


def run_reminder(context, dry_run=True):
    """执行一轮计划到期提醒（默认演练，不发真信）。"""
    summary = context.reminder.run(dry_run=dry_run)
    lines = [
        "计划到期提醒（%s）" % ("演练，不真正发送" if dry_run else "实际发送"),
        "候选计划：%d；发送成功：%d；跳过：%d；失败：%d" % (summary["candidates"], summary["sent"], summary["skipped"], summary["failed"]),
    ]
    lines.extend(["  " + item for item in summary["details"]])
    return "\n".join(lines)
