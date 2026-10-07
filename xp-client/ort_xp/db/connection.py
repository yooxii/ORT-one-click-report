# -*- coding: utf-8 -*-
"""SQLite 数据层：连接、事务、忙等待重试、表结构探测。

与主程序（FreeSql + System.Data.SQLite）**共用同一个库文件**，所以有几条硬约定：

- 表结构由主程序负责创建与迁移，这里只读写数据，**绝不改 ``journal_mode``**；
- 写操作走短事务（``BEGIN IMMEDIATE``），遇锁指数退避重试；
- SQLite 打不开 UNC 路径（``\\\\服务器\\共享\\...``），需要先映射成盘符；
- 任何表在使用前先用 :meth:`Database.table_exists` 探测（部分表由主程序按需创建）。
"""

import ctypes
import os
import sqlite3
import time

from .. import compat

DEFAULT_TIMEOUT_SECONDS = 30
DEFAULT_RETRIES = 5

DRIVE_REMOTE = 4
DRIVE_NO_ROOT_DIR = 1


class DatabaseError(Exception):
    """数据层错误（带排查提示）。"""


def _is_lock_error(exc):
    message = str(exc).lower()
    return "locked" in message or "busy" in message


def drive_type(path):
    """返回 Windows 盘符类型（UNC 直接判为远程）；非 Windows 返回 None。"""
    if not compat.is_windows():
        return None
    if is_network_path(path):
        return DRIVE_REMOTE
    if not path:
        return None
    root = os.path.splitdrive(os.path.abspath(path))[0]
    if not root:
        return None
    try:
        return ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(root + "\\"))
    except Exception:
        return None


def is_network_path(path):
    """UNC 或映射的网络盘都算网络路径。"""
    if not path:
        return False
    if path.startswith("\\\\") or path.startswith("//"):
        return True
    kind = None
    if compat.is_windows():
        root = os.path.splitdrive(os.path.abspath(path))[0]
        if root:
            try:
                kind = ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(root + "\\"))
            except Exception:
                kind = None
    return kind == DRIVE_REMOTE


class Database(object):
    def __init__(self, path, timeout_seconds=DEFAULT_TIMEOUT_SECONDS, retries=DEFAULT_RETRIES, logger=None):
        self.path = path
        self.timeout_seconds = timeout_seconds
        self.retries = retries
        self.logger = logger
        self._connection = None

    # ---------------------------------------------------------------- 基本信息

    def exists(self):
        return os.path.isfile(self.path)

    def size_bytes(self):
        try:
            return os.path.getsize(self.path)
        except OSError:
            return 0

    def is_network(self):
        return is_network_path(self.path)

    def _log(self, message, level="info"):
        if self.logger is None:
            return
        getattr(self.logger, level, self.logger.info)(message)

    # ---------------------------------------------------------------- 连接

    def connect(self):
        if self._connection is not None:
            return self._connection
        if is_network_path(self.path) and self.path.startswith(("\\\\", "//")):
            raise DatabaseError(
                "SQLite 不能直接打开 UNC 路径：%s\n"
                "请先把共享映射成盘符（例如 Z:\\），再把数据文件夹指向 Z:\\..." % self.path
            )
        if not os.path.isdir(os.path.dirname(self.path)):
            raise DatabaseError("数据目录不存在：%s" % os.path.dirname(self.path))
        try:
            # isolation_level=None → 自动提交，事务由 transaction() 显式控制
            connection = sqlite3.connect(self.path, self.timeout_seconds, isolation_level=None)
        except sqlite3.Error as exc:
            raise DatabaseError(self._open_failure_text(exc))
        connection.row_factory = sqlite3.Row
        try:
            connection.execute("PRAGMA busy_timeout=%d" % int(self.timeout_seconds * 1000))
        except sqlite3.Error:
            pass
        self._connection = connection
        return connection

    def _open_failure_text(self, exc):
        lines = [
            "打开数据库失败：%s" % exc,
            "  数据库文件：%s" % self.path,
            "  是否存在：%s" % ("是" if self.exists() else "否"),
            "  网络路径：%s" % ("是" if self.is_network() else "否"),
            "  日志模式：%s" % (self.journal_mode_safe() or "未知"),
            "",
            "常见原因：数据文件夹指向了网络共享；或库处于 WAL 模式而共享上不支持共享内存；"
            "或文件正被主程序独占、权限不足。",
        ]
        return "\n".join(lines)

    def close(self):
        if self._connection is not None:
            try:
                self._connection.close()
            finally:
                self._connection = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
        return False

    # ---------------------------------------------------------------- 查询

    def _run(self, func):
        last_error = None
        for attempt in range(self.retries + 1):
            try:
                return func()
            except sqlite3.OperationalError as exc:
                if not _is_lock_error(exc) or attempt >= self.retries:
                    raise
                last_error = exc
                delay = min(0.25 * (2 ** attempt), 5.0)
                self._log("数据库忙（%s），%.2f 秒后重试（第 %d/%d 次）" % (exc, delay, attempt + 1, self.retries), "warn")
                time.sleep(delay)
        raise last_error

    def query(self, sql, params=()):
        connection = self.connect()
        return self._run(lambda: connection.execute(sql, params).fetchall())

    def query_one(self, sql, params=()):
        connection = self.connect()
        return self._run(lambda: connection.execute(sql, params).fetchone())

    def scalar(self, sql, params=(), default=None):
        row = self.query_one(sql, params)
        if row is None:
            return default
        return row[0]

    def execute(self, sql, params=()):
        connection = self.connect()
        return self._run(lambda: connection.execute(sql, params).rowcount)

    def execute_many(self, sql, seq_of_params):
        connection = self.connect()
        return self._run(lambda: connection.executemany(sql, seq_of_params).rowcount)

    # ---------------------------------------------------------------- 事务

    def transaction(self, immediate=True):
        """事务上下文管理器（默认 ``BEGIN IMMEDIATE``，尽早拿写锁）。"""
        return _Transaction(self, immediate)

    def run_in_transaction(self, func, retries=None):
        """把整个函数放进一个事务执行，遇锁重试整个事务。"""
        limit = self.retries if retries is None else retries
        last_error = None
        for attempt in range(limit + 1):
            try:
                with self.transaction(immediate=True) as connection:
                    return func(connection)
            except sqlite3.OperationalError as exc:
                if not _is_lock_error(exc) or attempt >= limit:
                    raise
                last_error = exc
                delay = min(0.25 * (2 ** attempt), 5.0)
                self._log("事务遇锁（%s），%.2f 秒后重试整个事务" % (exc, delay), "warn")
                time.sleep(delay)
        raise last_error

    # ---------------------------------------------------------------- 结构探测

    def table_exists(self, name):
        try:
            row = self.query_one("SELECT name FROM sqlite_master WHERE type = 'table' AND name = ?", (name,))
        except sqlite3.Error:
            return False
        return row is not None

    def tables(self):
        rows = self.query("SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY name")
        return [row["name"] for row in rows]

    def columns(self, table):
        """返回 ``[{'name','type','notnull','pk','default'}]``。"""
        rows = self.query('PRAGMA table_info("%s")' % table.replace('"', '""'))
        result = []
        for row in rows:
            result.append(
                {
                    "name": row["name"],
                    "type": row["type"],
                    "notnull": row["notnull"],
                    "pk": row["pk"],
                    "default": row["dflt_value"],
                }
            )
        return result

    def column_names(self, table):
        return [column["name"] for column in self.columns(table)]

    def has_column(self, table, column):
        for item in self.columns(table):
            if item["name"].lower() == column.lower():
                return True
        return False

    def journal_mode(self):
        row = self.query_one("PRAGMA journal_mode")
        if row is None:
            return None
        return row[0]

    def journal_mode_safe(self):
        try:
            return self.journal_mode()
        except Exception:
            return None

    # ---------------------------------------------------------------- 健康检查

    def health(self):
        """返回一份可直接打印的诊断信息。"""
        info = {
            "path": self.path,
            "exists": self.exists(),
            "size": self.size_bytes(),
            "network": self.is_network(),
            "sqlite_version": sqlite3.sqlite_version,
            "python_sqlite": sqlite3.version,
            "tables": [],
            "journal_mode": None,
            "error": None,
        }
        if not info["exists"]:
            info["error"] = "数据库文件不存在"
            return info
        try:
            self.connect()
            info["journal_mode"] = self.journal_mode()
            info["tables"] = self.tables()
        except (sqlite3.Error, DatabaseError) as exc:
            info["error"] = str(exc)
        return info

    def describe(self):
        health = self.health()
        if health["error"]:
            return "数据库不可用：%s（%s）" % (health["error"].splitlines()[0], health["path"])
        return "数据库正常：%s（%d 张表，日志模式 %s，%s）" % (
            health["path"],
            len(health["tables"]),
            health["journal_mode"],
            "网络共享" if health["network"] else "本地磁盘",
        )


class _Transaction(object):
    def __init__(self, database, immediate=True):
        self.database = database
        self.immediate = immediate
        self.connection = None

    def __enter__(self):
        self.connection = self.database.connect()
        self.connection.execute("BEGIN IMMEDIATE" if self.immediate else "BEGIN")
        return self.connection

    def __exit__(self, exc_type, exc_value, traceback):
        try:
            if exc_type is None:
                self.connection.execute("COMMIT")
            else:
                self.connection.execute("ROLLBACK")
        except sqlite3.Error:
            try:
                self.connection.execute("ROLLBACK")
            except sqlite3.Error:
                pass
        return False
