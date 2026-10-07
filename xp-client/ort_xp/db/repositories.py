# -*- coding: utf-8 -*-
"""数据访问层：本范围用到的表（plans / requisitions / 下拉取值 / 变更日志）。

与主程序保持一致的地方（依据见 ``docs/02-数据契约.md``）：

- 列名 = C# 属性名，大小写敏感；表结构由主程序负责，本层**只读写、不建表**
  （唯一例外：``plan_change_logs`` 是追加型日志表，缺失时会按 FreeSql 完全相同的 DDL 补建，
  这样即使 XP 端先写第一条日志，主程序的自同步也是空操作）。
- 日期一律写成 ``yyyy-MM-dd HH:mm:ss`` 文本。
- 每次增删改都写 ``plan_change_logs``：主程序的领退与计划**共用**这张表，
  ``PlanId`` 存该表的记录 Id，``Action`` 用中文「新增/编辑/删除」，
  ``Summary`` 与主程序同格式（如 ``新增领退 RT2601 (X100)``），
  ``BeforeJson`` / ``AfterJson`` 是 PascalCase 的紧凑 JSON，日期用 ISO（``2026-10-01T00:00:00``）。
- 审计字段：新增写 ``CreatedBy`` / ``CreatedAt``，编辑写 ``UpdatedBy`` / ``UpdatedAt``。
"""

import collections
import datetime
import json
import sqlite3

from .. import compat
from . import schema_generated

#: 变更日志的动作词（与主程序一致）
ACTION_ADD = "新增"
ACTION_EDIT = "编辑"
ACTION_DELETE = "删除"

#: plan_change_logs 的建表语句（与 FreeSql 对 Models/AdminCatalogs.cs:PlanChangeLog 生成的完全相同）
CHANGE_LOG_DDL = (
    'CREATE TABLE IF NOT EXISTS "plan_change_logs" ('
    '"Id" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, '
    '"Action" NVARCHAR(16) NOT NULL, '
    '"PlanId" INTEGER NOT NULL, '
    '"Summary" NVARCHAR(256), '
    '"BeforeJson" TEXT, '
    '"AfterJson" TEXT, '
    '"Operator" NVARCHAR(64) NOT NULL, '
    '"CreatedAt" DATETIME NOT NULL)'
)


def now_text():
    return compat.format_datetime_text(datetime.datetime.now())


def snapshot_json(table, record):
    """按列顺序生成 PascalCase 紧凑 JSON（对齐 Newtonsoft 的默认序列化）。

    主程序用 ``JsonConvert.SerializeObject(entity)`` 存快照：属性名按声明顺序、不忽略 null、
    日期是 ISO（``2026-10-01T00:00:00``）、中文原样。这里照做，便于两边互相读懂。
    """
    if record is None:
        return None
    keys = record.keys() if hasattr(record, "keys") else record
    ordered = collections.OrderedDict()
    for column in schema_generated.columns_of(table):
        if column not in keys:
            continue
        value = record[column]
        if value is not None and schema_generated.sql_type_of(table, column) == "DATETIME":
            parsed = compat.parse_datetime_text(value)
            if parsed is not None:
                value = parsed.isoformat()
        ordered[column] = value
    return json.dumps(ordered, ensure_ascii=False, separators=(",", ":"))


def to_sql_value(value):
    """Python 值 → SQLite 值：日期转文本、布尔转 0/1。"""
    if value is None or isinstance(value, (int, float, str, bytes)):
        if isinstance(value, bool):
            return 1 if value else 0
        return value
    if isinstance(value, datetime.datetime):
        return compat.format_datetime_text(value)
    if isinstance(value, datetime.date):
        return compat.format_datetime_text(datetime.datetime(value.year, value.month, value.day))
    return str(value)


def order_by_clause(columns, column, descending=False):
    """把「界面点选的列」转成安全的 ``ORDER BY`` 片段。

    只允许白名单里的列（列名来自我们自己的界面定义，仍然显式校验，杜绝把界面字符串
    直接拼进 SQL）；空值/空串排最后，免得一升序就满屏空格子。
    """
    if column not in columns:
        raise ValueError("不允许按 %s 排序" % column)
    direction = "DESC" if descending else "ASC"
    return '("%s" IS NULL OR "%s" = \'\'), "%s" %s' % (column, column, column, direction)


class ChangeLogRepository(object):
    """``plan_change_logs``：领退与计划共用的变更日志（追加型）。"""

    TABLE = "plan_change_logs"
    _ensured = False

    def __init__(self, database, logger=None):
        self._db = database
        self._logger = logger

    def ensure(self):
        """表缺失时按与主程序相同的 DDL 补建（已在库中则什么都不做）。"""
        if self._db.table_exists(self.TABLE):
            return False
        self._db.execute(CHANGE_LOG_DDL)
        if self._logger is not None:
            self._logger.warn("plan_change_logs 表不存在，已按主程序 DDL 补建")
        return True

    def write(self, action, record_id, summary, before, after, operator, table=None):
        """写一条日志；``before`` / ``after`` 传行数据（sqlite3.Row 或 dict），可为 None。"""
        self.ensure()
        before_json = snapshot_json(table, before) if (table and before is not None) else None
        after_json = snapshot_json(table, after) if (table and after is not None) else None
        return self._db.insert(
            'INSERT INTO "plan_change_logs" ("Action", "PlanId", "Summary", "BeforeJson", "AfterJson", '
            '"Operator", "CreatedAt") VALUES (?, ?, ?, ?, ?, ?, ?)',
            (action, record_id, summary, before_json, after_json, operator or "", now_text()),
        )

    def recent(self, limit=50):
        if not self._db.table_exists(self.TABLE):
            return []
        return self._db.query(
            'SELECT "Id", "Action", "PlanId", "Summary", "Operator", "CreatedAt" '
            'FROM "plan_change_logs" ORDER BY "Id" DESC LIMIT ?',
            (limit,),
        )

    def count(self, record_id=None):
        """统计日志条数。

        注意：主程序的计划与领退**共用**这张表，``PlanId`` 存的是各自表的记录 Id（会重号），
        因此按 ``record_id`` 统计只能得到「两边 Id 相同」的合计；要精确区分请按 ``Summary`` 过滤。
        """
        if not self._db.table_exists(self.TABLE):
            return 0
        if record_id is None:
            return self._db.scalar('SELECT COUNT(*) FROM "plan_change_logs"', default=0)
        return self._db.scalar('SELECT COUNT(*) FROM "plan_change_logs" WHERE "PlanId" = ?', (record_id,), default=0)


class BaseTableRepository(object):
    """一张业务表的通用读写（列清单来自 ``schema_generated``）。"""

    TABLE = ""
    #: 用于日志摘要的显示列
    SUMMARY_FIELDS = ()

    def __init__(self, database, change_logs=None, logger=None):
        self._db = database
        self._change_logs = change_logs or ChangeLogRepository(database, logger)
        self._logger = logger

    # ---------------------------------------------------------------- 结构

    @property
    def columns(self):
        return schema_generated.columns_of(self.TABLE)

    def writable_columns(self):
        return tuple(name for name in self.columns if name != "Id")

    def require(self):
        if not self._db.table_exists(self.TABLE):
            raise sqlite3.OperationalError(
                "表 %s 不存在：请先用主程序打开一次该数据文件夹（表结构由主程序创建）" % self.TABLE
            )
        return True

    # ---------------------------------------------------------------- 读

    def get(self, record_id):
        self.require()
        return self._db.query_one('SELECT * FROM "%s" WHERE "Id" = ?' % self.TABLE, (record_id,))

    def list(self, keyword=None, limit=500, order="Id DESC"):
        self.require()
        sql = 'SELECT * FROM "%s"' % self.TABLE
        params = []
        if keyword:
            fields = self.SUMMARY_FIELDS or ()
            clauses = []
            for field in fields:
                clauses.append('"%s" LIKE ?' % field)
                params.append("%" + keyword + "%")
            if clauses:
                sql += " WHERE " + " OR ".join(clauses)
        sql += " ORDER BY " + order + " LIMIT ?"
        params.append(limit)
        return self._db.query(sql, tuple(params))

    def count(self):
        self.require()
        return self._db.scalar('SELECT COUNT(*) FROM "%s"' % self.TABLE, default=0)

    # ---------------------------------------------------------------- 写

    def _clean(self, values):
        known = dict((name, True) for name in self.writable_columns())
        result = collections.OrderedDict()
        for key, value in values.items():
            if key in known:
                result[key] = to_sql_value(value)
        return result

    def insert(self, values, operator):
        """新增一行并写日志，返回新记录 Id。"""
        self.require()
        payload = self._clean(values)
        stamp = now_text()
        payload.setdefault("CreatedBy", operator or "")
        payload.setdefault("CreatedAt", stamp)
        payload["UpdatedBy"] = operator or ""
        payload["UpdatedAt"] = stamp
        columns = list(payload.keys())
        placeholders = ", ".join(["?"] * len(columns))
        sql = 'INSERT INTO "%s" (%s) VALUES (%s)' % (
            self.TABLE,
            ", ".join('"%s"' % name for name in columns),
            placeholders,
        )
        record_id = self._db.insert(sql, tuple(payload[name] for name in columns))
        row = self.get(record_id)
        self._change_logs.write(ACTION_ADD, record_id, self.summary(ACTION_ADD, row), None, row, operator, self.TABLE)
        return record_id

    def update(self, record_id, changes, operator):
        """按字段更新并写日志（含前后快照）；没有任何字段变化时不写库也不写日志。

        与主程序一致：先比较内容，确认真的改了才写 ``UpdatedBy`` / ``UpdatedAt`` 与日志。
        """
        self.require()
        before = self.get(record_id)
        if before is None:
            raise sqlite3.OperationalError("%s 中不存在 Id=%s 的记录" % (self.TABLE, record_id))
        payload = self._clean(changes)
        changed_fields = []
        for name in payload.keys():
            if before[name] != payload[name]:
                changed_fields.append(name)
        if not changed_fields:
            return False
        payload["UpdatedBy"] = operator or ""
        payload["UpdatedAt"] = now_text()
        assignments = ", ".join('"%s" = ?' % name for name in payload.keys())
        params = [payload[name] for name in payload.keys()]
        params.append(record_id)
        self._db.execute('UPDATE "%s" SET %s WHERE "Id" = ?' % (self.TABLE, assignments), tuple(params))
        after = self.get(record_id)
        self._change_logs.write(
            ACTION_EDIT, record_id, self.summary(ACTION_EDIT, after), before, after, operator, self.TABLE
        )
        return True

    def delete(self, record_id, operator):
        self.require()
        before = self.get(record_id)
        if before is None:
            return False
        self._db.execute('DELETE FROM "%s" WHERE "Id" = ?' % self.TABLE, (record_id,))
        self._change_logs.write(ACTION_DELETE, record_id, self.summary(ACTION_DELETE, before), before, None, operator, self.TABLE)
        return True

    def summary(self, action, record):
        return "%s%s" % (action, self.describe(record))

    def describe(self, record):
        return ""

    # ---------------------------------------------------------------- 唯一性

    def exists_value(self, column, value, exclude_id=None):
        """判断某列取值是否已存在（对应库里的唯一索引）。"""
        if value is None or value == "":
            return False
        self.require()
        sql = 'SELECT COUNT(*) FROM "%s" WHERE "%s" = ?' % (self.TABLE, column)
        params = [value]
        if exclude_id:
            sql += ' AND "Id" <> ?'
            params.append(exclude_id)
        return self._db.scalar(sql, tuple(params), default=0) > 0


class PlanRepository(BaseTableRepository):
    TABLE = "plans"
    SUMMARY_FIELDS = ("JobNo", "ModelName", "TestItem", "Owner", "Remark")

    def job_no_exists(self, job_no, exclude_id=None):
        return self.exists_value("JobNo", job_no, exclude_id)

    def describe(self, record):
        return "计划 %s (%s)" % (record["JobNo"] or "", record["ModelName"] or "")

    def deadline_candidates(self, limit_text):
        """结束日期不晚于 ``limit_text`` 且未结案的记录（到期提醒用）。"""
        self.require()
        return self._db.query(
            'SELECT "Id", "JobNo", "ModelName", "Owner", "EndDate", "Status" FROM "plans" '
            'WHERE "EndDate" IS NOT NULL AND TRIM("EndDate") <> \'\' AND "EndDate" <= ? '
            'ORDER BY "EndDate"',
            (limit_text,),
        )


class RequisitionRepository(BaseTableRepository):
    TABLE = "requisitions"
    SUMMARY_FIELDS = ("RequisitionNo", "ModelName", "WorkOrder", "ReturnRtOrder", "Remark")

    def requisition_no_exists(self, requisition_no, exclude_id=None):
        return self.exists_value("RequisitionNo", requisition_no, exclude_id)

    def describe(self, record):
        return "领退 %s (%s)" % (record["RequisitionNo"] or "", record["ModelName"] or "")


class LookupRepository(object):
    """领退/计划表单的下拉取值（对应主程序 AdminService 的查询）。

    机种 → 客户/产品别的规则（主程序 ``AdminService``）：
    - 产品代码 = 机种名第 1-2 位；客户代码 = 机种名第 8-9 位（不足 9 位没有）；
    - 再去 ``code_mappings`` 按 ``CodeType``（P=产品 / C=客户）取名称。
    """

    def __init__(self, database):
        self._db = database

    def _simple_list(self, table, column, where=None):
        if not self._db.table_exists(table):
            return []
        sql = 'SELECT "%s" FROM "%s"' % (column, table)
        if where:
            sql += " WHERE " + where
        sql += ' ORDER BY "%s"' % column
        return [row[column] for row in self._db.query(sql)]

    def stages(self):
        if not self._db.table_exists("stages"):
            return []
        return [row["Name"] for row in self._db.query('SELECT "Name" FROM "stages" ORDER BY "Id"')]

    def test_items(self):
        return self._simple_list("test_items_catalog", "Name")

    def customers(self):
        return self._simple_list("customers", "Name")

    def products(self):
        return self._simple_list("products", "Name")

    def model_mappings(self):
        if not self._db.table_exists("model_mappings"):
            return []
        return self._db.query('SELECT "ModelName", "Product", "Customer" FROM "model_mappings" ORDER BY "ModelName"')

    def find_model_mapping(self, model_name):
        if not model_name or not self._db.table_exists("model_mappings"):
            return None
        return self._db.query_one(
            'SELECT "ModelName", "Product", "Customer" FROM "model_mappings" WHERE "ModelName" = ?',
            (model_name.strip(),),
        )

    def code_mapping(self, code_type, code):
        if not code or not self._db.table_exists("code_mappings"):
            return None
        row = self._db.query_one(
            'SELECT "Name" FROM "code_mappings" WHERE "CodeType" = ? AND "Code" = ?', (code_type, code)
        )
        return row["Name"] if row is not None else None

    @staticmethod
    def model_to_product_code(model_name):
        if not model_name:
            return None
        text = model_name.strip()
        return text[0:2] if len(text) >= 2 else None

    @staticmethod
    def model_to_customer_code(model_name):
        if not model_name:
            return None
        text = model_name.strip()
        return text[7:9] if len(text) >= 9 else None

    def find_product_by_model(self, model_name):
        return self.code_mapping("P", self.model_to_product_code(model_name))

    def find_customer_by_model(self, model_name):
        return self.code_mapping("C", self.model_to_customer_code(model_name))


class Repositories(object):
    """一次性拿到本范围的全部仓储（界面与脚本都用它，避免各自 new）。"""

    def __init__(self, database, logger=None):
        self.database = database
        self.change_logs = ChangeLogRepository(database, logger)
        self.plans = PlanRepository(database, self.change_logs, logger)
        self.requisitions = RequisitionRepository(database, self.change_logs, logger)
        self.lookups = LookupRepository(database)
