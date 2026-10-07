# -*- coding: utf-8 -*-
"""从主程序 C# 模型生成表结构清单（本客户端的数据层契约）。

输入：``../Models/*.cs``（主程序 FreeSql 实体类）
输出：
- ``docs/schema.md``                 —— 人读的表与字段说明
- ``ort_xp/db/schema_generated.py``  —— 机读的表→列映射（数据层校验用）

用法::

    python tools/extract_schema.py
    python tools/extract_schema.py --db "..\\bin\\Debug\\Data\\ort_plans.db"   # 顺便与真实库比对

生成文件不要手工编辑：改了主程序模型就重新跑一次。
"""

import argparse
import codecs
import os
import re
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)

from ort_xp import compat  # noqa: E402

MODELS_DIR = os.path.join(_PARENT, "..", "Models")
SCHEMA_DOC = os.path.join(_PARENT, "docs", "schema.md")
SCHEMA_MODULE = os.path.join(_PARENT, "ort_xp", "db", "schema_generated.py")

TABLE_RE = re.compile(r'\[Table\(Name\s*=\s*"([^"]+)"\)\]')
CLASS_RE = re.compile(r"^\s*public\s+(?:sealed\s+|partial\s+)?class\s+(\w+)")
PROPERTY_RE = re.compile(
    r"^\s*public\s+(?:static\s+)?([A-Za-z_][\w\.\<\>\[\]\?]*)\s+([A-Za-z_]\w*)\s*(\{|=>|$)"
)
COLUMN_RE = re.compile(r"\[Column\(([^)]*)\)\]")
INDEX_RE = re.compile(r"\[Index\(([^)]*)\)\]")
ATTR_PAIR_RE = re.compile(r"(\w+)\s*=\s*([^,)]+)")


def map_sql_type(csharp_type, string_length, is_nullable):
    """把 C# 类型映射成 SQLite 声明类型（与 FreeSql 生成的实际类型同风格）。"""
    value = csharp_type.replace("System.", "").strip()
    array = value.endswith("[]")
    if value.startswith("string"):
        return "NVARCHAR(%s)" % (string_length if string_length else 0)
    if value.startswith("DateTime"):
        return "DATETIME"
    if value.startswith("TimeSpan"):
        return "TIME"
    if value.startswith("bool"):
        return "BOOLEAN"
    if value.startswith(("long", "int", "short", "byte")):
        return "INTEGER"
    if value.startswith(("double", "float", "decimal")):
        return "DECIMAL"
    if array or value.startswith("byte"):
        return "BLOB"
    return value.upper()


def clean_doc(lines):
    text = []
    for line in lines:
        item = line.strip()
        item = item.lstrip("/").strip()
        if item.startswith("<") or item.startswith("</"):
            continue
        if item:
            text.append(item)
    return " ".join(text)


def parse_models(models_dir):
    tables = []  # [{'table':..,'class':..,'file':..,'indexes':[..],'columns':[..]}]
    if not os.path.isdir(models_dir):
        raise IOError("找不到模型目录：%s" % models_dir)

    for name in sorted(os.listdir(models_dir)):
        if not name.lower().endswith(".cs"):
            continue
        path = os.path.join(models_dir, name)
        stream = codecs.open(path, "r", "utf-8-sig")
        try:
            lines = stream.read().splitlines()
        finally:
            stream.close()

        attrs = []
        docs = []
        current = None
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("///"):
                docs.append(stripped)
                continue
            if stripped.startswith("["):
                attrs.append(stripped)
                continue
            if stripped.startswith("//") or stripped.startswith("#"):
                continue
            if stripped.startswith("using ") or stripped.startswith("namespace"):
                attrs, docs = [], []
                continue

            class_match = CLASS_RE.match(line)
            if class_match:
                table_name = None
                for attr in attrs:
                    found = TABLE_RE.search(attr)
                    if found:
                        table_name = found.group(1)
                entry = None
                if table_name:
                    entry = {
                        "table": table_name,
                        "class": class_match.group(1),
                        "file": name,
                        "summary": clean_doc(docs),
                        "indexes": [item for item in attrs if INDEX_RE.search(item)],
                        "columns": [],
                    }
                    tables.append(entry)
                current = entry
                attrs, docs = [], []
                continue

            property_match = PROPERTY_RE.match(line)
            if property_match and current is not None:
                csharp_type = property_match.group(1)
                prop_name = property_match.group(2)
                column_text = None
                for attr in attrs:
                    found = COLUMN_RE.search(attr)
                    if found:
                        column_text = found.group(1)
                pairs = {}
                if column_text:
                    for key, value in ATTR_PAIR_RE.findall(column_text):
                        pairs[key.strip()] = value.strip()
                ignored = pairs.get("IsIgnore", "false").lower() == "true"
                if not ignored:
                    is_nullable = pairs.get("IsNullable", "false").lower() == "true"
                    if csharp_type.endswith("?"):
                        is_nullable = True
                    string_length = compat.to_int(pairs.get("StringLength"), 0)
                    current["columns"].append(
                        {
                            "name": prop_name,
                            "csharp": csharp_type,
                            "sql": map_sql_type(csharp_type, string_length, is_nullable),
                            "nullable": is_nullable,
                            "primary": pairs.get("IsPrimary", "false").lower() == "true",
                            "identity": pairs.get("IsIdentity", "false").lower() == "true",
                            "summary": clean_doc(docs),
                        }
                    )
            attrs, docs = [], []
    return tables


def write_doc(tables, target):
    lines = [
        "# 表结构清单（主程序模型 → 本客户端数据契约）",
        "",
        "> 本文件由 `tools/extract_schema.py` 从主程序 `Models/*.cs` 自动生成，**请勿手改**。",
        "> 主程序改了模型后重新运行：`python tools/extract_schema.py`",
        "",
        "- 表数量：%d" % len(tables),
        "- 列数量：%d" % sum(len(item["columns"]) for item in tables),
        "",
    ]
    for entry in tables:
        lines.append("## `%s`" % entry["table"])
        lines.append("")
        if entry["summary"]:
            lines.append(entry["summary"])
            lines.append("")
        lines.append("来源：`Models/%s` → `%s`" % (entry["file"], entry["class"]))
        if entry["indexes"]:
            lines.append("")
            lines.append("索引：%s" % "；".join(item.strip() for item in entry["indexes"]))
        lines.append("")
        lines.append("| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |")
        lines.append("| --- | --- | --- | --- | --- | --- |")
        for column in entry["columns"]:
            markers = []
            if column["primary"]:
                markers.append("是")
            if column["identity"]:
                markers.append("自增")
            lines.append(
                "| `%s` | %s | %s | %s | %s | %s |"
                % (
                    column["name"],
                    column["csharp"],
                    column["sql"],
                    "是" if column["nullable"] else "否",
                    "、".join(markers) if markers else "",
                    column["summary"],
                )
            )
        lines.append("")
    compat.ensure_dir(os.path.dirname(target))
    compat.write_text(target, "\n".join(lines) + "\n")


def write_module(tables, target):
    lines = [
        "# -*- coding: utf-8 -*-",
        '"""表 → 列映射（由 tools/extract_schema.py 从主程序 Models/*.cs 生成，请勿手改）。',
        "",
        "重新生成：``python tools/extract_schema.py``",
        '"""',
        "",
        "TABLES = {",
    ]
    for entry in tables:
        lines.append('    "%s": (' % entry["table"])
        for column in entry["columns"]:
            lines.append('        ("%s", "%s", %s),' % (column["name"], column["sql"], column["nullable"]))
        lines.append("    ),")
    lines.append("}")
    lines += [
        "",
        "TABLE_NAMES = tuple(sorted(TABLES))",
        "",
        "",
        "def columns_of(table):",
        '    """返回列名元组；表不存在时返回空元组。"""',
        "    return tuple(item[0] for item in TABLES.get(table, ()))",
        "",
        "",
        "def sql_type_of(table, column):",
        "    for name, sql_type, _nullable in TABLES.get(table, ()):",
        "        if name == column:",
        "            return sql_type",
        "    return None",
        "",
    ]
    compat.ensure_dir(os.path.dirname(target))
    compat.write_text(target, "\n".join(lines))


def compare_with_database(tables, db_path):
    """与真实库比对（只读打开），返回差异描述列表。"""
    import sqlite3

    if not os.path.isfile(db_path):
        return ["数据库不存在：%s" % db_path]
    uri = "file:" + db_path.replace("\\", "/") + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True)
    try:
        cursor = connection.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY name")
        actual_tables = set(row[0] for row in cursor.fetchall() if not row[0].startswith("sqlite_"))
        messages = []
        for entry in tables:
            table = entry["table"]
            if table not in actual_tables:
                messages.append("库中尚无表 %s（主程序按需创建，或尚未用到）" % table)
                continue
            cursor.execute('PRAGMA table_info("%s")' % table)
            actual_columns = [row[1] for row in cursor.fetchall()]
            expected = [column["name"] for column in entry["columns"]]
            missing = [name for name in expected if name not in actual_columns]
            if missing:
                messages.append("表 %s 缺少列：%s" % (table, ", ".join(missing)))
            extra = [name for name in actual_columns if name not in expected]
            if extra:
                messages.append("表 %s 有模型未定义的列：%s" % (table, ", ".join(extra)))
        only_in_db = sorted(actual_tables - set(entry["table"] for entry in tables))
        if only_in_db:
            messages.append("仅存在于库中的表：%s" % ", ".join(only_in_db))
        return messages or ["与数据库一致"]
    finally:
        connection.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description="从主程序模型生成表结构清单")
    parser.add_argument("--db", help="顺便与真实数据库比对（只读）")
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    tables = parse_models(MODELS_DIR)
    write_doc(tables, SCHEMA_DOC)
    write_module(tables, SCHEMA_MODULE)

    if not args.quiet:
        compat.say("已生成：%s" % os.path.relpath(SCHEMA_DOC, _PARENT))
        compat.say("已生成：%s" % os.path.relpath(SCHEMA_MODULE, _PARENT))
        compat.say("表 %d 张，列 %d 个" % (len(tables), sum(len(item["columns"]) for item in tables)))
        for entry in tables:
            compat.say("  %-22s %2d 列  (%s)" % (entry["table"], len(entry["columns"]), entry["class"]))

    if args.db:
        compat.say("")
        compat.say("与数据库比对：%s" % args.db)
        for message in compare_with_database(tables, args.db):
            compat.say("  %s" % message)
    return 0


if __name__ == "__main__":
    sys.exit(main())
