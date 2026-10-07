# -*- coding: utf-8 -*-
"""Python 3.4 语法/API 下限检查（AST 扫描，零第三方依赖）。

XP 目标机只能用 Python 3.4，而开发机是 3.13：靠人记“这个写法 3.4 有没有”不可靠，
所以用本脚本把下限卡住。它自己也要能在 3.4 下运行。

用法::

    python tools/compat_check.py                  # 默认检查 ort_xp
    python tools/compat_check.py ort_xp tests     # 指定目录或文件
    python tools/compat_check.py --quiet          # 只输出汇总

退出码：0 = 通过；1 = 发现问题。
"""

import ast
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)

from ort_xp import compat  # noqa: E402

#: 3.4 上不存在的模块（模块名 → 引入版本）
FORBIDDEN_MODULES = {
    "typing": "3.5",
    "dataclasses": "3.7",
    "secrets": "3.6",
    "contextvars": "3.7",
    "zoneinfo": "3.9",
    "graphlib": "3.9",
    "tomllib": "3.11",
    "importlib.resources": "3.7",
}

#: 3.4 上不存在的“模块.属性”
FORBIDDEN_ATTRS = {
    "os.scandir": "3.5",
    "os.fspath": "3.5",
    "math.isclose": "3.5",
    "random.choices": "3.6",
    "time.time_ns": "3.7",
    "ssl.TLSVersion": "3.7",
}

#: 只要属性名出现就报（这些名字在任何对象上都是 3.5+ 才有的）
FORBIDDEN_ATTR_NAMES = {
    "removeprefix": "3.9",
    "removesuffix": "3.9",
    "fromisoformat": "3.7",
}

#: 关键字参数规则：(关键字, 适用的函数名集合, 引入版本, 说明)
KEYWORD_RULES = (
    ("capture_output", ("subprocess.run", "run"), "3.7", "subprocess.run(capture_output=)"),
    ("text", ("subprocess.run", "run"), "3.7", "subprocess.run(text=)"),
    (
        "encoding",
        (
            "logging.FileHandler",
            "FileHandler",
            "logging.handlers.RotatingFileHandler",
            "RotatingFileHandler",
            "logging.handlers.TimedRotatingFileHandler",
            "TimedRotatingFileHandler",
        ),
        "3.9",
        "日志处理器的 encoding= 参数（本项目改用 codecs.open 包装）",
    ),
)

_JoinedStr = getattr(ast, "JoinedStr", None)
_NamedExpr = getattr(ast, "NamedExpr", None)
_AnnAssign = getattr(ast, "AnnAssign", None)
_AsyncFunctionDef = getattr(ast, "AsyncFunctionDef", None)
_AsyncFor = getattr(ast, "AsyncFor", None)
_AsyncWith = getattr(ast, "AsyncWith", None)
_Await = getattr(ast, "Await", None)
_MatMult = getattr(ast, "MatMult", None)


def _dotted(node):
    """把属性访问还原成 ``a.b.c`` 形式的字符串（尽力而为）。"""
    parts = []
    current = node
    while isinstance(current, ast.Attribute):
        parts.append(current.attr)
        current = current.value
    if isinstance(current, ast.Name):
        parts.append(current.id)
        parts.reverse()
        return ".".join(parts)
    return ".".join(reversed(parts))


def _call_name(node):
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        return _dotted(node)
    return ""


def _check_node(node, findings, filename):
    def report(lineno, message):
        findings.append((filename, lineno or 0, message))

    if _JoinedStr is not None and isinstance(node, _JoinedStr):
        report(getattr(node, "lineno", 0), "f-string 是 3.6+，请改用 \"{}\".format() 或 %")
    if _NamedExpr is not None and isinstance(node, _NamedExpr):
        report(getattr(node, "lineno", 0), "海象运算符 := 是 3.8+")
    if _AnnAssign is not None and isinstance(node, _AnnAssign):
        report(node.lineno, "变量注解（x: int = ...）是 3.6+")
    if _AsyncFunctionDef is not None and isinstance(node, _AsyncFunctionDef):
        report(node.lineno, "async def 是 3.5+")
    if _AsyncFor is not None and isinstance(node, _AsyncFor):
        report(node.lineno, "async for 是 3.5+")
    if _AsyncWith is not None and isinstance(node, _AsyncWith):
        report(node.lineno, "async with 是 3.5+")
    if _Await is not None and isinstance(node, _Await):
        report(node.lineno, "await 是 3.5+")
    if _MatMult is not None and isinstance(node, _MatMult):
        report(getattr(node, "lineno", 0), "@ 矩阵乘是 3.5+")

    if isinstance(node, ast.Import):
        for alias in node.names:
            version = FORBIDDEN_MODULES.get(alias.name)
            if version:
                report(node.lineno, "import %s（%s）" % (alias.name, version))
    elif isinstance(node, ast.ImportFrom):
        module = node.module or ""
        version = FORBIDDEN_MODULES.get(module)
        if version:
            report(node.lineno, "from %s import ...（%s）" % (module, version))

    if isinstance(node, ast.Attribute):
        dotted = _dotted(node)
        version = FORBIDDEN_ATTRS.get(dotted)
        if version:
            report(node.lineno, "%s 是 %s 引入的" % (dotted, version))
        version = FORBIDDEN_ATTR_NAMES.get(node.attr)
        if version:
            report(node.lineno, ".%s() 是 %s 引入的" % (node.attr, version))

    if isinstance(node, ast.Call):
        name = _call_name(node.func)
        for keyword, names, version, description in KEYWORD_RULES:
            if name not in names:
                continue
            for item in node.keywords:
                if item.arg == keyword:
                    report(getattr(item.value, "lineno", node.lineno), "%s 是 %s 引入的" % (description, version))


def check_source(source, filename):
    """检查一段源码，返回 ``[(文件, 行号, 说明)]``。"""
    findings = []
    try:
        tree = ast.parse(source, filename)
    except SyntaxError as exc:
        return [(filename, exc.lineno or 0, "语法错误：%s" % exc.msg)]
    for node in ast.walk(tree):
        _check_node(node, findings, filename)
    return findings


def check_file(path):
    text = compat.read_text(path)
    if text is None:
        return [(path, 0, "无法读取")]
    return check_source(text, path)


def check_path(path):
    findings = []
    if os.path.isdir(path):
        files = sorted(compat.walk_files(path, ".py"))
        for item in files:
            findings.extend(check_file(item))
        return findings
    if os.path.isfile(path):
        return check_file(path)
    return [(path, 0, "路径不存在")]


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    quiet = False
    if "--quiet" in argv:
        quiet = True
        argv.remove("--quiet")
    targets = argv or [os.path.join(_PARENT, "ort_xp")]

    total = 0
    for target in targets:
        findings = check_path(target)
        for filename, lineno, message in findings:
            if not quiet:
                compat.say("%s:%d: %s" % (os.path.relpath(filename, _PARENT), lineno, message))
        total += len(findings)
        if not quiet and not findings:
            compat.say("%s：通过（无 3.4 不兼容写法）" % target)
    compat.say("Python 3.4 语法下限检查：%d 处问题" % total)
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
