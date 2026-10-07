# -*- coding: utf-8 -*-
"""ORT 实验室管理系统 · XP 精简客户端。

主程序（.NET/WPF）的子项目，两者共用同一份 SQLite 库与同一套设置键。
本包内所有代码必须保持 Python 3.4 语法（XP 上可用的最后一个 CPython）。
"""

from .version import APP_NAME, APP_FOLDER_NAME, VERSION, BUILD_STAGE

__all__ = ["APP_NAME", "APP_FOLDER_NAME", "VERSION", "BUILD_STAGE"]
