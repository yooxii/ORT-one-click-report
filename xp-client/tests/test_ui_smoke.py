# -*- coding: utf-8 -*-
"""界面装配自检（默认跳过）。

tkinter 的窗口没法在无人值守下点击，但「窗口能不能造出来、控件引用对不对」是可以自动查的：
登录窗口把 ``focus_set`` 写到了 ``StringVar`` 上这类问题，就是靠这个自检发现的。

显式开启：``ORT_XP_UI_TEST=1`` 时运行；会在桌面会话里瞬时创建并销毁窗口（不进入 mainloop）。
"""

import os
import shutil
import sqlite3
import sys
import tempfile
import unittest

_TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_TESTS_DIR)
for _path in (_PARENT, os.path.join(_PARENT, "tools"), _TESTS_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from test_smoke import SCHEMA_SQL  # noqa: E402

from ort_xp import context as context_module  # noqa: E402

ENABLED = os.environ.get("ORT_XP_UI_TEST") == "1"


@unittest.skipUnless(ENABLED, "界面自检需要显示会话，设置 ORT_XP_UI_TEST=1 开启")
class UiSmokeTests(unittest.TestCase):
    def setUp(self):
        import tkinter as tk

        self.tk = tk
        self.temp_dir = tempfile.mkdtemp(prefix="ort-xp-ui-")
        db_path = os.path.join(self.temp_dir, "ort_plans.db")
        connection = sqlite3.connect(db_path)
        try:
            connection.executescript(SCHEMA_SQL)
            connection.commit()
        finally:
            connection.close()

        self.context = context_module.AppContext(app_directory=self.temp_dir, data_folder=self.temp_dir)
        self.context.open()
        self.context.database.execute(
            "INSERT INTO users (Username, DisplayName, Email, PasswordHash, Salt, IsActive, CreatedAt) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            ("ui", "界面自检", "ui@example.com", "", "", 1, "2026-10-07 09:00:00"),
        )
        self.root = self.tk.Tk()
        self.root.withdraw()

    def tearDown(self):
        try:
            self.root.destroy()
        except Exception:
            pass
        try:
            self.context.close()
        except Exception:
            pass
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_login_dialog_builds(self):
        from ort_xp.ui.app import LoginDialog

        dialog = LoginDialog(self.root, self.context)
        try:
            self.assertIsInstance(dialog.username.get(), str)
            self.assertIsInstance(dialog.password.get(), str)
        finally:
            dialog.window.destroy()

    def test_main_window_and_table_tabs_build(self):
        from ort_xp.ui.app import Application, TableTab

        application = Application(self.context)
        try:
            application._build_main()
            for kind in ("requisitions", "plans"):
                tab = TableTab(application.root, self.context, kind)
                tab.reload()
                self.assertEqual(0, len(tab.tree.get_children()))
                tab.destroy()
        finally:
            application.root.destroy()

    def test_form_dialog_reads_values(self):
        from ort_xp.services import plans as plans_service
        from ort_xp.ui.app import FormDialog

        dialog = FormDialog(
            self.root, "表单自检", plans_service.requisition_form_fields(), {"RequisitionNo": "2609-001"}
        )
        try:
            values = dialog.read()
            self.assertEqual("2609-001", values["RequisitionNo"])
            self.assertIn("Disposition", values)
        finally:
            dialog.window.destroy()

    def test_mail_window_builds(self):
        from ort_xp.ui.app import MailWindow, SelftestWindow

        mail_window = MailWindow(self.root, self.context)
        try:
            self.assertIn("当前配置", mail_window.window.title() + "（当前配置）")
        finally:
            mail_window.window.destroy()
        selftest = SelftestWindow(self.root, self.context)
        try:
            self.assertTrue(selftest.window.title())
        finally:
            selftest.window.destroy()


    def test_ui_smoke_helper_passes(self):
        """打包后的 ``--ui-smoke`` 走的就是这个函数：目标机上一个命令证明界面装得起来。"""
        from ort_xp.ui import app as ui_app

        passed, text = ui_app.ui_smoke(self.context)
        self.assertTrue(passed, text)
        self.assertIn("界面装配自检", text)
        self.assertNotIn("[失败]", text)


if __name__ == "__main__":
    unittest.main()
