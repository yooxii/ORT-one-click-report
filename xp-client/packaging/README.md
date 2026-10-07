# 打包说明（XP 发布包）

## 为什么只能在 Python 3.4 下打包

PyInstaller 的 Windows 启动器在 3.3 版本时把最低目标系统提到了 Vista 以上，
**3.3.1 才重新支持 XP 目标**；现代版本（4.x/5.x/6.x）明确要求 Windows 8+。
XP 上可用的最后一个 CPython 系列是 **3.4**（3.5 起改用 VS2015 的 UCRT，不支持 XP）。
注意 `python-3.4.10.exe` 并不存在：3.4.5 之后 python.org 只发源码包，
带 Windows 安装包的最后一版是 **3.4.4**（`python-3.4.4.msi`），本项目的打包机就用它。

因此：**Python 3.4.4（32 位）+ PyInstaller 3.3.1** 是唯一组合。

## 构建

```powershell
# 一次性准备（Python 3.4 环境里）
D:\Python34-32\python.exe -m pip install "pyinstaller==3.3.1"

# 打包
.\tools\build_xp.ps1 -Python D:\Python34-32\python.exe
```

产物：`xp-client\dist\ORT-XP\`（onedir），其中 `ORT-XP.exe` 是入口。

## 部署到 XP 机器

> **硬性要求：程序目录必须是纯 ASCII 路径（例如 `C:\ORT-XP`）。**
> Windows 下模块加载器用 **ANSI 代码页**编码路径：路径里出现该代码页表示不了的字符
> （例如本机 cp950 繁体环境下，仓库名里的简体「一键报告」），导入 tkinter 时就会以
> `UnicodeEncodeError` 崩掉。程序启动时会先自检，遇到这种路径会弹框提示并要求换路径
> （`--selftest` 时以退出码 2 结束）。

1. 把整个 `ORT-XP` 目录拷到 XP 机器，例如 `C:\ORT-XP`。
2. 双击 `ORT-XP.exe`；**首次运行会弹框让你选数据文件夹**（应与主程序一致）。
   选中的目录（或它的 `Data` 子目录）里必须有 `ort_plans.db`，选完会写进
   `ORT-XP\Data\local_settings.json`，下次启动直接生效。
   也可以预先放好这个文件：

   ```json
   { "DataFolder": "Z:\\ORT数据", "AteDataPath": null, "EmiDataPath": null }
   ```

   或设环境变量 `ORT_XP_DATA_FOLDER`，或在快捷方式里加 `--data-folder "Z:\ORT数据"`（跳过弹框）。
3. 共享目录必须映射成盘符（SQLite 打不开 `\\服务器\共享\...`）。
4. `msvcr100.dll`（32 位 VC++2010 运行库）**已经打进包里**（约 774 KB，Microsoft 签名），
   XP 机器上不需要另装运行库。
5. 排错：窗口子系统看不到控制台，**所有启动期失败都会写日志并弹框**。
   - 运行日志：`ORT-XP\Logs\ort_xp.log`（UTF-8；`--selftest` / `--ui-smoke` 的结果也写在这里）；
   - 未捕获异常的完整回溯：`ORT-XP\Logs\fatal_<时间戳>.log`；
   - 不想让弹框打断脚本时加 `--no-dialog`（或设 `ORT_XP_NO_DIALOG=1`）。

## 已在本机验证过的内容（2026-10-07）

```
Python 3.4.4 (32 bit) + PyInstaller 3.3.1  →  dist\ORT-XP（15.3 MB，917 个文件）
ORT-XP.exe 的 PE 头 MajorOperatingSystemVersion = 5.1   ← 启动器面向 Windows XP
Python 3.4.4 下 95 项单测全通过（ORT_XP_UI_TEST=1 时含界面装配与可见性检查）
ORT-XP.exe --version         → 退出码 0，中文输出正常
ORT-XP.exe --selftest        → 退出码 0，14 项检查全通过（含 DPAPI 往返、SQLite 3.8.11、
                               OpenSSL 1.0.2d + PROTOCOL_TLSv1_2、真实库 20 张表）
ORT-XP.exe --check-plans-email → 退出码 0（演练，68 条候选计划，未发信）
ORT-XP.exe --ui-smoke --data-folder "..."  → 退出码 0，7 个窗口逐个 winfo_viewable() 通过
无数据文件夹时双击           → 不再静默退出：日志写 [ERROR]，并弹出「没有找到数据库」对话框
                               （列出数据文件夹、库路径、4 种解决办法，可直接选文件夹）
有数据文件夹时双击           → 登录窗口正常显示（已用 PrintWindow 截图确认窗口内容）
在含简体中文的路径下运行     → 退出码 2，并提示「请把程序放到纯英文/数字路径」
```

仍然待办：在**真实 XP 机器（或 XP 虚拟机）**上跑一遍（含界面登录与编辑操作）。

## 与主程序的已知差异

| 差异 | 说明 |
| --- | --- |
| 收件人分隔符 | XP 端比主程序多认全角逗号「，」、顿号「、」与斜杠「/」。真实库里「负责人」有 187 条写成 `李剛/李志斌`，主程序不拆斜杠 → 这些提醒在两端都发不出去；XP 端拆开后能解析出两位负责人。要完全一致的话主程序也要加这个分隔符 |
| 邮件口令 | XP 端**不写** `mail.passwordEnc`（那是主程序用**它那台机器**的 DPAPI 加密的，XP 端写进去主程序解不开），口令存在 XP 机器本机（DPAPI）。因此 XP 端发信要用本机口令 |
| 邮件模板与抄送管理员开关 | `mail.template.*` / `mail.ccAdmin.*` 由主程序设置窗口维护，XP 端只读不改 |
| 其他设置项 | XP 端设置界面只写 SMTP 相关的那 17 个 `mail.*` 键（多键一次事务），其余设置键不动 |
| 授权/审核流程 | 不在范围内：XP 端没有审核提交，改动直接落库并写变更日志（主程序会看到这些改动） |

## 已知打包注意事项

| 事项 | 说明 |
| --- | --- |
| 用 onedir，不要 onefile | onefile 每次启动解压到 `%TEMP%`，XP 上慢且易被杀软误报 |
| 不要开 UPX | XP 时代杀软对 UPX 壳敏感，spec 里已关闭 |
| 控制台 | `console=False`（GUI 程序）；排错时临时改成 `True` 可看到日志 |
| 证书包 | 需要校验证书时把 `cacert.pem` 放到 exe 同级（`ort_xp/context.py` 会优先找它） |
| 体积 | 约 15–25 MB（仅标准库 + tkinter） |

## 排错顺序

1. XP 上先在**命令行**里跑 `ORT-XP.exe --selftest`：
   - 从命令行启动时能继承控制台，输出直接看得见；双击启动时输出会写进 `ORT-XP\Logs\ort_xp.log`。
   - 界面起不起来用 `ORT-XP.exe --ui-smoke --data-folder "..."`（造窗口后立刻退出，退出码 0 = 通过）。
2. 「双击没反应 / 只写了两行日志」：先看 `Logs\ort_xp.log` 有没有 `[ERROR]` 行 ——
   没有数据库、路径含当前代码页表示不了的字符、tkinter 装不起来，这三种都会留下 `[ERROR]` 并弹框。
   2026-10-07 就是把「`Data` 目录不存在 → 静默退出 2」这条补成可见的。
   如果日志显示数据库已连上、界面却一个窗口都没有（进程活着、响应中），那是
   **登录窗口 transient 到已 withdraw 的主窗口后被 Tk 一起隐藏**了（Tk 8.6 在 Windows 上的行为），
   已在 `LoginDialog` 里去掉 transient 修复；`--ui-smoke` 会逐个窗口检查 `winfo_viewable()`。
3. 数据层问题看 `docs/02-数据契约.md`（路径解析、日期格式、WAL/共享目录）。
4. 打包器本身跑不起来时的替代方案见 `docs/01-环境搭建.md` 第 6 节。
