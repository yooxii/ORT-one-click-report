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
2. 双击 `ORT-XP.exe`；首次运行会提示选择数据文件夹（应与主程序一致）。
   也可以在 `ORT-XP\Data\local_settings.json` 里预置：

   ```json
   { "DataFolder": "Z:\\ORT数据", "AteDataPath": null, "EmiDataPath": null }
   ```

   或设环境变量 `ORT_XP_DATA_FOLDER`。
3. 共享目录必须映射成盘符（SQLite 打不开 `\\服务器\共享\...`）。
4. `msvcr100.dll`（32 位 VC++2010 运行库）**已经打进包里**（约 774 KB，Microsoft 签名），
   XP 机器上不需要另装运行库。
5. 排错：窗口子系统看不到控制台，异常会写进 `ORT-XP\Logs\fatal_<时间戳>.log`（UTF-8）并弹框提示；
   普通运行日志在同目录 `ort_xp.log`。

## 已在本机验证过的内容（2026-10-07）

```
Python 3.4.4 (32 bit) + PyInstaller 3.3.1  →  dist\ORT-XP（15.3 MB，917 个文件）
ORT-XP.exe 的 PE 头 MajorOperatingSystemVersion = 5.1   ← 启动器面向 Windows XP
ORT-XP.exe --version         → 退出码 0，中文输出正常
ORT-XP.exe --selftest        → 退出码 0，14 项检查全通过（含 DPAPI 往返、SQLite 3.8.11、
                               OpenSSL 1.0.2d + PROTOCOL_TLSv1_2、真实库 20 张表）
ORT-XP.exe --check-plans-email → 退出码 0（演练，68 条候选计划，未发信）
在含简体中文的路径下运行 → 退出码 2，并提示「请把程序放到纯英文/数字路径」
```

仍然待办：在**真实 XP 机器（或 XP 虚拟机）**上跑一遍（含界面登录与编辑操作）。

## 已知打包注意事项

| 事项 | 说明 |
| --- | --- |
| 用 onedir，不要 onefile | onefile 每次启动解压到 `%TEMP%`，XP 上慢且易被杀软误报 |
| 不要开 UPX | XP 时代杀软对 UPX 壳敏感，spec 里已关闭 |
| 控制台 | `console=False`（GUI 程序）；排错时临时改成 `True` 可看到日志 |
| 证书包 | 需要校验证书时把 `cacert.pem` 放到 exe 同级（`ort_xp/context.py` 会优先找它） |
| 体积 | 约 15–25 MB（仅标准库 + tkinter） |

## 排错顺序

1. XP 上先跑 `ORT-XP.exe --selftest`（在命令行里）看自检输出；
   若没有界面或闪退，日志在 `ORT-XP\Logs\ort_xp.log`（UTF-8）。
2. 数据层问题看 `docs/02-数据契约.md`（路径解析、日期格式、WAL/共享目录）。
3. 打包器本身跑不起来时的替代方案见 `docs/01-环境搭建.md` 第 6 节。
