# 打包说明（XP 发布包）

## 为什么只能在 Python 3.4 下打包

PyInstaller 的 Windows 启动器在 3.3 版本时把最低目标系统提到了 Vista 以上，
**3.3.1 才重新支持 XP 目标**；现代版本（4.x/5.x/6.x）明确要求 Windows 8+。
XP 上可用的最后一个 CPython 是 3.4.10（3.5 起官方不再支持 XP）。

因此：**Python 3.4.10（32 位）+ PyInstaller 3.3.1** 是唯一组合。

## 构建

```powershell
# 一次性准备（Python 3.4 环境里）
D:\Python34-32\python.exe -m pip install "pyinstaller==3.3.1"

# 打包
.\tools\build_xp.ps1 -Python D:\Python34-32\python.exe
```

产物：`xp-client\dist\ORT-XP\`（onedir），其中 `ORT-XP.exe` 是入口。

## 部署到 XP 机器

1. 把整个 `ORT-XP` 目录拷到 XP 机器（例如 `C:\ORT-XP`）。
2. 双击 `ORT-XP.exe`；首次运行会提示选择数据文件夹（应与主程序一致）。
   也可以在 `ORT-XP\Data\local_settings.json` 里预置：

   ```json
   { "DataFolder": "Z:\\ORT数据", "AteDataPath": null, "EmiDataPath": null }
   ```

   或设环境变量 `ORT_XP_DATA_FOLDER`。
3. 共享目录必须映射成盘符（SQLite 打不开 `\\服务器\共享\...`）。
4. 若提示缺少 `msvcr100.dll`：把 Python 3.4 安装目录下的 `msvcr100.dll`
   拷进 `ORT-XP` 目录（正常情况下 PyInstaller 会自动带上）。

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
