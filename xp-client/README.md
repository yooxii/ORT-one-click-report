# ORT 实验室管理系统 · XP 精简客户端（xp-client）

主程序（仓库根目录的 .NET / WPF 版「ORT实验室管理系统」）的**子项目**，给只能跑 Windows XP SP3 的
机器用的精简客户端。两者**共用同一份数据**（同一个 SQLite 库、同一套设置键），不是替代关系。

## 为什么用 Python

| 约束 | 影响 |
| --- | --- |
| XP 上最高只能装 .NET Framework 4.0.3 | 主程序（net48）无法在 XP 运行 |
| XP 的 Schannel 只支持到 TLS 1.0 | .NET 4.0 发不出 TLS 1.2 的邮件 → **邮件需求决定了必须换技术栈** |
| Python 自带 OpenSSL，不走系统 Schannel | XP 上照样能 TLS 1.2 发信 |
| 本范围不需要 Excel/Word/PDF 处理 | **零第三方运行时依赖**，只用标准库：`sqlite3` / `tkinter` / `smtplib` / `ssl` / `hashlib` / `ctypes` |

## 范围

**做**：数据层（读写主程序的 `ort_plans.db`）、登录、领用表与计划表的简单交互、邮件（SMTP + 计划到期提醒）。

**不做**：报告生成（Excel/Word/PDF/OLE）、审核、流程窗口、管理端、多语言与主题体系、计划索引与报告扫描、
列筛选/列宽记忆/拖拽等交互细节。完整清单与验收标准见 [docs/00-范围与验收.md](docs/00-范围与验收.md)。

## 目标环境

| | 版本 | 说明 |
| --- | --- | --- |
| 运行 | Windows XP SP3（32 位）+ Python 3.4（32 位）+ PyInstaller 3.3.1（onedir） | XP 可用的最后一组组合 |
| 开发 | 本机 Python 3.13.3（32 位）+ 源码保持 3.4 语法 | 用 `tools/compat_check.py` 卡住语法下限 |
| 打包 | 必须在 Python 3.4 环境执行 | PyInstaller 3.3.1–3.4 才能生成 XP 可用的启动器 |

细节、安装步骤与替代方案见 [docs/01-环境搭建.md](docs/01-环境搭建.md)。

## 目录结构

```
xp-client/
├─ README.md                    # 本文件
├─ docs/                        # 设计与接口文档
│  ├─ 00-范围与验收.md           # 做什么 / 不做什么 / 怎么算做完
│  ├─ 01-环境搭建.md             # 目标机与开发机环境、安装步骤、语法约束
│  ├─ 02-数据契约.md             # 与主程序共库的全部约定（路径、设置键、日期格式、并发）
│  ├─ 03-里程碑与工时.md         # 阶段拆分、工时、当前进度
│  └─ schema.md                 # 表与字段清单（由 tools/extract_schema.py 从 C# 模型生成，勿手改）
├─ ort_xp/                      # 应用源码（Python 3.4 兼容）
│  ├─ compat.py                 # 3.4 兼容层：编码、控制台、.NET 值解析
│  ├─ config.py                 # 数据目录 / 数据库路径 / app_settings 读取
│  ├─ dpapi.py                  # DPAPI（ctypes，零依赖）加解密
│  ├─ logging_setup.py          # 日志（UTF-8 文件 + 控制台安全输出）
│  ├─ db/                       # 数据层：连接、事务、重试、仓储、表结构
│  ├─ services/                 # auth（登录）/ mail（SMTP 与到期提醒）
│  └─ ui/                       # tkinter 界面：登录、主窗口、领用与计划
├─ tools/                       # 开发/运维脚本（不随程序发布）
│  ├─ check_env.py              # 环境自检（解释器、依赖、数据库连通性、DPAPI）
│  ├─ compat_check.py           # Python 3.4 语法/API 下限检查
│  ├─ extract_schema.py         # 从主程序 C# 模型生成 docs/schema.md 与 schema_generated.py
│  ├─ verify_roundtrip.py       # 共库往返验证（在真实库副本上读写，证明与主程序格式一致）
│  ├─ run_dev.ps1               # 开发机启动
│  └─ build_xp.ps1              # XP 包构建（需在 Python 3.4 环境执行）
├─ packaging/                   # PyInstaller 配置与打包说明
├─ requirements/                # 依赖清单（运行时为零依赖）
└─ tests/                       # 标准库 unittest，无第三方测试框架
```

## 快速开始

```powershell
# 1) 环境自检（解释器、依赖、语法下限、数据库连通性、DPAPI）
python .\tools\check_env.py

# 2) 单元测试
python -m unittest discover -s tests -v

# 3) 共库往返验证（在真实库的副本上读写，证明两边格式一致且不动原库）
python .\tools\verify_roundtrip.py

# 4) 启动界面（开发机；用 --data-folder 指向主程序数据目录）
.\tools\run_dev.ps1 -DataFolder "D:\source\repos\ORT一键报告\bin\Debug\Data"

# 5) 无界面自检（不连界面，验证数据层与邮件配置）
python -m ort_xp --selftest
```

## 与主程序的约定（重要）

1. **表结构由主程序负责**（FreeSql `UseAutoSyncStructure`）——XP 客户端只读写，**不建表、不改表**。
2. 每次改业务数据必须写 `plan_change_logs`，否则主程序的审核模块看不到 XP 这边的改动。
3. 日期一律按主程序格式读写：`yyyy-MM-dd HH:mm:ss`，含微秒时为 `yyyy-MM-dd HH:mm:ss.ffffff`。
4. 设置走 `app_settings` 键值对，布尔值文本是 .NET 风格 `True` / `False`。
5. 邮件密码在主程序里是 DPAPI(CurrentUser) 加密后存库的，**换机器/换用户解不开** → XP 客户端自带本机口令文件。
6. 两套客户端可能同时打开同一个库 → 短事务 + 忙等待重试，不主动改 `journal_mode`。

全部细节见 [docs/02-数据契约.md](docs/02-数据契约.md)。

## 当前进度

- [x] 里程碑 0：子项目骨架、环境自检、数据层连通性验证（Python 3.13 直读主程序库成功）
- [x] 里程碑 2：数据层（仓储 / 变更日志 / 下拉取值）与登录、记住登录
- [x] 里程碑 3：领用表与计划表的新增/编辑/删除（校验规则与自动编号对齐主程序，改动写变更日志）；
  61 项单测 + 真实库副本 32 项检查全通过
- [ ] 里程碑 1：技术验证 —— **打包与本机运行已完成**（`D:\Python34-32` 装好 Python 3.4.4 + PyInstaller 3.3.1，
  `dist\ORT-XP` 15.3 MB 可运行，3.4 下 66 项单测全绿）；**待办：在真实 XP 机器/虚拟机上跑一遍界面**
- [ ] 里程碑 3 收尾：界面在 XP 实机上的点击回归、列表点列头排序
- [ ] 里程碑 4：邮件（SMTP 发送与到期提醒服务已就绪，设置界面写库待做）
- [ ] 里程碑 5：打包、XP 回归、交付

工时估算见 [docs/03-里程碑与工时.md](docs/03-里程碑与工时.md)。
