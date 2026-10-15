# 架构证据索引 — ORT实验室管理系统

> 配套 `system-model.structurizr.dsl`（结构真源）与 `business-flow.dot`（业务主路径）。
> 逆向自代码，程序集版本 `0.6.2.0`（`Properties/AssemblyInfo.cs:54`），生成日期 2026-10-15。
>
> **置信度口径**
> - `high` — 代码、配置、schema 或权威文档直接证明
> - `medium` — 多个部分信号一致，但没有单一来源直接确认该关系
> - `low` — 由命名、目录结构或约定推断
> - `unknown` — 刻意留空的待验证区域
>
> **状态口径**：本套模型全部为 `current`（当前状态），无 target/proposed 节点。

---

## 1. 节点证据

### 1.1 使用者（L1 角色）

| ID | 标签 | 类型 | 置信度 | 证据 |
| --- | --- | --- | --- | --- |
| `guest` | 游客 | actor | high | `Services/PermissionService.cs:19-21`（未登录视为游客）、`App.xaml.cs:192-201`（`login.IsGuestMode` 以 `CurrentUser=null` 进入）、`README.md:62-63` |
| `generalUser` | 普通用户 | actor | high | `Services/PermissionService.cs:38,58-62,83-86`、`Models/User.cs:74`（`user_roles` 表） |
| `technician` | 技术员 | actor | high | `Services/PermissionService.cs:38,66,78-80,94-96` |
| `reviewer` | 审核员 | actor | high | `Services/PermissionService.cs:39,67,89-91,97-98` |
| `admin` | 管理员 | actor | high | `Services/PermissionService.cs:40,68,101-105`、`Admin/Views/WindowAdmin.xaml.cs:590-603`（清空数据需两次确认） |

**权限矩阵逐格证据**：`Services/PermissionService.cs:64-107`（已逐行核对）。其中 `plan.view` 恒 true 见 `:74-75`；`report.use` 仅技术员与管理员见 `:94-96`；default 分支 `return isAdmin` 见 `:104-105`。

### 1.2 外部系统

| ID | 标签 | 类型 | 置信度 | 证据 |
| --- | --- | --- | --- | --- |
| `smtp` | SMTP 邮件服务器 | external-system | high | `Services/SmtpMailClient.cs:19`（自实现，明文/STARTTLS/隐式 SSL）、`Services/MailService.cs:11`、`README.md:118` |
| `office` | Microsoft Office | external-system | high | **Excel = 运行时后期绑定**：`Utils/ExcelOleEmbedder.cs:74,110-111`（注释明示"不依赖 Interop PIA / office.dll，只要求装了 Excel"）、`:111`（`Type.GetTypeFromProgID("Excel.Application")`）、`:121-132`（`dynamic` + `Activator.CreateInstance`）、`:433,447-453`（dynamic 封送失败与 COM 释放兜底）；全仓检索 `Microsoft.Office.Interop.Excel` **零命中**。**Word = 编译期 COM 引用**：`ORT一键报告.csproj:637-645`（`Microsoft.Office.Interop.Word` PIA）、`Utils/Docx2Pdf.cs:1`。另有 `csproj:628-636`（`Microsoft.Office.Core`）与 `:646-654`（`VBIDE`） |
| `srcExcel` | 计划表 / 领用表 Excel 源文件 | external-system | high | `Services/PlanExcelService.cs:62,187,471`（表头前 10 行按去空白关键字定位）、`:900,910`（年份从文件名提取）、`:80,691`（S/N 列 OLE 提取）、`README.md:111-112` |
| `reportFolder` | 报告文件夹 | external-system | high | `Services/ReportScanService.cs:56`（路径取设置项 `ReportDir`）、`:88-98`（递归 4 层、目录名包含工作编号即匹配）、`:109-127`（要求 `Report` 子目录 + 概览 xls） |
| `winEnv` | Windows 客户端环境 | external-system | high | `Services/SingleInstanceManager.cs:11-25`（`Local\` 互斥体 + 命名管道）、`Services/StartupManager.cs:13-15`（HKCU Run）、`MainWindow.xaml.cs:44,241-249`（WinForms NotifyIcon 托盘）、`Services/LoginCredentialStore.cs:9-14`（DPAPI + `%LocalAppData%`）、`Services/WindowThemeHelper.cs`（DWM 深色标题栏）、`README.md:141`（远程路径禁用开机自启） |

### 1.3 容器（L2）

| ID | 标签 | 类型 | 置信度 | 证据 |
| --- | --- | --- | --- | --- |
| `wpf` | WPF 桌面客户端 | container | high | `ORT一键报告.csproj`（net48、输出 `ORT实验室管理系统.exe`）、`App.xaml.cs:41-259`（启动编排）、`:111-148`（DI 注册全部 Services 与 ViewModels）、`MainWindow.xaml.cs:217`（关窗转后台）、`:690-730`（注销时按权限关闭窗口） |
| `xp` | Windows XP 精简客户端 | container | high | `xp-client/README.md:3-4`（主程序的**子项目**，两者**共用同一份数据**——同一个 SQLite 库、同一套设置键，不是替代关系）、`:8-14`（换 Python 的动因：XP 最高 .NET 4.0.3、Schannel 只到 TLS 1.0，发不出 TLS 1.2 邮件；零第三方运行时依赖，只用标准库 `sqlite3`/`tkinter`/`smtplib`/`ssl`/`hashlib`/`ctypes`）、`:18-21`（**做**：数据层读写 `ort_plans.db`、登录、领用表与计划表简单交互、邮件到期提醒；**不做**：报告生成、审核、流程窗口、管理端、多语言与主题、计划索引与报告扫描）、`:25-29`（XP SP3 32 位 + Python 3.4 32 位 + PyInstaller 3.3.1 onedir）；跨容器契约文档 `xp-client/docs/02-数据契约.md` 与 `xp-client/docs/schema.md`；另有 `xp-client/docs/00-范围与验收.md`、`01-环境搭建.md`、`03-里程碑与工时.md`、`04-部署与交付.md`，代码在 `xp-client/ort_xp`，语法下限由 `xp-client/tools/compat_check.py` 卡住 |
| `db` | SQLite 主库 `ort_plans.db` | database | high | `Services/DatabaseService.cs:133,146-148`（库文件名与唯一主连接）、`:104-133`（UNC 映射、并发建议）、`:190-195`（本地 WAL / 网络 rollback journal）、`README.md:177`（`UseAutoSyncStructure(true)` 自动迁移） |
| `fileStore` | 附件·配图·备份文件区 | database（文件存储） | high | `Services/PlanExcelService.cs:16-19,691`（`OleFiles`）、`Services/DatabaseBackupService.cs:49,66,85-89`（`Backups`）、`README.md:170`（`OleFiles` / `PlanImages`）、`更新日志.md:67`（备份范围只含数据库本身，不含附件与配图） |
| `templates` | 报告模板目录 `Templates` | database（文件存储） | high | `Services/ReportTemplateService.cs:327-333`（按 `Templates` 下模板目录名定位）、`:281-333`（多根探测）、`Reports/Views/WindowMainReport.xaml.cs:93`、`Reports/ViewModels/BaseReportPageViewModel.cs:62`、`README.md:207,238`（随生成复制到输出目录） |

### 1.4 组件（L3，均在 `wpf` 容器内）

| ID | 标签 | 分组 | 置信度 | 证据 |
| --- | --- | --- | --- | --- |
| `shell` | 应用外壳与启动编排 | Platform | high | `App.xaml.cs:50-59`（单实例 + 监听重复启动）、`:65`（`OnExplicitShutdown`）、`:71,74`（语言/主题初始化）、`:79-101`（全窗口 Loaded 类处理器：深色标题栏、字体、开窗防闪）、`:108-109`（自定义本地化 Provider）、`:111-148`（DI）、`:153-173`（数据文件夹检查与回退）、`:179`（幂等修正公式脏数据）、`:192-201`（登录前置）、`:204-211`（手建主窗口）、`:206-207,219-223`（`--background` 进托盘）、`:235,245-247`（启动扫描与备份调度）、`:306-331`（退出清理）；`MainWindow.xaml.cs:617-620`（到期邮件定时器）、`:690-730`（注销按权限关窗）；`Services/SingleInstanceManager.cs`、`Services/StartupManager.cs` |
| `authz` | 认证与权限判定 | Platform | high | `Services/AuthService.cs:51-73`（首启建默认 admin）、`:98`（salt+hash 校验）、`:103-107`（空密码登录）、`:110-116`（读写 users/user_roles）、`:216-227`（7 天免登录）；`Services/PermissionService.cs:34-41,58-62,64-107`；`Services/LoginCredentialStore.cs:9-14` |
| `plans` | 领退与计划 | Business | high | `Plans/Views/WindowPlans.xaml(.cs)`、`Plans/ViewModels/PlansViewModel.cs:50,1666-1758`（暂存-提交、先删→再改→后增、逐条变更日志）、`:871-888`（领退↔计划三级匹配）、`Services/PlanExcelService.cs:62,187,100-107`（导入与多条件合并）、`:530-552`（公式脏数据改写 `#N/A`）、`Utils/PlanValidation.cs`、`Utils/OrtPlanParser.cs`、`Plans/ViewModels/FlowViewModel.cs`（流程状态与分支）、`更新日志.md:12-14,40-42,66`（回线/入库/报废规则、批量工具、回线工令编号检查） |
| `reports` | 一键报告 | Business | high | `Reports/Views/WindowReportTemplate.xaml.cs:22`（入口）、`Services/ReportTemplateService.cs:389`（排期跳周末）、`:406-489`（整目录复制 + 写四表）、`:433-474`（以已有报告为骨架）、`:107-120`（`ExistingReportInfo`）、`Services/ReportGenerationService.cs:33`（`Generate`）、`:171`（NPOI 生成 + 嵌附件）、`:229-238`（**EMI 分支尚未迁移**）、`Services/ReportService.cs:24-64`（右键预填与已有报告优先）、`Reports/ViewModels/BaseReportPageViewModel.cs:62,155-159,167-180`（模板目录、保存路径、Excel COM 嵌 OLE）、`Reports/Models/ReportInputModels.cs:16-113`、`更新日志.md:72-75` |
| `review` | 审核 | Business | high | `Review/Views/WindowReview*.xaml.cs`、`Services/ReviewService.cs:12-15,61,92,190,214`（提交/通过/驳回并触发通知）、`更新日志.md:32-35`（逐字段旧值→新值比对、原始数据折叠） |
| `adminMgmt` | 管理 | Business | high | `Admin/Views/WindowAdmin.xaml.cs:529,547`（导入入口）、`:556-562`（同步客户/产品/测试项目字典）、`:590-603`（清空数据两次确认）、`:223`（重置密码通知）、`Services/AdminService.cs`、`Services/TestPlanService.cs`、`README.md:42,181-190`、`更新日志.md:30`（移除用户不能移除自己） |
| `dataAccess` | 数据访问与数据文件夹 | Platform | high | `Services/DatabaseService.cs:133,146-148`（唯一主连接）、`:104-133`（UNC 自动映射盘符、并发建议）、`:108-110`、`:190-195`（WAL/rollback）、`App.xaml.cs:150-173`（回退只对本次运行生效、不改配置；连默认目录也起不来才致命退出）、`Services/DataFolderMigrator.cs`、`Utils/NetworkDriveMapper.cs`、`Utils/SqliteFileUtil.cs` |
| `background` | 后台调度与耗时任务协调 | Platform | high | `Services/HighCostTaskCoordinator.cs:9-12`（三任务互斥）、`Services/ReportScanScheduler.cs:10-75`（30 秒轮询、闲置默认 300 秒）、`Services/PlanIndexScheduler.cs:9-12`、`Services/PlanIndexService.cs`、`Services/ReportScanService.cs:85,264-281`（`Task.Run` 异步可取消、进度经 Dispatcher 回 UI）、`App.xaml.cs:235`、`README.md:126-129` |
| `backupRestore` | 数据库备份与快照还原 | Platform | high | `Services/DatabaseBackupScheduler.cs:9-20`（启动 90 秒首查、之后每 30 分钟）、`Services/DatabaseBackupService.cs:46-50,63,102-119,165`（≥7 天做全量 `VACUUM INTO`，否则当日增量逐表比对写 `.json`）、`:59-66,84-94`（默认 `Backups`，可用 `backup.folder` 改）、`:68-72,644`（`.ort_backup.lock` 跨机互斥，超 30 分钟视为残留）、`:363-369`（还原链）、`:391`（还原前保底全量）、`:69`（`.ort_restore_tmp.db`）、`:622`（独立连接）、`App.xaml.cs:245-247`、`更新日志.md:67` |
| `mail` | 邮件提醒 | Platform | high | `Services/MailNotifier.cs:16,23,153,187`（`NotifyReviewSubmitted` / `NotifyReviewDecided`）、`Services/MailService.cs:11`、`Services/SmtpMailClient.cs:19`、`MainWindow.xaml.cs:617-620`（到期定时器）、`Main/Views/WindowAppSettings.xaml.cs:512`（手动立即检查）、`Admin/Views/WindowAdmin.xaml.cs:223`（密码变更通知）、`Services/ReviewService.cs:61,92,190,214`（审核通知触发点）、`README.md:118-122`（SMTP 安全方式、DPAPI、模板标记） |
| `settingsI18n` | 设置·本地化·主题 | Platform | high | `Services/AppSettingsService.cs:33`（`app_settings`）、`Services/LocalSettingsStore.cs:12,27,34-35,55`（`local_settings.json`）、`:17,30,33,138`（`db_path.json`/`auth_cookie.json` 迁移后删除）、`Services/PathService.cs:24-32`（Ookii 文件夹对话框）、`Services/LanguageService.cs:30`（`current_language.txt`）、`Services/ThemeService.cs:10-18`（三主题 + `current_theme.txt`/`current_theme_path.txt`）、`Services/OrtLocalizationProvider.cs`（绕开中文程序集名解析失败，`App.xaml.cs:106-109`）、`Services/WindowThemeHelper.cs`、`README.md:96-101,166-173` |
| `officeInterop` | Excel / Word 文件处理与校验 | Platform | high | `Utils/ExcelNpoi.cs:14`（NPOI 兼容层，注释明示替代 EPPlus）、`Utils/ExcelOleEmbedder.cs` + `Utils/ExcelOleWriter.cs:15`（COM 嵌入与兜底）、`Utils/Docx2Pdf.cs:1`、`Utils/DocxTextReplacer.cs`、`Utils/SmartTableExtractor.cs`（OpenXml）、`Utils/OrtPlanParser.cs`、`Utils/PlanValidation.cs`、`Utils/SnVerification.cs`、`ORT一键报告.csproj:628-706`（NPOI 2.7.4 / OpenXml 3.5.1 / itext 9.7.0 / Office COM） |
| `domain` | 领域实体与绑定模型 | Business | high | `Models/Plan.cs:12,14`（表名 + `ObservableObject`）、`Models/Requisition.cs:11`、`Models/User.cs:26,74`、`Models/TestPlan.cs:45,116,117,177,191,349,433,502,556`、`Models/AdminCatalogs.cs:11-126`、`Models/AppSettings.cs:10,34,35,43`、`Models/Review.cs:10,85,111`、`Models/Mail.cs:95`、`Models/DatabaseBackup.cs`、`Models/ReportStatusKind.cs`、`Models/RequisitionDispositionKind.cs`、`Models/Report.cs`（纯 UI 模型，无表特性）、`Reports/Models/ReportInputModels.cs` |

---

## 2. 关系证据

### 2.1 L1 关系（使用者 → 系统，系统 → 外部系统）

| 关系 | 类型 | 同步 | 协议 | 置信度 | 证据 |
| --- | --- | --- | --- | --- | --- |
| 五级角色 → `ort` | calls | sync | in-process | high | `Services/PermissionService.cs:64-107`、`README.md:75-81`（主界面四入口与所需权限） |
| `ort` → `smtp` | calls | async | SMTP | high | `Services/SmtpMailClient.cs:19`、`Services/MailNotifier.cs:10-14`（失败不影响主流程） |
| `ort` → `srcExcel` | reads | sync | file | high | `Services/PlanExcelService.cs:62,187` |
| `ort` → `reportFolder` | reads + writes | sync/async | file | high | 写：`Services/ReportTemplateService.cs:406-489`；读：`Services/ReportScanService.cs:88-127` |
| `ort` → `office` | calls | sync | COM | high | `Utils/Docx2Pdf.cs:1`、`Utils/ExcelOleEmbedder.cs`、`Reports/ViewModels/BaseReportPageViewModel.cs:167-180` |
| `ort` → `winEnv` | calls | sync | Win32 / 注册表 | high | 见 1.2 `winEnv` 行 |

**已确认不存在的出站关系**（检索为空，high 置信的"无"）：HTTP/HTTPS 调用、自动更新检查、消息队列、外部数据库、Web 服务。依据：全仓 `HttpClient`/`WebRequest`/`RestSharp` 检索仅命中 XML 命名空间常量。

### 2.2 L2 关系（容器之间）

| 关系 | 类型 | 置信度 | 证据 |
| --- | --- | --- | --- |
| `wpf` → `db` | reads + writes | high | `Services/DatabaseService.cs:133,146-148`（唯一主连接） |
| `wpf` → `fileStore` | reads + writes | high | `Services/PlanExcelService.cs:691`、`Services/DatabaseBackupService.cs:84-94` |
| `wpf` → `templates` | reads | high | `Services/ReportTemplateService.cs:281-333`、`Reports/Views/WindowMainReport.xaml.cs:93` |
| `xp` → `db` | reads + writes（共用） | high | `xp-client/README.md:1-28`、`xp-client/docs/02-数据契约.md` |
| `xp` → `smtp` | calls | high | `xp-client/README.md:1-28`（XP 上 .NET 4.0 发不出 TLS 1.2，故 Python 自实现） |
| `technician` → `xp` | uses | **low（`Inferred`）** | `xp-client/README.md` 只说"面向 Windows XP 客户端"，**未指明由哪些角色使用**。DSL 中以虚线 + `Inferred` 标签呈现 |

### 2.3 L3 关系（组件之间，仅列需要单独说明的）

绝大多数 L3 关系为 high 置信，证据即 1.4 表中对应组件的行号。以下三条需要特别标注：

| 关系 | 置信度 | 说明与证据 |
| --- | --- | --- |
| `shell` → `mail` | high | 到期提醒定时器**挂在 `MainWindow` 而非独立 Scheduler**：`MainWindow.xaml.cs:617-620`（`_mailTimer.Tick += ... CheckPlanDeadlinesInBackground()`）。这条关系与"后台调度"组件无关，改后台任务时容易漏掉它 |
| `plans` → `review` | high | `Services/ReviewService.cs:61,92`（提交请求）由 `PlansViewModel` 的提审路径调用；提审条件 `Services/PermissionService.cs:58-62` |
| `utils/Report.cs` 反向依赖 `Reports` | high | `Utils/Report.cs:5-7` 引用 `Reports.Views` / `Reports.ViewModels`，说明 `Utils` 不是纯底层工具层。这是一条**层次反向依赖**，在 DSL 中未单独建模（`officeInterop` 组件描述里未涵盖），改动 `Reports` 命名空间时需一并检查 |

### 2.4 模块间跨层引用（代码直接证明）

| 引用方 | 被引用方 | 证据 |
| --- | --- | --- |
| `Plans` | `Main.Views` + `Reports.Views` | `Plans/Views/WindowPlans.xaml.cs:3,7` |
| `Reports` | `Main.Views` + `Admin.Views` | `Reports/Views/WindowReportTemplate.xaml.cs:3-4` |
| `Admin` | `Main.Views` | `Admin/Views/UserControlTestPlans.xaml.cs:3` |
| `Utils` | `Reports.Views` + `Reports.ViewModels` | `Utils/Report.cs:5-7` |

`Plans` 与 `Reports` 之间存在**双向耦合**（`Plans` → `Reports.Views`，且 `Reports/Views/WindowReportTemplate` 从计划右键进入）。

---

## 3. 第三方依赖证据

来源：COM 引用见 `ORT一键报告.csproj:628-655`，`PackageReference` 见 `:656-708`。版本号为逐个核对 `csproj` 所得，使用方经 grep 确认。下表为**完整清单**。

| 库 | 版本 | 使用方（证据） | 置信度 |
| --- | --- | --- | --- |
| FreeSql.Provider.Sqlite | 3.5.310 | `DatabaseService`、`DatabaseBackupService`、`PlanIndexService`、`TestPlanService`、`DataFolderMigrator`、`Models/` 实体特性 | high |
| NPOI | 2.7.4 | `PlanExcelService`、`ReportGenerationService`、`ReportTemplateService`、`ReportStatusReader`、`AdminService`、`Utils/ExcelNpoi.cs:14` 等 20+ 文件 | high |
| DocumentFormat.OpenXml | 3.5.1 | `Utils/DocxTextReplacer.cs`、`Utils/SmartTableExtractor.cs`、`Main/Views/WindowTestPdf.xaml.cs` | high |
| itext | 9.7.0 | `Main/Views/WindowTestPdf.xaml.cs:3-4` | high |
| Microsoft.Office.Interop.Word（COM PIA） | — | `csproj:637-645`、`Utils/Docx2Pdf.cs:1` | high |
| Microsoft.Office.Core / VBIDE（COM PIA） | — | `csproj:628-636`、`:646-654` | high |
| Excel COM（**后期绑定，无 PIA**） | — | `Utils/ExcelOleEmbedder.cs:110-132`；全仓无 `Interop.Excel` 引用 | high |
| CommunityToolkit.Mvvm | 8.4.2 | 全部 `ObservableObject` 实体与 ViewModel | high |
| Microsoft.Extensions.DependencyInjection | 10.0.10 | `App.xaml.cs:111-148` | high |
| WPFLocalizeExtension | 3.10.0 | `Services/LanguageService.cs`、`OrtLocalizationProvider.cs`、`App.xaml.cs:108` | high |
| NLog | 6.1.4 | 全局；`App.xaml.cs:23`、`NLog.config:9-11`（`logs/${shortdate}.log`） | high |
| Newtonsoft.Json | 13.0.4 | `LocalSettingsStore`、`LoginCredentialStore`、`AuthService`、`ReviewService`、`DatabaseBackupService`、`WindowReviewDetail` | high |
| Ookii.Dialogs.Wpf | 5.0.1 | `Services/PathService.cs:30-32` | high |
| Microsoft.Extensions.Logging | 10.0.10 | 随 DI 包引入（`csproj:675-677`） | high |
| NLog.Schema | 6.1.4 | `NLog.xsd` 的 schema 来源（`csproj:690-692`） | high |
| System.Text.Json | 10.0.10 | `csproj:702-704`；与 Newtonsoft.Json 并存 | high |
| System.IO.Compression | 4.3.0 | `csproj:693-695`；zip 处理见 `Services/FileService.cs:16-19`、`Utils/XlsxCleaner.cs` | high |
| System.Runtime | 4.3.1 | `csproj:696-698`（net48 兼容垫片） | high |
| System.Security.Cryptography.Xml | 10.0.10 | `csproj:699-701`；OpenXml 签名相关依赖 | medium |
| MdXaml | 1.27.0 | `Main/Views/WindowChangelog.xaml:7,13`（`xmlns:mdx`）、`Main/Views/WindowChangelog.xaml.cs:13,47`（把 `更新日志.md` 渲染为 WPF FlowDocument，只保留各版本小节） | high |
| EPPlus | — | **已不再使用**，仅 `TmpTools/` 遗留 dll；`App.xaml.cs:67-68` 注明 Excel 读写统一改用 NPOI | high（确认弃用） |

---

## 4. 未证实项与验证任务

按对改动安全性的影响排序。

| # | 未证实项 | 当前状态 | 验证步骤 |
| --- | --- | --- | --- |
| 1 | **导入与提交保存的事务边界** | 静态检索确认 `PlanExcelService` 与 `PlansViewModel.SaveChanges` 中**无 `Transaction` 使用**，均为逐行写入。中途失败时的一致性表现**未测** | 构造中途异常场景（例如第 N 行违反唯一索引）实测库的最终状态；确认"先删→再改→后增"顺序（`PlansViewModel.cs:1666-1758`）在失败时是否留下半完成状态 |
| 2 | **XP 客户端的角色模型** | 未验证是否与主程序五级角色一致 | 读 `xp-client/ort_xp` 的登录与权限判定代码，逐格比对 `Services/PermissionService.cs:64-107` |
| 3 | **XP 客户端的使用者角色** | `low`（DSL 中标 `Inferred` 虚线） | 向业务方确认，或在 `xp-client/docs` 中补记；确认后把 DSL 里 `technician -> xp` 的 `Inferred` 标签去掉或改为实际角色 |
| 4 | **数据库实体 / `app_settings` 键变化对 XP 客户端的波及面** | 已知两边共用库与设置键，但**没有自动校验** | 改 `Models/` 或 `app_settings` 键时人工比对 `xp-client/docs/02-数据契约.md`；`Tests/` 是否覆盖跨端契约需确认 |
| 5 | **EMI 报告的生成路径** | 已知**未迁移**到 `ReportGenerationService`（`:229-238` 明示），仍走 `EMIReportViewModel` | 改报告生成逻辑时必须同时检查两条路径 |
| 6 | **`Plan` ↔ `Requisition` 三级匹配的实际命中率** | 匹配规则已证实（备注含回线RT工令 → 含工令 → 机种+日期），历史数据上的命中率未知 | 对真实库跑一次匹配统计，量化"匹配不到"的比例 |
| 7 | **计划索引的读写路径** | 只确认了 7 张 `plan_index_*` 表的存在与调度触发方式，未追索引的构建与查询细节 | 读 `Services/PlanIndexService.cs` 全文，补出索引的写入时机与查询入口 |
| 8 | **报告模板目录的多根探测优先级** | 确认 `ReportTemplateService.cs:281-333` 做多根探测，未逐个根列出顺序 | 读该区间，列出探测顺序与失败兜底 |
| 9 | **到期提醒在主窗口未创建期间的行为** | 已知定时器挂在 `MainWindow`（`:617-620`），推断登录窗口阶段与 `--background` 早期不会触发；**未实测** | 以 `--background` 启动后观察到期提醒是否按周期发送 |

---

## 5. 证据来源类型统计

| 证据类型 | 占比情况 |
| --- | --- |
| `code` | 主体。全部节点与绝大多数关系均由 `.cs` / `.xaml` / `.csproj` 的行号支撑 |
| `document` | `README.md`（业务与设置说明）、`更新日志.md`（功能演进）、`修改记录.md`（实现细节）、`xp-client/README.md` 与 `xp-client/docs/02-数据契约.md`、`AGENTS.md`（协作约定） |
| `config` | `App.config`（仅 bindingRedirect）、`NLog.config:9-11`、`Data/local_settings.json`、`current_theme.txt`、`current_language.txt` |
| `data` | `Models/` 的 FreeSql 表名特性；无独立 migration 文件（依赖 `UseAutoSyncStructure(true)` 自动迁移） |
| `runtime` | **无**。本轮未采集日志、进程或运行时遥测；所有运行时结论均来自代码静态阅读 |
| `contract` | 无 OpenAPI/AsyncAPI 等机器可读契约。跨容器契约是**人写文档**：`xp-client/docs/02-数据契约.md` 与 `xp-client/docs/schema.md` |
| `human-assumption` | 仅一处：XP 客户端由哪些角色使用（见 §4 第 3 项）。已在 DSL 中以 `Inferred` 标签 + 虚线呈现，未混入确认路径 |

**本轮未做的事**（刻意划出范围，避免被误读为已覆盖）：未运行程序、未采集运行时日志、未执行 `Tests/run-tests.ps1`、未做架构风险评级与技术债盘点（那属于 `risk-quality-reviewer`）、未做部署拓扑与发布流水线分析（属于 `deployment-topology-analyzer`，本文件第 5 节与 `README.md:273-283` 只给出发布形态的事实性描述）、未做演进规划（属于 `evolution-planner`）。
