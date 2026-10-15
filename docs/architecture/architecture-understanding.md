# ORT实验室管理系统 — 当前架构理解

> **这份文档回答的问题**：这个系统是什么、边界在哪里、由哪些部分组成、谁在用、数据与状态归谁管、一条业务主路径怎么走完。
> **读者**：需要在这个仓库里安全改动的开发者，以及需要向实验室业务方解释系统构成的人。
> **范围**：**当前状态**（as-is），不含目标架构、不含改进建议、不含风险评级。
> **依据**：逆向自代码，对应程序集版本 `0.6.2.0`（`Properties/AssemblyInfo.cs:54`），生成日期 2026-10-15。
> **产出方式**：`architecture-visualization` 技能族的 `system-modeler`（场景）+ `c4model`/Structurizr DSL 与 `graphviz`（图源码格式）。

---

## 1. 阅读顺序与产物清单

| 文件 | 回答什么 | 怎么看 |
| --- | --- | --- |
| `system-model.structurizr.dsl` | 结构：L0/L1 全景 → L1 上下文 → L2 容器 → L3 组件，四个视图层层下钻 | 用 Qoder 的 **Structurizr DSL 查看器**打开该文件即可预览；它是本套产物的**唯一真源** |
| `business-flow.dot` | 行为：从 Excel 源表到报告状态回写的端到端业务主路径（7 个阶段） | 用 Qoder 的 **DOT 查看器**打开，或 `dot -Tsvg business-flow.dot -o business-flow.svg` |
| `system-model.evidence.md` | 可信度：每个节点/关系的证据位置、置信度、以及未证实项与验证步骤 | 改代码前先查这里，确认你要动的关系是被证实的还是推断的 |
| 本文件 | 业务语言的解释：能力、角色、实体、状态、边界 | 直接读 |

**先读哪张图**：想快速知道"这系统跟谁打交道"看 `SystemContext`；想知道"有几个可部署的东西、数据在哪"看 `Containers`；想知道"改某个功能会碰到哪些内部件"看 `Components`；想知道"一笔领用怎么走到报告完成"看 `business-flow.dot`。

---

## 2. 系统在做什么（L0 业务能力）

实验室 ORT（Ongoing Reliability Test）业务的桌面管理系统，六项能力（`README.md:37-47`）：

1. **样品领用与退库登记** — 谁领了哪些单体、领了多少、什么时候退。
2. **测试计划排程** — 机种在哪个阶段做哪些测试项、起止日期怎么排。
3. **样品终态登记** — 回线、入库、报废、单体归还，且这四者之间有硬性顺序与互斥约束。
4. **测试报告一键生成** — Thermal Shock / Burn In / EMI / ATE 四类报告，以及给机种生成报告模板骨架。
5. **计划变更审核** — 权限不够的人改动不直接生效，进队列由审核员逐字段比对后通过或驳回。
6. **到期提醒与后台自动化** — 计划临近结束发邮件；空闲时自动建计划索引、自动扫报告夹回写报告状态；数据库定时备份。

---

## 3. 谁在用它（L1 角色与权限边界）

角色是**五级**，权限是**角色 → 动作字符串的硬编码映射**，全部集中在 `Services/PermissionService.cs:64-107`（已逐行核对，高置信）：

| 动作键 | 游客 | 普通用户 | 技术员 | 审核员 | 管理员 |
| --- | :---: | :---: | :---: | :---: | :---: |
| `plan.view` 浏览领退和计划 | ✓ | ✓ | ✓ | ✓ | ✓ |
| `plan.add` / `plan.edit` / `plan.delete` | — | ✓ **需提审** | ✓ | ✓ | ✓ |
| `plan.import` / `plan.export` 导入导出 | — | — | ✓ | ✓ | ✓ |
| `plan.review` / `review.view` 审核 | — | — | — | ✓ | ✓ |
| `report.use` / `report.submit` 一键报告 | — | — | ✓ | **—** | ✓ |
| `report.review` 报告审核 | — | — | — | ✓ | ✓ |
| `admin.manage` 管理模块 | — | — | — | — | ✓ |
| 其他未列出的动作键（default 分支） | — | — | — | — | ✓ |

三个容易被忽略、但改动时必须知道的边界事实：

- **`plan.view` 恒为 `true`**（`PermissionService.cs:74-75`），不看角色。所以"游客能看领退和计划"不是配置出来的，是写死的。
- **审核员拿不到 `report.use`**（`PermissionService.cs:94-96` 只发给技术员与管理员），审核员拿到的是 `report.review`。想给审核员开一键报告，必须改这处映射。
- **default 分支 `return isAdmin`**（`PermissionService.cs:104-105`）：新加一个动作键而忘了在 switch 里登记时，默认只有管理员能用（fail-closed）。

**是否提审由 `PlanEditNeedsReview` 决定**（`PermissionService.cs:58-62`）：只有"是普通用户、且不同时具备技术员/审核员/管理员任一身份"时才需要提审。也就是说角色是**可叠加**的，一个账号兼有普通用户与技术员身份时走直接生效路径。

**登录是纯本地的**：口令 salt+hash 校验、首次启动建默认管理员账号（`Services/AuthService.cs:51-73,98`）；"记住登录"是本地 7 天免登录凭据，用 DPAPI 加密后存 `%LocalAppData%`（`AuthService.cs:216-227`、`Services/LoginCredentialStore.cs:14`）。**全仓没有 HTTP 客户端调用，没有网络登录，也没有自动更新检查**（`HttpClient`/`WebRequest` 检索仅命中 XML 命名空间常量）。

---

## 4. 系统边界与外部依赖（L1）

边界外只有五样东西，其中**只有一样是网络出站**：

| 外部系统 | 关系 | 证据 |
| --- | --- | --- |
| **SMTP 邮件服务器** | 唯一网络出站目标。到期警告 / 审核通知 / 密码变更通知。自实现客户端，支持明文 25、STARTTLS 587、隐式 SSL 465 | `Services/SmtpMailClient.cs:19`、`Services/MailNotifier.cs:10-14` |
| **计划表 / 领用表 Excel 源文件** | 导入源。表头在前 10 行内按去空白关键字定位，年份可从文件名提取，S/N 列可内嵌 OLE 附件 | `Services/PlanExcelService.cs:471,900,910,80,691` |
| **报告文件夹** | 双向：既是一键报告与报告模板工具的输出地，也是报告扫描的输入源。路径由设置项 `ReportDir` 决定 | `Services/ReportScanService.cs:56,88-127` |
| **Microsoft Office** | 仅两个场景需要，且**依赖方式不同**：Excel 走运行时后期绑定（`ProgID` + `dynamic`，不依赖 Interop PIA，只要求本机装了 Excel）嵌 OLE 附件；Word 走编译期 COM 引用（Interop PIA）做 docx→pdf。纯 NPOI 路径两者都不需要 | `Utils/ExcelOleEmbedder.cs:74,110-132`、`ORT一键报告.csproj:637-645`、`Utils/Docx2Pdf.cs:1` |
| **Windows 客户端环境** | 单实例互斥体与命名管道、托盘图标、HKCU Run 开机自启、DPAPI、DWM 深色标题栏 | `Services/SingleInstanceManager.cs:11-25`、`Services/StartupManager.cs:13-15` |

---

## 5. 部署形态（L2 容器）

**一个软件系统，两个客户端，共用一个数据库。** 这是本系统最重要的结构事实：

| 容器 | 是什么 | 技术 |
| --- | --- | --- |
| **WPF 桌面客户端** | `ORT实验室管理系统.exe`。单实例；登录窗口前置决定身份后才创建主窗口；关窗可收进托盘常驻，后台调度与邮件提醒照跑 | .NET Framework 4.8 · WPF · CommunityToolkit.Mvvm · MS.Extensions.DependencyInjection |
| **Windows XP 精简客户端** | 主程序的**子项目，不是替代品**。做：读写 `ort_plans.db`、登录、领用表与计划表的简单交互、SMTP 到期提醒。不做：报告生成、审核、流程窗口、管理端、多语言与主题、计划索引与报告扫描 | Python 3.4（32 位）· tkinter · 仅标准库 · PyInstaller 3.3.1 onedir |
| **SQLite 主库 `ort_plans.db`** | 全部业务数据。FreeSql 开启 `UseAutoSyncStructure(true)`，新增实体列自动迁移 | SQLite · FreeSql 3.5.310 |
| **附件·配图·备份文件区** | 数据文件夹下的 `OleFiles`（从领用表 S/N 列提取的 OLE 附件）、`PlanImages`（计划配图）、`Backups`（周全量 `.db` 快照 + 日增量 `.json` 补丁） | 文件系统 |
| **报告模板目录 `Templates`** | 随程序发布并复制到输出目录的 Excel 模板，含报告模板工具用的整目录骨架 | Excel 文件（NPOI 读写） |

**关于 XP 客户端要知道的两件事**（`xp-client/README.md:1-28`）：

1. 它是**主程序的子项目而非独立产品**，两边靠**数据库表结构与 `app_settings` 键**耦合，而契约是**人写文档**——`xp-client/docs/02-数据契约.md` 与 `xp-client/docs/schema.md`，没有任何自动校验。**改 `Models/` 实体或 `app_settings` 键时，必须同时比对这两份文档与 `xp-client/ort_xp` 代码**，这是本仓库最容易被漏掉的跨容器影响面。
2. 改用 Python 的动因是**邮件**，不是技术栈偏好：XP 上最高只能装 .NET Framework 4.0.3，而 XP 的 Schannel 只支持到 TLS 1.0，.NET 4.0 发不出 TLS 1.2 的邮件；Python 自带 OpenSSL、不走系统 Schannel，XP 上照样能发。因此该客户端**只用标准库**（`sqlite3` / `tkinter` / `smtplib` / `ssl` / `hashlib` / `ctypes`），源码保持 3.4 语法，由 `xp-client/tools/compat_check.py` 卡住语法下限。

**多机共用一套数据的方式**：数据文件夹默认在程序目录 `Data` 下，可在设置里改到网络共享（改后需重启）。SQLite 不能直接用 UNC 路径，所以程序会自动把 `\\服务器\共享` 映射成盘符（`Services/DatabaseService.cs:104-133`、`Utils/NetworkDriveMapper.cs`）。本地库用 WAL，网络共享上退回 rollback journal；建议并发不超过 10 台。配置路径起不来时自动回退到程序目录默认 `Data`，回退**只对本次运行生效、不改配置文件**，连默认目录也起不来才抛致命错误退出（`App.xaml.cs:150-173`）。

---

## 6. WPF 客户端内部（L3 组件）

13 个组件，分**业务**（图中蓝色）与**平台**（图中灰色）两组。

### 业务组件

| 组件 | 负责 | 关键代码 |
| --- | --- | --- |
| **领退与计划** | 核心工作台：领用/退库登记、计划排期、回线·入库·报废·单体归还、批量回线/批量删除/筛选待入库、列筛选与搜索、领退↔计划互转定位、流程窗口、回线工令编号检查 | `Plans/Views`、`Plans/ViewModels`、`Services/PlanExcelService.cs`、`Utils/PlanValidation.cs`、`Utils/OrtPlanParser.cs` |
| **一键报告** | 两条产出路径：报告模板工具（生成机种报告骨架）与一键报告（Thermal Shock / Burn In / EMI / ATE） | `Reports/Views`、`Reports/ViewModels`、`Services/ReportGenerationService.cs`、`ReportTemplateService.cs`、`ReportStatusReader.cs`、`ReportService.cs` |
| **审核** | 计划变更请求队列，逐字段与库中现值比对并标出旧值→新值 | `Review/Views`、`Services/ReviewService.cs` |
| **管理** | 人员与角色、基础数据字典维护与 Excel 导入、邮件服务与数据文件夹设置、重置密码、移除用户、清空全部数据 | `Admin/Views`、`Services/AdminService.cs`、`TestPlanService.cs` |
| **领域实体与绑定模型** | 数据库实体**同时兼作** UI 绑定模型（继承 `ObservableObject`） | `Models/`、`Reports/Models/` |

### 平台组件

| 组件 | 负责 | 关键代码 |
| --- | --- | --- |
| **应用外壳与启动编排** | 启动顺序、单实例、托盘常驻、关窗转后台、注销时按权限强制关闭各模块窗口、**到期邮件定时器** | `App.xaml.cs:41-259`、`MainWindow.xaml.cs:217,241-249,617-620,690-730` |
| **认证与权限判定** | 本地口令校验、记住登录、五级角色映射 | `Services/AuthService.cs`、`PermissionService.cs`、`LoginCredentialStore.cs` |
| **数据访问与数据文件夹** | 唯一主连接、UNC 映射、WAL/rollback 选择、启动回退、数据文件夹迁移、启动时幂等修正历史脏数据 | `Services/DatabaseService.cs`、`DataFolderMigrator.cs`、`Utils/NetworkDriveMapper.cs`、`SqliteFileUtil.cs` |
| **后台调度与耗时任务协调** | 三个高耗时任务互斥；计划索引与报告扫描的闲置触发 | `Services/PlanIndexScheduler.cs`、`ReportScanScheduler.cs`、`HighCostTaskCoordinator.cs`、`PlanIndexService.cs`、`ReportScanService.cs` |
| **数据库备份与快照还原** | 周全量 + 日增量、跨机锁文件互斥、还原链回放 | `Services/DatabaseBackupService.cs`、`DatabaseBackupScheduler.cs` |
| **邮件提醒** | 业务事件触发 → 模板组装 → 自实现 SMTP 投递 | `Services/MailNotifier.cs`、`MailService.cs`、`SmtpMailClient.cs` |
| **设置·本地化·主题** | 设置分两处存放、三语 resx、三套主题运行时换肤、字号等比缩放 | `Services/AppSettingsService.cs`、`LocalSettingsStore.cs`、`LanguageService.cs`、`ThemeService.cs`、`WindowThemeHelper.cs`、`UiScale.cs` |
| **Excel / Word 文件处理与校验** | NPOI 统一读写 xlsx（已取代 EPPlus）、Excel 后期绑定嵌 OLE、Word PIA 转 pdf、OpenXml 解析、日期归一化与序列号核对 | `Utils/ExcelNpoi.cs`、`ExcelOleEmbedder.cs`、`Docx2Pdf.cs`、`DocxTextReplacer.cs`、`SmartTableExtractor.cs`、`SnVerification.cs` |

### 组件层面的三条结构事实（改动前必读）

- **`HighCostTaskCoordinator` 是三任务的单一互斥点**：一键报告 / 计划索引 / 报告扫描同一时间只跑一个，启动新任务先取消旧任务、弹等待窗、等旧任务收尾再启动（`Services/HighCostTaskCoordinator.cs:9-12`）。新增第四类高耗时任务时必须接进来，否则会与扫描抢库。
- **到期邮件定时器不在调度器里，而在 `MainWindow`**：`_mailTimer.Tick` 直接调 `CheckPlanDeadlinesInBackground()`（`MainWindow.xaml.cs:617-620`）。这与计划索引/报告扫描/备份各有独立 Scheduler 的做法不一致——**主窗口未创建时（例如登录窗口阶段）到期提醒不会跑**。
- **模块之间存在跨层与反向引用**（代码直接证明）：`Plans/Views/WindowPlans.xaml.cs:3,7` 同时引用 `Main.Views` 与 `Reports.Views`；`Reports/Views/WindowReportTemplate.xaml.cs:3-4` 引用 `Main.Views` 与 `Admin.Views`；`Utils/Report.cs:5-7` 反向引用 `Reports.Views`/`Reports.ViewModels`。也就是说 `Utils` 并非纯粹的底层工具层。

---

## 7. 领域实体与数据归属

**唯一数据所有者是 SQLite 主库**，全部实体在 `Models/`，用 FreeSql 特性声明表名：

| 表 | 实体 | 证据 |
| --- | --- | --- |
| `plans` | `Plan` | `Models/Plan.cs:12` |
| `requisitions` | `Requisition` | `Models/Requisition.cs:11` |
| `users` / `user_roles` | `User` / `UserRole` | `Models/User.cs:26,74` |
| `test_plans` / `test_plan_items` / `plan_item_templates` / `plan_index_jobs` / `plan_index_entries` / `plan_index_raw_items` / `plan_item_images` | `TestPlan` 族 | `Models/TestPlan.cs:45,116,177,349,433,502,556` |
| `code_mappings` / `stages` / `test_categories` / `products` / `model_mappings` / `plan_change_logs` | 基础数据字典族 | `Models/AdminCatalogs.cs:11-126` |
| `app_settings` / `report_links` | `AppSettings` / `ReportLink` | `Models/AppSettings.cs:10,34` |
| `review_requests` / `customers` / `test_items_catalog` | `ReviewRequest` / `Customer` / `TestItemsCatalog` | `Models/Review.cs:10,85,111` |
| `mail_logs` | `MailLog` | `Models/Mail.cs:95` |
| 备份记录 | `DatabaseBackup` | `Models/DatabaseBackup.cs` |

**表之间没有外键约束，全靠逻辑关联**，这是改动时最容易踩的地方：

- `TestPlanItem.PlanId` → `test_plans`（`Models/TestPlan.cs:191`，唯一显式的外键字段）
- `ReportLink.JobNo` ↔ `Plan.JobNo`，靠唯一索引 `uk_report_link_job` 保证一对一（`Models/AppSettings.cs:35,43`）
- `test_plans` 以 机种名 + 阶段 唯一（`Models/TestPlan.cs:117`）
- `Plan` ↔ `Requisition` **没有键关联**，靠 备注含回线RT工令 → 含工令 → 机种 + 日期 三级匹配（`Plans/ViewModels/PlansViewModel.cs:871-888`、`Services/ReportService.cs:24-64`）

**配置数据的归属按"是否要跟着数据走"一分为二**（这是刻意设计，为的是多台电脑共用一套数据）：

| 存哪 | 存什么 | 证据 |
| --- | --- | --- |
| `app_settings` 表（跟着数据走） | 常规设置、邮件服务、备份开关与目录、报告扫描与计划索引参数 | `Services/AppSettingsService.cs:33` |
| `Data/local_settings.json`（留在本机） | 数据文件夹、ATE/EMI 路径、表格列布局与列宽、后台运行开关 | `Services/LocalSettingsStore.cs:12,27,34-35,55` |
| `%LocalAppData%/ORT实验室管理系统/login.json`（当前 Windows 用户） | 记住登录的用户名与 DPAPI 加密口令 | `Services/LoginCredentialStore.cs:9-14` |
| 程序目录 `current_theme.txt` / `current_theme_path.txt` / `current_language.txt` | 当前主题与语言 | `Services/ThemeService.cs:17-18`、`LanguageService.cs:30` |

历史遗留的 `db_path.json` 与 `auth_cookie.json` 在启动时自动迁移到 `local_settings.json` 后删除（`LocalSettingsStore.cs:17,30,33,138`）。邮件密码以 DPAPI（当前 Windows 用户）加密后入库，不存明文。

---

## 8. 状态与生命周期

系统里并行跑着**四套独立的状态口径**，改动任一处都要确认另外几处不受影响：

| 状态口径 | 取值 | 谁写 | 关键规则 |
| --- | --- | --- | --- |
| **单体去向** `Requisition.Disposition` | 入库 / 报废 | 领退新增编辑窗口；旧数据启动时按有无回线RT工令幂等回填 | 必填。驱动后续分支：入库必须先登记回线，报废与入库互斥（已报废不能再入库，已入库不能再报废，均为**硬拦截**不是二次确认） |
| **报告状态** `Plan.ReportStatus` | 已完成 / 进行中 / 无要求 | 报告扫描增量回写；用户可手工设 | 用户手工设为「无要求」后**扫描不再覆盖**（`ReportStatusKind.IsUserLocked`），改回其他值则下次扫描重新接管（`Services/ReportStatusReader.cs:12-15`、`ReportScanService.cs:191-220`） |
| **计划状况** | 进行中 / 预排测试 / 已结案 | 由登记数据与日期推导，搜索栏右侧彩色徽标实时统计 | 展示口径，见 `更新日志.md:10`；判定逻辑在 `Utils/PlanValidation.cs`（`PlanStatusKind`，有单元测试覆盖） |
| **流程步骤状态** | 已完成 / 进行中 / 未到达 / 无需 / 已驳回 | 流程窗口按登记数据与扫描结果现算 | QRT 流程另有「单体归还」节点；分支未登记时标为「计划走此分支」 |

**暂存与提交是一个显式的两段式生命周期**：界面上的一切编辑（含右键回线/入库/报废、批量登记、批量删除）只改内存暂存并高亮为"待提交"，点「提交保存」才写库，写入顺序为**先删 → 再改 → 后增**以避开唯一索引，并逐条写 `plan_change_logs`（`Plans/ViewModels/PlansViewModel.cs:50,1666-1758`）。关闭领退与计划窗口时若有未提交暂存会先确认。

---

## 9. 端到端业务主路径

完整图见 `business-flow.dot`，七个阶段：

1. **数据进入系统** — 技术员及以上从管理窗口导入领用表/计划表（`plan.import`），或任何人从新增窗口手工登记。导入时按 領料單据號 → 工令 → 日期+机种 做多条件关联合并；日期支持 `8月7日`、`2026/8/18` 等任意格式归一化。
2. **暂存 → 提交保存** — 提交前做必填与工作编号重号的整体校验，重号按表格当前数据（含未提交修改）判定。
3. **审核分支** — `PlanEditNeedsReview` 为真（仅普通用户）时生成 `review_requests`，审核员逐字段比对旧值→新值后通过或驳回；通过才落库并写变更日志，驳回则改动不生效。
4. **样品终态登记** — 按单体去向分入库/报废两条互斥分支，入库前必须先有回线；报废需上传序列号清单与领用清单逐条核对，多一个少一个或清单内重复都列出差异并阻止。技术员及以上直接生效，普通用户提审。
5. **报告产出** — 两条独立路径：**报告模板工具**优先以"同机种周期号最大的已有报告"为骨架（读不到才退回机种测试计划从头生成），自动排期跳周末、超 24 小时按跨天计天数，整目录复制模板到报告夹并写 Cover / ORT Plan / Waterfall / TestStatus 四表，与领用表冲突时以领用表为准并逐条提示；**一键报告**从计划右键携带预填数据，按模板 setup 表的单元格地址写内容，保存后用 Excel COM 嵌 OLE。
6. **报告扫描回写** — 闲置默认 300 秒触发，递归 4 层枚举目录，目录名**包含**计划工作编号即匹配，要求夹内有 `Report` 子目录与概览 xls；读概览 TestStatus 表判定（`UNITS UNDER TEST == TOTAL + FAIL` 则该项完成，全部满足→已完成，否则→进行中，找不到表/表头/测试项行→不改动）；`report_links` 全量刷新，`plans.ReportStatus` 增量回写。**报告状态是扫描环节写的，不是报告生成环节写的。**
7. **到期提醒** — 主窗口定时器周期检查计划到期并后台发送，设置窗口可手动"立即检查并发送警告"。

---

## 10. 事务与并发边界

- **未找到显式数据库事务**：Excel 导入（`PlanExcelService`）与提交保存（`PlansViewModel.SaveChanges`）都是**逐行写入**，静态检索未见 `Transaction` 使用。这是一条**结构观察**，不是运行时行为验证——中途失败时的实际一致性表现未测。
- **多机并发靠两处互斥**：数据库层是 SQLite 自身的锁（本地 WAL / 网络 rollback journal）；备份层是 `Backups` 目录内的 `.ort_backup.lock` 锁文件，超 30 分钟视为残留锁（`Services/DatabaseBackupService.cs:68-72,644`）。
- **进程内并发靠协调器**：三个高耗时任务经 `HighCostTaskCoordinator` 串行化。
- **单实例靠互斥体 + 命名管道**：同一 Windows 用户重复双击不会开出第二个进程，而是唤醒首实例把窗口置前；同一台电脑的不同 Windows 用户各自独立运行（`Services/SingleInstanceManager.cs:11-25`）。
- **还原是危险窗口**：还原前会先给当前库做保底全量，经 `.ort_restore_tmp.db` 临时文件切换，完成后需重启程序，其它电脑上开着的程序要先关闭（`DatabaseBackupService.cs:363-369,391,69`）。

---

## 11. AI 协作指引覆盖情况

`AGENTS.md` **讲清了仓库协作流程，但基本没讲业务领域**：

- **已覆盖**：每轮对话开局的 Git 状态检查与收尾的提交推送、更新日志的"预览版累积 + 用户主动发布"规则与三处版本号一致性、发布打包脚本的 UTF-8 BOM 约束、构建命令为何要带 `SignManifests=false`、重建前要确认 exe 未在运行（`SQLite.Interop.dll` 会被占用）、单元测试用字节加载运行器。
- **已覆盖的唯一一条领域/数据约束**：「数据库为 SQLite + FreeSql（`UseAutoSyncStructure(true)`），新增实体列会自动迁移」。
- **未覆盖（散落在 `README.md`、`更新日志.md` 与代码注释里）**：五级角色与权限矩阵、`report.use` 不给审核员这类反直觉映射、"无回线不能入库 / 报废与入库互斥"的流程不变量、回线RT工令 `RTAH+年月+编号` 的编号规则与历史数据的两种口径、报告状态的四种判定与用户锁定语义、暂存-提交两段式与先删→再改→后增的写入顺序、**改数据库实体或 `app_settings` 键会波及 XP 客户端**。

也就是说：按 `AGENTS.md` 操作能保证提交规范，但**不能保证不破坏业务不变量**。上面这些规则需要靠本文件与 `system-model.evidence.md` 补齐。

---

## 12. 已证实 vs 未证实

**高置信（代码直接证明）**：第 3、4、5、6、7、8、9 节的全部结论，以及第 10 节的互斥与单实例机制。逐条行号见 `system-model.evidence.md`。

**中/低置信与未知项**（改动前需验证）：

| 项 | 状态 | 验证方式 |
| --- | --- | --- |
| 导入与提交保存中途失败时的一致性表现 | **未知** | 只确认了"无显式事务"这一静态事实；需要构造中途异常场景实测 |
| XP 精简客户端自己的角色模型是否与主程序五级角色一致 | **未验证** | 读 `xp-client/ort_xp` 的登录与权限判定代码，比对 `PermissionService.cs:64-107` |
| XP 客户端具体由哪些角色使用 | **推断**（图中 `Inferred` 虚线） | `xp-client/README.md` 只说"面向 Windows XP 客户端"，未指明角色 |
| `Plan` ↔ `Requisition` 三级匹配在历史数据上的实际命中率 | **未知** | 匹配规则本身已证实，命中率需跑真实库统计 |
| 计划索引（`plan_index_*` 七张表）的具体查询路径 | **未展开** | 本轮只确认了表的存在与调度触发，未追索引的读写细节 |
| 报告模板目录的多根探测顺序 | **部分** | 确认了 `ReportTemplateService.cs:281-333` 做多根探测，未逐个根列出优先级 |
| 到期提醒在主窗口未创建期间的行为 | **推断未实测** | 定时器挂在 `MainWindow.xaml.cs:617-620`，推断登录窗口阶段与 `--background` 早期不触发；以 `--background` 启动后观察是否按周期发送 |

**本轮已核实、原先属于推断的两项**（记录在此以免被当成仍待验证）：

- **EMI 报告确未迁移**：`Services/ReportGenerationService.cs:229-238` 的 `GenerateEMIReport` 直接 `throw new NotImplementedException`，注释 TODO 明示"待迁移：`EMIReportViewModel.DoReport` → `ReportGenerationService.GenerateEMIReport`"。改报告生成逻辑必须同时检查两条路径。
- **`MdXaml` 的使用方**：`Main/Views/WindowChangelog.xaml:7,13` 与 `WindowChangelog.xaml.cs:13,47`，把 `更新日志.md` 渲染成 WPF FlowDocument，只保留各「## 版本号」小节。

---

## 13. 如何查看与重新生成

**查看**：`system-model.structurizr.dsl` 用 Qoder 的 Structurizr DSL 查看器打开（四个视图：`Landscape` / `SystemContext` / `Containers` / `Components`）；`business-flow.dot` 用 Qoder 的 DOT 查看器打开，或本机装了 Graphviz 后 `dot -Tsvg business-flow.dot -o business-flow.svg`。

**重新生成**：这套产物是**逆向自代码的快照**，不会自动跟随代码更新。改动落在下列位置时应手动同步本套文件：

- `Services/PermissionService.cs` 的动作映射 → 第 3 节权限矩阵 + DSL 的 L1 角色描述
- `Models/` 增删实体或表名 → 第 7 节表格 + DSL 的 `domain` 组件 + `db` 容器描述
- `App.xaml.cs` 的 DI 注册或启动顺序 → DSL 的 `shell` 组件描述
- 新增高耗时后台任务 → DSL 的 `background` 组件 + 第 10 节并发边界
- 数据库表结构或 `app_settings` 键变化 → 额外检查 `xp-client/docs/02-数据契约.md` 与 XP 客户端代码
- 新增外部依赖或网络出站 → 第 4 节 + DSL 的外部系统与 L1 关系

同步后请更新本文件顶部的程序集版本与生成日期，并在 `system-model.evidence.md` 里补上新的证据行号。
