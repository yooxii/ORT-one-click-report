workspace "ORT实验室管理系统" "ORT（Ongoing Reliability Test）实验室业务桌面管理系统的当前架构模型。逆向自代码，对应程序集版本 0.6.2.0，生成日期 2026-10-15。证据索引见同目录 system-model.evidence.md，业务语言说明见 architecture-understanding.md。标签 Inferred 表示按命名或部分信号推断、尚未被代码直接证实；标签 Unknown 表示刻意留空的待验证区域。" {

    model {

        // ========== L1 使用者（对应 PermissionService 的五级角色） ==========

        guest = person "游客" "无账号时以 CurrentUser=null 进入界面浏览。plan.view 对所有身份恒为 true，因此游客能看「领退和计划」，其余业务入口按权限置灰或关闭窗口。" "Guest"

        generalUser = person "普通用户" "浏览 + 新增/编辑/删除领退与计划记录，但 PlanEditNeedsReview 为真：改动生成审核请求进队列，审核通过后才生效。" "GeneralUser"

        technician = person "技术员" "普通用户的全部权限，外加计划表/领用表导入导出（plan.import / plan.export）与一键报告（report.use / report.submit）。报废登记等技术员及以上直接生效、不走审核。" "Technician"

        reviewer = person "审核员" "审核计划变更请求（review.view / plan.review）与报告审核（report.review），并有导入导出权限。注意：report.use 只发给技术员与管理员，审核员拿不到一键报告入口。" "Reviewer"

        admin = person "管理员" "全部权限，含管理模块（admin.manage）：人员与角色、基础数据字典、邮件服务、数据文件夹、数据库备份与还原。清空全部数据需 admin.manage 加两次弹窗确认。" "Administrator"

        // ========== 外部系统 ==========

        smtp = softwareSystem "SMTP 邮件服务器" "计划到期警告、审核通知、密码变更通知的投递端，也是本系统唯一的网络出站目标。全仓无 HTTP 客户端调用，程序没有自动更新检查。" "External"

        office = softwareSystem "Microsoft Office" "两处依赖方式不同。Excel：运行时后期绑定（Type.GetTypeFromProgID 加 dynamic，不依赖 Interop PIA，只要求本机装了 Excel），用于把 ATE 数据与图片以 OLE 对象嵌入报告。Word：编译期 COM 引用（Microsoft.Office.Interop.Word PIA），用于 docx 转 pdf。纯 NPOI 路径两者都不需要。" "External"

        srcExcel = softwareSystem "计划表 / 领用表 Excel 源文件" "ORT Test Schedule（计划表）与成品領用記錄（领用表）。表头在前 10 行内按去空白关键字定位，年份可从文件名提取，S/N 列可内嵌 OLE 附件。默认目录由设置项给出。" "External"

        reportFolder = softwareSystem "报告文件夹" "各机种的报告目录（需含 Report 子目录与报告概览 xls）。既是一键报告与报告模板工具的输出地，也是报告扫描回写报告状态的输入源。路径由设置项 ReportDir 决定。" "External"

        winEnv = softwareSystem "Windows 客户端环境" "单实例互斥体与命名管道、托盘图标、HKCU Run 开机自启项、DPAPI 凭据加密、%LocalAppData% 登录凭据、DWM 深色标题栏。程序位于网络共享/映射盘/SUBST 虚拟盘时开机自启被禁用。" "External"

        // ========== 目标系统 ==========

        ort = softwareSystem "ORT实验室管理系统" "实验室 ORT 业务的桌面管理系统：样品领用与退库、测试计划排程、测试报告一键生成、计划变更审核、到期邮件提醒。数据存 SQLite，数据文件夹可放网络共享，多台电脑共用一套数据（建议并发不超过 10）。" {

            wpf = container "WPF 桌面客户端" "ORT实验室管理系统.exe。单实例运行；登录窗口前置决定身份后才创建主窗口；关闭主窗口可收进托盘常驻，后台调度与邮件提醒照跑。主界面按权限显示一键报告 / 领退和计划 / 管理 / 审核四个入口，另有设置、日志查看、更新日志、用户中心与报告模板工具。" ".NET Framework 4.8 · WPF · CommunityToolkit.Mvvm · Microsoft.Extensions.DependencyInjection" "Client" {

                shell = component "应用外壳与启动编排" "App.OnStartup 顺序：取单实例互斥体并开始监听重复启动 → ShutdownMode 改显式退出 → 初始化语言与主题 → 注册全窗口 Loaded 类处理器（深色标题栏、字体应用、开窗防闪）→ 装自定义本地化 Provider → 建 DI 容器注册全部 Services 与 ViewModels → 检查数据文件夹可用性（不可用则回退默认 Data 并告警，连默认目录也起不来才致命退出）→ 幂等修正历史公式脏数据 → 弹登录窗口 → 手建 MainWindow（支持 --background 直接进托盘）→ 启动报告扫描调度与数据库备份调度。MainWindow 负责托盘常驻、关窗转后台、到期邮件定时器，以及注销时按权限强制关闭各模块窗口。" "App.xaml.cs · MainWindow.xaml.cs · SingleInstanceManager · StartupManager" "Platform"

                authz = component "认证与权限判定" "AuthService：本地口令 salt+hash 校验、首次启动建默认管理员账号、「记住登录」为本地 7 天免登录凭据（DPAPI 加密存 %LocalAppData%）。PermissionService：五级角色到动作字符串的硬编码映射，并提供 PlanEditNeedsReview 决定普通用户改动是否进审核队列。没有任何网络登录。" "AuthService · PermissionService · LoginCredentialStore" "Platform"

                plans = component "领退与计划" "核心工作台：领用/退库登记、测试计划排期、回线·入库·报废·单体归还登记、批量回线 / 批量删除 / 筛选待入库、列筛选与搜索（250ms 防抖）、领退表与计划表按回线RT工令/工令/机种+日期互转定位、流程窗口、回线工令编号检查。界面改动先暂存并高亮，点「提交保存」才写库，写入顺序为先删→再改→后增以避开唯一索引，并逐条记变更日志。流程依赖规则：无回线不能入库，报废与入库互斥。" "Plans/Views · Plans/ViewModels · PlanExcelService · PlanValidation · OrtPlanParser · FlowViewModel · ReturnRtCodeCheck" "Business"

                reports = component "一键报告" "两条产出路径。① 报告模板工具：按机种取测试计划自动排期（跳周末，超 24 小时按跨天计天数），或优先以「同机种周期号最大的已有报告」为骨架，整目录复制模板到报告夹并写 Cover / ORT Plan / Waterfall / TestStatus 四表；已有报告与领用表冲突时以领用表为准并逐条提示。② 一键报告：Thermal Shock / Burn In / EMI / ATE，从计划右键携带 MatchedPlan、MatchedRequisition、MatchedReportDir 预填，按模板 setup 表的单元格地址写表头与单体数据，保存后用 Excel COM 嵌 OLE 附件。" "Reports/Views · Reports/ViewModels · ReportGenerationService · ReportTemplateService · ReportStatusReader · ReportService" "Business"

                review = component "审核" "计划变更请求队列：逐字段与库中现值比对并标出「旧值 → 新值」，审计字段与界面计算字段不展示，完整 JSON 收在可折叠的原始数据区。审核员与管理员可通过或驳回。" "Review/Views · ReviewService" "Business"

                adminMgmt = component "管理" "基础数据与人员权限维护：人员与角色、客户、测试项目、产品别、测试阶段、机种映射字典的维护与 Excel 导入；邮件服务与数据文件夹设置（仅管理员可见）；重置用户密码；移除用户（不能移除自己，保证至少剩一个管理员）；清空全部数据（需 admin.manage 加两次确认）。" "Admin/Views · AdminService · TestPlanService" "Business"

                dataAccess = component "数据访问与数据文件夹" "DatabaseService 是唯一主连接：启动时按本机设置的 DataFolder 初始化，失败自动回退到程序目录默认 Data（回退只对本次运行生效、不改配置文件），连默认目录也起不来才抛致命错误退出；UNC 路径自动映射成盘符；本地用 WAL、网络共享退回 rollback journal。DataFolderMigrator 负责换数据文件夹时的数据迁移。" "DatabaseService · DataFolderMigrator · NetworkDriveMapper · SqliteFileUtil" "Platform"

                background = component "后台调度与耗时任务协调" "三个高耗时任务（一键报告 / 计划索引 / 报告扫描）经 HighCostTaskCoordinator 互斥：启动新任务先取消旧任务、弹等待窗、等旧任务收尾再启动。PlanIndexScheduler 空闲 N 分钟后建计划索引以加快搜索筛选；ReportScanScheduler 每 30 秒轮询、闲置默认 300 秒触发扫描，扫描可中断，进度回主窗口状态栏。" "PlanIndexScheduler · ReportScanScheduler · HighCostTaskCoordinator · PlanIndexService · ReportScanService" "Platform"

                backupRestore = component "数据库备份与快照还原" "DispatcherTimer 启动 90 秒后首查、之后每 30 分钟检查一次：距上次全量已满 7 天做全量（VACUUM INTO 生成 .db 快照），否则当日做增量（逐表比对差异写 .json 补丁）。多台电脑共用同一库时靠 Backups 目录内的锁文件互斥，超 30 分钟视为残留锁。还原＝基准全量加上按序回放增量，还原前先给当前库做保底全量，经临时库文件切换，完成后需重启程序。" "DatabaseBackupService · DatabaseBackupScheduler" "Platform"

                mail = component "邮件提醒" "MailNotifier 按业务事件触发 → MailService 按模板组装（标记 {{变量}} 与 {{变量|格式}}）→ 自实现 SmtpMailClient 投递，支持明文 25 / STARTTLS 587 / 隐式 SSL 465，可忽略证书错误、可用 Windows 集成验证。发件密码以 DPAPI 加密后入库，发送记录写 mail_logs。发送失败不影响主流程。" "MailNotifier · MailService · SmtpMailClient" "Platform"

                settingsI18n = component "设置·本地化·主题" "设置分两类存放：跟着数据走的存 app_settings 表（常规、邮件、备份、报告扫描、计划索引），留在本机的存 Data 目录下的 local_settings.json（数据文件夹、ATE/EMI 路径、表格列布局、后台运行开关），历史遗留的 db_path.json 与 auth_cookie.json 启动时自动迁移后删除。语言为简中/繁中/英文三套 resx，经自定义本地化 Provider 绕开中文程序集名解析失败。主题为 Fluent / Material / DarkLab，以 ResourceDictionary 合并加 DynamicResource 运行时换肤，支持导入自定义 XAML 主题；字号变化时窗口、标签与表格列宽等比放大。" "AppSettingsService · LocalSettingsStore · PathService · LanguageService · OrtLocalizationProvider · ThemeService · WindowThemeHelper · UiScale · ToastService" "Platform"

                officeInterop = component "Excel / Word 文件处理与校验" "NPOI 统一承担 xlsx 读写（已取代 EPPlus，EPPlus 仅剩 TmpTools 里的遗留 dll）；ExcelOleEmbedder 以运行时后期绑定调 Excel COM 嵌 OLE 附件（ProgID 加 dynamic，不依赖 PIA；dynamic 封送失败与 Excel 崩溃都有兜底）并有 ExcelOleWriter 兜底；Docx2Pdf 走 Word COM PIA；DocxTextReplacer 与 SmartTableExtractor 走 OpenXml。解析与校验工具负责任意格式日期归一化、计划文本解析、必填与重号校验、报废序列号清单逐条核对。" "Utils/ExcelNpoi · ExcelOleEmbedder · ExcelOleWriter · Docx2Pdf · DocxTextReplacer · SmartTableExtractor · OrtPlanParser · PlanValidation · SnVerification" "Platform"

                domain = component "领域实体与绑定模型" "数据库实体同时兼作 UI 绑定模型（继承 ObservableObject）：Plan、Requisition、User 与 UserRole、TestPlan 族（测试计划、测试项、测试项模板、计划索引任务/条目/原始项、测试项配图）、基础数据字典族（机种映射、阶段、测试类别、产品、代码映射、计划变更日志）、AppSettings、ReportLink、ReviewRequest、Customer、TestItemsCatalog、MailLog、DatabaseBackup。表之间没有外键约束，靠逻辑关联：TestPlanItem.PlanId 指向测试计划，ReportLink.JobNo 与 Plan.JobNo 靠唯一索引对应，Plan 与 Requisition 靠机种与序列号匹配。" "Models/ · Reports/Models/" "Business"
            }

            xp = container "Windows XP 精简客户端" "面向 Windows XP SP3 的精简客户端。做：读写主程序的 ort_plans.db、登录、领用表与计划表的简单交互、SMTP 计划到期提醒。不做：报告生成、审核、流程窗口、管理端、多语言与主题体系、计划索引与报告扫描。与 WPF 主程序共用同一个 ort_plans.db 与同一批 app_settings 设置键（是子项目而非替代关系），数据契约写在 xp-client/docs/02-数据契约.md。改用 Python 的动因是 XP 最高只能装 .NET 4.0.3、其 Schannel 只到 TLS 1.0，发不出 TLS 1.2 邮件。" "Python 3.4（32 位）· tkinter · 仅标准库（sqlite3 / smtplib / ssl / hashlib / ctypes）· PyInstaller 3.3.1 onedir" "Client"

            db = container "SQLite 主库 ort_plans.db" "计划、领用、审核请求、用户与角色、基础数据字典、邮件发送记录、报告链接、备份记录、计划索引。FreeSql 开启 UseAutoSyncStructure(true)，新增实体列自动迁移，升级不需手工改库。默认位于程序目录 Data 下，可在设置里改到网络共享（改后需重启，并可选择把现有数据复制过去）。" "SQLite · FreeSql 3.5.310" "Database"

            fileStore = container "附件·配图·备份文件区" "数据文件夹下的三个子目录：OleFiles 存从领用表 S/N 列提取的 OLE 附件，PlanImages 存计划配图，Backups 存周全量 .db 快照与日增量 .json 补丁（位置可用设置项 backup.folder 改到别处）。备份范围只有数据库本身，附件与配图不在内。" "文件系统" "Database"

            templates = container "报告模板目录 Templates" "随程序发布并复制到输出目录的 Excel 报告模板：Thermal Shock Test、ORT Burn-In Test Report、ATE report，以及报告模板工具用的整目录骨架（名为 # ORT Test Report (WK#)_# 的文件夹）。生成时按多根探测定位模板目录，再整目录复制到报告文件夹。" "Excel 文件（NPOI 读写）" "Database"
        }

        // ========== L1 关系：使用者与系统 ==========

        guest -> ort "浏览领退和计划（只读）"
        generalUser -> ort "登记领用与退库、排测试计划；改动提交审核"
        technician -> ort "上述全部，外加导入导出计划表与领用表、一键报告、报废登记直接生效"
        reviewer -> ort "审核计划变更请求与报告，导入导出"
        admin -> ort "维护人员权限与基础数据、邮件服务、数据文件夹、备份与还原"

        // ========== L1 关系：系统与外部系统 ==========

        ort -> smtp "发到期警告、审核通知、密码变更邮件" "SMTP 明文 25 / STARTTLS 587 / 隐式 SSL 465（自实现客户端）"
        ort -> srcExcel "读取并导入计划表与领用表，提取 S/N 列内嵌的 OLE 附件" "NPOI 读 xlsx / xls"
        ort -> reportFolder "扫描报告夹回写报告状态；生成的报告写入报告夹" "文件 IO"
        ort -> office "嵌入 OLE 附件、docx 转 pdf" "Excel / Word COM 互操作"
        ort -> winEnv "单实例互斥体与命名管道、托盘、HKCU Run 开机自启、DPAPI 加密凭据、DWM 深色标题栏" "Win32 API / 注册表"

        // ========== L2 关系：使用者与容器 ==========

        guest -> wpf "浏览"
        generalUser -> wpf "日常登记与排程"
        technician -> wpf "日常登记、导入导出、一键报告"
        reviewer -> wpf "审核"
        admin -> wpf "管理与设置"
        technician -> xp "在 Windows XP SP3 机器上用精简客户端查领退与计划、收到期提醒" "" "Inferred"

        // ========== L2 关系：容器之间 ==========

        wpf -> db "读写全部业务数据" "FreeSql / SQLite（本地 WAL，网络共享退回 rollback journal）"
        wpf -> fileStore "写 OLE 附件与计划配图；读写全量与增量备份" "文件 IO"
        wpf -> templates "读取报告模板与整目录骨架" "文件 IO"
        xp -> db "共用同一主库与同一批 app_settings 设置键" "Python sqlite3" "Shared"
        xp -> smtp "到期提醒" "SMTP（XP 上 .NET 4.0 发不出 TLS 1.2，故改用 Python 实现）"

        // ========== L3 关系：组件之间 ==========

        shell -> authz "启动时先弹登录窗口决定身份，登录失败或选退出则结束程序"
        shell -> settingsI18n "启动时初始化语言与主题，并对每个窗口应用字体与深色标题栏"
        shell -> dataAccess "启动时检查数据文件夹可用性，不可用则回退默认目录或致命退出"
        shell -> background "主窗口显示后启动报告扫描调度与计划索引调度"
        shell -> backupRestore "主窗口显示后启动数据库定时备份调度"
        shell -> mail "主窗口内的定时器周期检查计划到期并后台发送警告"

        authz -> dataAccess "读写 users 与 user_roles，校验口令散列"
        authz -> winEnv "「记住登录」凭据以 DPAPI 加密存 %LocalAppData%"

        plans -> domain "绑定 Plan 与 Requisition 实体，暂存改动并做快照对比"
        plans -> dataAccess "提交保存时按先删→再改→后增写库，并逐条写计划变更日志"
        plans -> officeInterop "导入导出 Excel、日期与文本归一化、必填与重号校验、报废序列号核对"
        plans -> review "普通用户的改动生成审核请求" "in-process"
        plans -> reports "右键一键报告或报告模板时携带匹配到的计划、领退记录与报告夹预填"
        plans -> background "手动触发计划索引与报告扫描，经协调器与一键报告互斥"
        plans -> srcExcel "导入计划表与领用表" "NPOI"

        reports -> templates "按模板骨架生成报告" "NPOI"
        reports -> officeInterop "写 xlsx、以 Excel COM 嵌 OLE 附件、以 Word COM 转 pdf"
        reports -> reportFolder "生成的报告写入报告夹；测试信息与图片优先读绑定报告夹里的已有报告" "文件 IO"
        reports -> dataAccess "读计划、领用与机种测试计划"
        reports -> background "一键报告经协调器与计划索引、报告扫描互斥"
        reports -> domain "消费报告输入模型与报告数据模型"

        review -> dataAccess "读写 review_requests，通过或驳回后落库并写变更日志"
        review -> mail "提交、通过、驳回时发审核通知"
        review -> domain "绑定 ReviewRequest 实体"

        adminMgmt -> dataAccess "维护 users 与 user_roles、基础数据字典；执行清空全部数据"
        adminMgmt -> officeInterop "基础数据导入读 Excel"
        adminMgmt -> settingsI18n "维护邮件服务与数据文件夹设置（仅管理员可见）"
        adminMgmt -> mail "重置用户密码时发密码变更通知"

        background -> dataAccess "写计划索引表、报告链接表，并增量回写计划的报告状态（用户锁定为无要求时不覆盖）"
        background -> reportFolder "递归 4 层扫描报告夹，目录名包含工作编号即匹配" "文件 IO"
        background -> domain "绑定计划索引与报告链接实体"

        backupRestore -> db "全量用 VACUUM INTO 生成快照，增量逐表比对差异" "独立 SQLite 连接"
        backupRestore -> fileStore "备份文件写入 Backups，靠目录内锁文件在多台电脑间互斥" "文件 IO"
        backupRestore -> dataAccess "读取数据库路径与备份设置"

        settingsI18n -> dataAccess "跟着数据走的设置存 app_settings 表"
        settingsI18n -> winEnv "主题与语言记录写程序目录文本文件；开机自启写 HKCU Run"

        dataAccess -> db "唯一主连接" "FreeSql / SQLite"
        dataAccess -> fileStore "管理附件与配图目录，换数据文件夹时迁移数据" "文件 IO"

        officeInterop -> office "OLE 嵌入与 docx 转 pdf" "Excel / Word COM"
        officeInterop -> srcExcel "读取导入源表" "NPOI"

        mail -> smtp "投递邮件" "自实现 SMTP 客户端"
        mail -> dataAccess "读设置与计划数据，写 mail_logs 发送记录"

        domain -> db "FreeSql 实体映射，开启自动结构同步" "ORM 映射"
    }

    views {

        systemLandscape "Landscape" "L0/L1 全景：五级角色、ORT实验室管理系统、以及它依赖的外部系统" {
            include *
            autoLayout tb
        }

        systemContext ort "SystemContext" "L1 系统上下文：系统边界内只有 ORT实验室管理系统，边界外是使用者的五种角色与五个外部依赖" {
            include *
            autoLayout tb
        }

        container ort "Containers" "L2 容器：两个客户端共用一个 SQLite 主库，另有附件配图备份文件区与报告模板目录" {
            include *
            autoLayout tb
        }

        component wpf "Components" "L3 组件：WPF 客户端内部的业务组件（蓝）与平台组件（灰）" {
            include *
            autoLayout tb
        }

        styles {
            element "Person" {
                shape person
                background #1f6f8b
                color #ffffff
                fontSize 16
            }
            element "Software System" {
                background #1168bd
                color #ffffff
                fontSize 16
            }
            element "External" {
                background #8c8496
                color #ffffff
            }
            element "Container" {
                background #438dd5
                color #ffffff
            }
            element "Client" {
                background #2f6f9f
                color #ffffff
            }
            element "Database" {
                shape cylinder
                background #3d7a52
                color #ffffff
            }
            element "Component" {
                background #b5cdf0
                color #000000
            }
            element "Business" {
                background #4f9dd8
                color #ffffff
            }
            element "Platform" {
                background #c9d4e0
                color #000000
            }
            element "Shared" {
                border dashed
            }
            element "Inferred" {
                border dashed
                opacity 70
            }
            element "Unknown" {
                border dotted
                opacity 55
            }
            relationship "Relationship" {
                fontSize 13
            }
        }
    }
}
