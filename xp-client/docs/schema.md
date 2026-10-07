# 表结构清单（主程序模型 → 本客户端数据契约）

> 本文件由 `tools/extract_schema.py` 从主程序 `Models/*.cs` 自动生成，**请勿手改**。
> 主程序改了模型后重新运行：`python tools/extract_schema.py`

- 表数量：23
- 列数量：217

## `code_mappings`

代码映射（code_mappings 表）：Cust. Code 工作表的两位代码 → 名称。 CodeType：C=客户别（B、C 列：Cust. Code→ENDCUSTOMER），P=产品别（G、H 列：Code→Product Type）。 查询规则：客户别 = 机种名称第 8 位起的 2 位；产品别 = 机种名称开始的 2 位。

来源：`Models/AdminCatalogs.cs` → `CodeMapping`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `CodeType` | string | NVARCHAR(1) | 否 |  | 代码类型：C=客户别，P=产品别 |
| `Code` | string | NVARCHAR(8) | 否 |  | 两位代码 |
| `Name` | string | NVARCHAR(64) | 否 |  | 对应名称（客户名/产品类型名） |

## `stages`

阶段字典（stages 表）：阶段名 + 描述。初始值 MP/EVT/DVT/PVT/RMA

来源：`Models/AdminCatalogs.cs` → `Stage`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Name` | string | NVARCHAR(32) | 否 |  | 阶段名（如 MP/EVT/DVT/PVT/RMA） |
| `Description` | string | NVARCHAR(256) | 是 |  | 描述 |

## `test_categories`

测试种类字典（test_categories 表）：测试项目的归类（RELIABILITY TEST / EMC / 不确定…）， 报告的 ORT Plan 分类行与 TestStatus 分组行都按它来分，可在管理界面增删改。

来源：`Models/AdminCatalogs.cs` → `TestCategory`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Name` | string | NVARCHAR(64) | 否 |  |  |
| `Description` | string | NVARCHAR(256) | 是 |  |  |

## `products`

产品别字典（products 表），数据源为计划表 Cust. Code 工作表的 G、H 列（Code, Product Type）

来源：`Models/AdminCatalogs.cs` → `Product`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Name` | string | NVARCHAR(64) | 否 |  |  |
| `Code` | string | NVARCHAR(32) | 是 |  | 产品代码（Cust. Code 表 G 列） |
| `Remark` | string | NVARCHAR(256) | 是 |  |  |

## `model_mappings`

机种映射（model_mappings 表）：还原计划表公式关系—— 输入机种名称即可带出对应的产品别与客户别

来源：`Models/AdminCatalogs.cs` → `ModelMapping`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `ModelName` | string | NVARCHAR(128) | 否 |  |  |
| `Product` | string | NVARCHAR(64) | 是 |  |  |
| `Customer` | string | NVARCHAR(64) | 是 |  |  |

## `plan_change_logs`

计划数据变更日志（plan_change_logs 表）：记录每次提交的更改前后快照，便于追溯与回滚

来源：`Models/AdminCatalogs.cs` → `PlanChangeLog`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Action` | string | NVARCHAR(16) | 否 |  | 操作：新增 / 编辑 / 删除 |
| `PlanId` | long | INTEGER | 否 |  | 计划记录Id（新增时为提交后的Id） |
| `Summary` | string | NVARCHAR(256) | 是 |  | 变更摘要 |
| `BeforeJson` | string | TEXT | 是 |  | 变更前快照（JSON，新增时为null） |
| `AfterJson` | string | TEXT | 是 |  | 变更后快照（JSON，删除时为null） |
| `Operator` | string | NVARCHAR(64) | 否 |  | 操作人 |
| `CreatedAt` | DateTime | DATETIME | 否 |  |  |

## `app_settings`

设置键值对实体（app_settings 表）：所有设置项以键值对形式保存到数据库。 注：数据库路径设置项单独保存在程序目录文件中（避免自引用）。

来源：`Models/AppSettings.cs` → `AppSetting`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Key` | string | NVARCHAR(64) | 否 |  | 设置键（如 ui.fontFamily / paths.report） |
| `Value` | string | NVARCHAR(512) | 是 |  | 设置值 |

## `report_links`

报告链接实体（report_links 表）：记录按 RT 工作编号在报告路径下找到的报告文件夹。 报告夹结构：文件夹名包含工作编号，内含 Report 子文件夹与一个 Excel 报告概览文件。

来源：`Models/AppSettings.cs` → `ReportLink`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `JobNo` | string | NVARCHAR(64) | 否 |  | 工作编号（与计划表对应） |
| `ReportDir` | string | NVARCHAR(512) | 是 |  | Report 子文件夹完整路径（打开报告文件夹目标） |
| `OverviewFile` | string | NVARCHAR(512) | 是 |  | 报告概览 Excel 文件完整路径（与 Report 同级） |
| `UpdatedAt` | System.DateTime? | DATETIME | 是 |  | 扫描时间 |

## `mail_logs`

邮件发送记录（mail_logs 表）：用于去重（同一对象同类邮件间隔）与排查

来源：`Models/Mail.cs` → `MailLog`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Kind` | string | NVARCHAR(16) | 否 |  |  |
| `RefType` | string | NVARCHAR(32) | 是 |  |  |
| `RefKey` | string | NVARCHAR(128) | 是 |  |  |
| `Recipients` | string | NVARCHAR(1024) | 是 |  |  |
| `Subject` | string | NVARCHAR(512) | 是 |  |  |
| `Success` | bool | BOOLEAN | 否 |  |  |
| `Error` | string | NVARCHAR(1024) | 是 |  |  |
| `CreatedAt` | DateTime | DATETIME | 否 |  |  |

## `plans`

计划表实体（plans 表）：ORT Test Schedule 数据。 与领退表（requisitions 表）分表存储。 工作編號唯一（QRT 前缀为非领用计划，RT 前缀为正常领用计划）。 实现属性通知以便表格内编辑/机种联动自动带出时单元格即时刷新。

来源：`Models/Plan.cs` → `Plan`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 | 自增主键索引 |
| `ModelName` | string | NVARCHAR(128) | 是 |  | 機種名/Part No |
| `TestItem` | string | NVARCHAR(128) | 是 |  | 測試項目/Test Item |
| `Remark` | string | NVARCHAR(512) | 是 |  | 備 考/Remark |
| `JobNo` | string | NVARCHAR(64) | 是 |  | 工作編號/Job No（唯一） |
| `Product` | string | NVARCHAR(64) | 是 |  | 產品別/Product |
| `Customer` | string | NVARCHAR(64) | 是 |  | 客戶別/Customer |
| `Stage` | string | NVARCHAR(32) | 是 |  | 階 段/Stage |
| `SampleSize` | string | NVARCHAR(32) | 是 |  | 樣品數/Sample Size |
| `TestPeriod` | string | NVARCHAR(32) | 是 |  | 試驗時間/Test Period |
| `Owner` | string | NVARCHAR(64) | 是 |  | 負責人/Owner |
| `StartDate` | System.DateTime? | DATETIME | 是 |  | 開始日期/Start Date（日期类型） |
| `EndDate` | System.DateTime? | DATETIME | 是 |  | 結束日期/End Date（日期类型） |
| `Status` | string | NVARCHAR(32) | 是 |  | 完成狀況/Status（Close/Ongoing/Pending） |
| `ReportStatus` | string | NVARCHAR(16) | 是 |  | 报告状态（已完成 / 进行中 / 无要求）：扫描报告文件夹时按 TestStatus 表自动写入； 用户手工设为「无要求」后不再被扫描覆盖，改回其他值则下次扫描重新接管。 null 表示尚未扫描且用户未设置。 |
| `UploadELab` | string | NVARCHAR(32) | 是 |  | 上傳系統/Upload e-lab |
| `UnitReturnDate` | System.DateTime? | DATETIME | 是 |  | 單體歸還日期（其他部門申請測試流程：完成報告後歸還單體， 登记此日期后流程查看里的「單體歸還」步骤判定为完成） |
| `CreatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `CreatedAt` | System.DateTime? | DATETIME | 是 |  |  |
| `UpdatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `UpdatedAt` | System.DateTime? | DATETIME | 是 |  |  |

## `requisitions`

领退表实体（requisitions 表）：记录成品領用/回线信息。 与计划表（plans 表）分表存储，通过领料单据号/WorkOrder/回线RT工令关联。

来源：`Models/Requisition.cs` → `Requisition`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 | 自增主键索引 |
| `RequisitionDate` | System.DateTime? | DATETIME | 是 |  | 領用日期（必填，日期类型） |
| `RequisitionNo` | string | NVARCHAR(64) | 是 |  | 領料單据號（必填，唯一） |
| `ModelName` | string | NVARCHAR(128) | 是 |  | 機種名稱（必填） |
| `OutQty` | string | NVARCHAR(32) | 是 |  | 領出數量（必填） |
| `Disposition` | string | NVARCHAR(16) | 是 |  | 單體去向（入库 / 报废，二选一）：领退表新增/编辑时必选； 报废→走报废分支、入库→走回线入库分支（流程查看据此标出计划走的分支） |
| `SN` | string | NVARCHAR(2048) | 是 |  | S/N（必填，字符串或附件形式） |
| `SnFilePath` | string | NVARCHAR(512) | 是 |  | SN附件文件路径（S/N 为附件形式时） |
| `Rev` | string | NVARCHAR(32) | 是 |  | REV.（必填） |
| `WorkOrder` | string | NVARCHAR(64) | 是 |  | Work Order（必填） |
| `DC` | string | NVARCHAR(32) | 是 |  | D/C（自动补全：WorkOrder 倒数第三位起的两位表示第多少周） |
| `LineNo` | string | NVARCHAR(32) | 是 |  | 線別（自动补全：WorkOrder 倒数第六位起的三位字符串） |
| `ReturnRtOrder` | string | NVARCHAR(64) | 是 |  | 回线RT工令（可选自动生成：RTAH{当前年月}{编号}） |
| `ReturnQty` | string | NVARCHAR(32) | 是 |  | 回線數量 |
| `ReturnDate` | System.DateTime? | DATETIME | 是 |  | 回線日期 |
| `StockInNo` | string | NVARCHAR(64) | 是 |  | 入庫退料單据號 |
| `StockInQty` | string | NVARCHAR(32) | 是 |  | 入庫數量 |
| `StockInDate` | System.DateTime? | DATETIME | 是 |  | 入庫日期 |
| `ScrapNo` | string | NVARCHAR(64) | 是 |  | 報廢單据號（可空） |
| `ScrapQty` | string | NVARCHAR(32) | 是 |  | 報廢數量 |
| `ScrapDate` | System.DateTime? | DATETIME | 是 |  | 報廢日期 |
| `ScrapSnText` | string | TEXT | 是 |  | 報廢序列號清單（文本模式；与 ScrapSnFilePath 二选一） |
| `ScrapSnFilePath` | string | NVARCHAR(512) | 是 |  | 報廢序列號文件（文件模式，存 OleDir 相对文件名；与 ScrapSnText 二选一） |
| `Remark` | string | NVARCHAR(512) | 是 |  | 备注 |
| `CreatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `CreatedAt` | System.DateTime? | DATETIME | 是 |  |  |
| `UpdatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `UpdatedAt` | System.DateTime? | DATETIME | 是 |  |  |

## `review_requests`

审核请求实体（review_requests 表），工作流形式： 请求方提交（如计划表单更改），审核员通过（应用更改）或驳回。

来源：`Models/Review.cs` → `ReviewRequest`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Type` | string | NVARCHAR(32) | 否 |  | 请求类型（当前仅"计划表单"，预留"报告"类型） |
| `Action` | string | NVARCHAR(16) | 否 |  | 操作类型：新增 / 编辑 / 删除 |
| `TargetId` | long? | INTEGER | 是 |  | 目标记录Id（编辑/删除时有值） |
| `Summary` | string | NVARCHAR(256) | 是 |  | 请求摘要（列表展示用） |
| `PayloadJson` | string | TEXT | 是 |  | 更改内容（Plan 序列化 JSON） |
| `RequesterName` | string | NVARCHAR(64) | 否 |  | 请求人 |
| `AssigneeName` | string | NVARCHAR(64) | 是 |  | 当前待审核人（提交时自动指派给待办最少的审核员；审核完成后由 ReviewerName 记录实际审核人） |
| `Status` | string | NVARCHAR(16) | 否 |  | 状态：待审核 / 已通过 / 已驳回 |
| `ReviewerName` | string | NVARCHAR(64) | 是 |  | 审核人 |
| `ReviewComment` | string | NVARCHAR(512) | 是 |  | 审核意见 |
| `CreatedAt` | DateTime | DATETIME | 否 |  |  |
| `ReviewedAt` | DateTime? | DATETIME | 是 |  |  |

## `customers`

客户实体（customers 表），数据源为计划表 Cust. Code 工作表的 B、C 列（Cust. Code, ENDCUSTOMER）

来源：`Models/Review.cs` → `Customer`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Name` | string | NVARCHAR(64) | 否 |  | 客户名称（ENDCUSTOMER） |
| `Code` | string | NVARCHAR(32) | 是 |  | 客户代码（Cust. Code） |
| `Remark` | string | NVARCHAR(256) | 是 |  |  |

## `test_items_catalog`

测试项目字典（test_items_catalog 表），数据源为计划表的"Test Items"工作表

来源：`Models/Review.cs` → `TestItemCatalog`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Name` | string | NVARCHAR(128) | 否 |  | 试验项目名 |
| `Period` | string | NVARCHAR(32) | 是 |  | 试验时间（小时） |
| `Category` | string | NVARCHAR(64) | 是 |  | 测试种类（RELIABILITY TEST / EMC / 不确定…）：报告 ORT Plan 与 TestStatus 的分组行按此归类， 计划索引时按历史报告归入的类别自动填，认不出来的归"不确定"，由用户在管理界面手工归类 |
| `Owner` | string | NVARCHAR(64) | 是 |  | 负责人（显示用文本，多个以"/"分隔；由 OwnerIds 对应的显示名同步维护，导入时存原始文本） |
| `OwnerIds` | string | NVARCHAR(256) | 是 |  | 负责人账号Id（多个以"/"分隔）——负责人以 uid 锚定技术员， 显示名修改后负责人栏展示随之更新，不会因改名而失去关联 |
| `Remark` | string | NVARCHAR(256) | 是 |  |  |

## `plan_item_templates`

测试项模板（plan_item_templates 表）：按测试项目名保存"多个机种重复出现"的公共文本， 各机种计划只在差异字段上另存（见 <see cref="TestPlanItem"/>），实现"重复内容只保存一次"。

来源：`Models/TestPlan.cs` → `PlanItemTemplate`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `TestItemName` | string | NVARCHAR(128) | 否 |  |  |
| `Category` | string | NVARCHAR(128) | 是 |  |  |
| `SamplingPlan` | string | TEXT | 是 |  |  |
| `TestCondition` | string | TEXT | 是 |  |  |
| `PassCriterion` | string | TEXT | 是 |  |  |
| `Remark` | string | TEXT | 是 |  |  |
| `Period` | string | NVARCHAR(64) | 是 |  |  |
| `UsageCount` | int | INTEGER | 否 |  |  |
| `IsManual` | bool | BOOLEAN | 否 |  |  |
| `CreatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `CreatedAt` | DateTime? | DATETIME | 是 |  |  |
| `UpdatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `UpdatedAt` | DateTime? | DATETIME | 是 |  |  |

## `test_plans`

测试计划（test_plans 表）：按「机种 + 阶段」规划要做哪些测试。 普通机种一般只有 MP 计划；新机种才有 NPI 计划。

来源：`Models/TestPlan.cs` → `TestPlan`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `ModelName` | string | NVARCHAR(128) | 否 |  |  |
| `Stage` | string | NVARCHAR(16) | 否 |  |  |
| `Remark` | string | TEXT | 是 |  |  |
| `Source` | string | NVARCHAR(16) | 是 |  |  |
| `CreatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `CreatedAt` | DateTime? | DATETIME | 是 |  |  |
| `UpdatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `UpdatedAt` | DateTime? | DATETIME | 是 |  |  |

## `test_plan_items`

测试计划明细（test_plan_items 表）：计划里的一项测试。 抽样计划/测试条件/通过判定/备注/周期 只在「与模板不同」时保存（为空表示沿用模板），

来源：`Models/TestPlan.cs` → `TestPlanItem`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `PlanId` | long | INTEGER | 否 |  |  |
| `OrderNo` | int | INTEGER | 否 |  |  |
| `Category` | string | NVARCHAR(128) | 是 |  |  |
| `TestItemName` | string | NVARCHAR(128) | 否 |  |  |
| `TemplateId` | long? | INTEGER | 是 |  |  |
| `SamplingPlan` | string | TEXT | 是 |  |  |
| `TestCondition` | string | TEXT | 是 |  |  |
| `PassCriterion` | string | TEXT | 是 |  |  |
| `Remark` | string | TEXT | 是 |  |  |
| `Period` | string | NVARCHAR(64) | 是 |  |  |
| `OverriddenFields` | string | NVARCHAR(256) | 是 |  |  |
| `Confirmed` | bool | BOOLEAN | 否 |  |  |
| `SourceVariants` | string | TEXT | 是 |  |  |
| `FromIndex` | bool | BOOLEAN | 否 |  | 是否由计划索引生成（重建索引时会被最新报告刷新或清理； 用户手工新增的明细为 false，重建时保留） |
| `UpdatedBy` | string | NVARCHAR(64) | 是 |  |  |
| `UpdatedAt` | DateTime? | DATETIME | 是 |  |  |

## `plan_index_jobs`

计划索引任务（plan_index_jobs 表）：一次「建立计划索引」的执行记录。 进度与待处理明细都落库，因此同一数据库上的任一客户端都能接着跑（断点继续）。

来源：`Models/TestPlan.cs` → `PlanIndexJob`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Status` | string | NVARCHAR(16) | 否 |  |  |
| `RootPath` | string | NVARCHAR(512) | 是 |  |  |
| `Total` | int | INTEGER | 否 |  |  |
| `Processed` | int | INTEGER | 否 |  |  |
| `Failed` | int | INTEGER | 否 |  |  |
| `Message` | string | NVARCHAR(512) | 是 |  |  |
| `StartedBy` | string | NVARCHAR(64) | 是 |  |  |
| `ClaimedBy` | string | NVARCHAR(128) | 是 |  |  |
| `ClaimedAt` | DateTime? | DATETIME | 是 |  |  |
| `StartedAt` | DateTime? | DATETIME | 是 |  |  |
| `FinishedAt` | DateTime? | DATETIME | 是 |  |  |
| `UpdatedAt` | DateTime? | DATETIME | 是 |  |  |

## `plan_index_entries`

计划索引明细（plan_index_entries 表）：一份报告 = 一条待处理明细。 每条独立记录状态与认领信息，客户端中途退出后其他客户端可从 Pending/超时认领的记录继续。

来源：`Models/TestPlan.cs` → `PlanIndexEntry`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `JobId` | long | INTEGER | 否 |  |  |
| `FolderName` | string | NVARCHAR(256) | 是 |  |  |
| `ModelName` | string | NVARCHAR(128) | 是 |  |  |
| `Stage` | string | NVARCHAR(16) | 是 |  |  |
| `OverviewFile` | string | NVARCHAR(512) | 是 |  |  |
| `Status` | string | NVARCHAR(16) | 否 |  |  |
| `ItemCount` | int | INTEGER | 否 |  |  |
| `Note` | string | NVARCHAR(1024) | 是 |  |  |
| `Error` | string | NVARCHAR(512) | 是 |  |  |
| `ClaimedBy` | string | NVARCHAR(128) | 是 |  |  |
| `ClaimedAt` | DateTime? | DATETIME | 是 |  |  |
| `UpdatedAt` | DateTime? | DATETIME | 是 |  |  |

## `plan_index_raw_items`

计划索引原始抽取结果（plan_index_raw_items 表）：从报告 ORT Plan 表里原样抽出的每项测试。 先落库再统一归并，归并过程可重复执行（重跑不会丢数据，也便于人工核对来源）。

来源：`Models/TestPlan.cs` → `PlanIndexRawItem`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `JobId` | long | INTEGER | 否 |  |  |
| `EntryId` | long | INTEGER | 否 |  |  |
| `ModelName` | string | NVARCHAR(128) | 是 |  |  |
| `Stage` | string | NVARCHAR(16) | 是 |  |  |
| `Category` | string | NVARCHAR(128) | 是 |  |  |
| `TestItemName` | string | NVARCHAR(128) | 是 |  |  |
| `OrderNo` | int | INTEGER | 否 |  |  |
| `SamplingPlan` | string | TEXT | 是 |  |  |
| `TestCondition` | string | TEXT | 是 |  |  |
| `PassCriterion` | string | TEXT | 是 |  |  |
| `Remark` | string | TEXT | 是 |  |  |

## `plan_item_images`

测试项目配图（plan_item_images 表）：计划索引时从历史报告的 ORT Plan 表里抽出图片， 按锚点所在行归到对应的测试项目上；生成新报告模板时把图片一起放进 ORT Plan 表 （历史报告里每个测试项目旁边都有设备/测试现场照片）。 图片文件落在数据库目录下的 PlanImages 文件夹，这里只存相对文件名。

来源：`Models/TestPlan.cs` → `PlanItemImage`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `NameKey` | string | NVARCHAR(128) | 否 |  |  |
| `TestItemName` | string | NVARCHAR(128) | 是 |  |  |
| `ModelName` | string | NVARCHAR(128) | 是 |  |  |
| `SourceFile` | string | NVARCHAR(512) | 是 |  |  |
| `FileName` | string | NVARCHAR(256) | 是 |  |  |
| `WidthPx` | int | INTEGER | 否 |  |  |
| `HeightPx` | int | INTEGER | 否 |  |  |
| `AnchorColumn` | int | INTEGER | 否 |  | 历史报告里这张图的锚点列（1 基，即"原来那一格"）： 生成 ORT Plan 时图片要放回这一格、位于该格文字下方 |
| `AnchorColumn2` | int | INTEGER | 否 |  |  |
| `OrderNo` | int | INTEGER | 否 |  |  |
| `UpdatedAt` | DateTime? | DATETIME | 是 |  |  |

## `users`

用户实体（users 表）。密码以 SHA256(Salt+密码) 散列存储。

来源：`Models/User.cs` → `User`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `Username` | string | NVARCHAR(64) | 否 |  | 登录名（唯一） |
| `DisplayName` | string | NVARCHAR(64) | 是 |  | 显示名称 |
| `Email` | string | NVARCHAR(128) | 是 |  | 邮箱（技术员/审核员登录时若为空会提示完善） |
| `PasswordHash` | string | NVARCHAR(128) | 否 |  | 密码散列 SHA256(Salt+密码) |
| `Salt` | string | NVARCHAR(64) | 否 |  | 密码盐 |
| `IsActive` | bool | BOOLEAN | 否 |  | 是否启用 |
| `CreatedAt` | DateTime | DATETIME | 否 |  |  |

## `user_roles`

用户-角色关联（user_roles 表）。一个用户可拥有多个身份。

来源：`Models/User.cs` → `UserRoleRow`

| 列 | C# 类型 | SQLite 类型 | 可空 | 主键 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `Id` | long | INTEGER | 否 | 是、自增 |  |
| `UserId` | long | INTEGER | 否 |  |  |
| `Role` | string | NVARCHAR(32) | 否 |  | 角色名（UserRole枚举的字符串形式：GeneralUser/Technician/Reviewer/Administrator） |

