using CommunityToolkit.Mvvm.ComponentModel;
using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 预领·月排程明细（pre_schedule_items 表）。
    /// 原始工作表是「年月 / 線別 / 1日 / 2日 … 31日」的平铺矩阵，一个单元格写着「機種/數量」；
    /// 落库时拆成一行一个（月份 + 線別 + 機種 + 日），这样周视图、整月视图、变动比对都能直接查。
    /// 数据按月归档：按月整体替换（导入更新），历史月份保留可翻阅。
    /// </summary>
    [Table(Name = "pre_schedule_items")]
    [Index("idx_pre_schedule_month", nameof(Month), false)]
    public class PreScheduleItem
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>年月（yyyyMM，例如 202611）</summary>
        [Column(StringLength = 6, IsNullable = true)]
        public string Month { get; set; }

        /// <summary>線別（例如 P01、P/R）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Line { get; set; }

        /// <summary>機種</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Model { get; set; }

        /// <summary>日（1..31，配合 Month 组成完整日期）</summary>
        public int Day { get; set; }

        /// <summary>預計生產數量</summary>
        public int Qty { get; set; }

        /// <summary>原始单元格文本（「機種/數量」）；数量解析不出来时保留原样便于核对</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string RawText { get; set; }

        /// <summary>来源文件名（更新记录里用）</summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string SourceFile { get; set; }

        public DateTime? ImportedAt { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string ImportedBy { get; set; }
    }

    /// <summary>
    /// 预领·ORT 看板行（ort_board_rows 表）。
    /// 只保留原始导出表 A~O 共 15 列（P 列及其后是人工加的列，不导入）；
    /// 原始导出的第 1 行不是表头，**表头在第 2 行**，导入时从第 2 行读表头、第 3 行起读数据。
    /// 月份取自「日期」列（yyyymmdd → yyyyMM）。
    /// </summary>
    [Table(Name = "ort_board_rows")]
    [Index("idx_ort_board_month", nameof(Month), false)]
    public class OrtBoardRow : ObservableObject
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>月份（yyyyMM，由「日期」列推出）</summary>
        [Column(StringLength = 6, IsNullable = true)]
        public string Month { get; set; }

        /// <summary>日期（原始文本，通常是 yyyymmdd）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReportDate { get; set; }

        /// <summary>類別</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Category { get; set; }

        /// <summary>機種</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Model { get; set; }

        /// <summary>生產線別</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Line { get; set; }

        /// <summary>生產日期</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ProduceDate { get; set; }

        /// <summary>工令</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string JobNo { get; set; }

        /// <summary>生產數量</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Qty { get; set; }

        /// <summary>生產時數</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Hours { get; set; }

        /// <summary>65報工否</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Reported65 { get; set; }

        /// <summary>SIM87入庫否（更新变动主要看这一列：N → Y 表示可以领用了）</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Stocked87
        {
            get => _stocked87;
            set
            {
                if (SetProperty(ref _stocked87, value))
                {
                    OnPropertyChanged(nameof(Stocked));
                }
            }
        }
        private string _stocked87;

        /// <summary>收料日期</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReceiveDate { get; set; }

        /// <summary>料號</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string PartNo { get; set; }

        /// <summary>材料剩餘庫存</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string MaterialStock { get; set; }

        /// <summary>測試結果</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string TestResult { get; set; }

        /// <summary>SIM94禁止出貨</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Forbidden94 { get; set; }

        /// <summary>是否已经转成「准备领用」的待开单记录（界面配色用，随库存储）</summary>
        public bool HasRequisition
        {
            get => _hasRequisition;
            set
            {
                if (SetProperty(ref _hasRequisition, value))
                {
                    OnPropertyChanged(nameof(Stocked87Hint));
                }
            }
        }
        private bool _hasRequisition;

        /// <summary>看板上的机种是否已入库（SIM87入庫否 = Y），只用于界面配色</summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public bool Stocked => PreIssueCodes.IsYes(Stocked87);

        /// <summary>行状态提示（配色 / 悬浮提示用）</summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public string Stocked87Hint => Stocked ? (HasRequisition ? "已转待开单" : "已入库·待转") : "";

        public DateTime? ImportedAt { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string ImportedBy { get; set; }

        /// <summary>比对用的键：工令优先，没有工令时用「機種|生產日期」</summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public string MatchKey => !string.IsNullOrWhiteSpace(JobNo)
            ? JobNo.Trim()
            : $"{Model}|{ProduceDate}";
    }

    /// <summary>
    /// 预领·准备领用（pre_issue_requisitions 表），对应 ORT 领用。
    /// 两个阶段：
    /// 1) 待开单：先只填机种名（从 ORT 看板 SIM87入庫否=Y 的机种推过来，或手工添加）；
    /// 2) 已开单：下载开单数据（11.XLS 一类）后导入，按 VCHNO 或机种补全 10 列；
    ///    再次导入时主要更新 VCH_STATUS、CLOSE_CODE 两列，其余列也一并同步变动并记录。
    /// </summary>
    [Table(Name = "pre_issue_requisitions")]
    [Index("idx_pre_issue_req_month", nameof(Month), false)]
    public class PreIssueRequisition : ObservableObject
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>月份（yyyyMM）</summary>
        [Column(StringLength = 6, IsNullable = true)]
        public string Month { get; set; }

        /// <summary>VCHNO（领用单号；待开单阶段为空）</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Vchno
        {
            get => _vchno;
            set
            {
                if (SetProperty(ref _vchno, value))
                {
                    OnPropertyChanged(nameof(PhaseText));
                }
            }
        }
        private string _vchno;

        /// <summary>WARE_CODE</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string WareCode { get; set; }

        /// <summary>TR_CODE</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string TrCode { get; set; }

        /// <summary>PART_NO（機種）</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string PartNo { get; set; }

        /// <summary>REQ_QTY（领用数量）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReqQty { get; set; }

        /// <summary>IVER_ID</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string IverId { get; set; }

        /// <summary>VCH_STATUS（导入更新时重点跟踪的一列）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string VchStatus
        {
            get => _vchStatus;
            set
            {
                if (SetProperty(ref _vchStatus, value))
                {
                    OnPropertyChanged(nameof(PhaseText));
                }
            }
        }
        private string _vchStatus;

        /// <summary>ISSUE_DEPT</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string IssueDept { get; set; }

        /// <summary>CLOSE_CODE（导入更新时重点跟踪的一列）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string CloseCode { get; set; }

        /// <summary>PC</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Pc { get; set; }

        /// <summary>阶段：待开单 / 已开单（见 <see cref="PreIssueCodes"/>）</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string State { get; set; }

        /// <summary>关联的看板工令（从看板推进来时带上，便于回溯）</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string JobNo { get; set; }

        /// <summary>阶段（待开单 / 已开单）：看有没有开单号即可，界面上单独一列显示</summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public string PhaseText => PreIssueCodes.IsPendingState(State) || string.IsNullOrWhiteSpace(Vchno)
            ? PreIssueCodes.PendingState
            : PreIssueCodes.OpenedState;

        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get; set; }

        public DateTime? CreatedAt { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get; set; }

        public DateTime? UpdatedAt { get; set; }
    }

    /// <summary>
    /// 预领·ORT 黑名单（ort_blacklist 表）：有些机种 ORT 做不了，不需要领用，
    /// 看板变动时不再推进到「准备领用」。
    /// </summary>
    [Table(Name = "ort_blacklist")]
    [Index("uk_ort_blacklist_model", nameof(Model), true)]
    public class OrtBlacklistItem
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>机种（唯一）</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Model { get; set; }

        /// <summary>加入原因 / 备注</summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string Reason { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get; set; }

        public DateTime? CreatedAt { get; set; }
    }

    /// <summary>
    /// 预领·变动记录（pre_issue_change_logs 表）：
    /// 三个表每次导入更新都逐条记录变动（新增 / 删除 / 修改 / 入库转领用），
    /// 旧月份的数据被更新时同样记录，便于随时翻阅。
    /// </summary>
    [Table(Name = "pre_issue_change_logs")]
    [Index("idx_pre_issue_change_month", nameof(Month), false)]
    public class PreIssueChangeLog
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>月份（yyyyMM）</summary>
        [Column(StringLength = 6, IsNullable = true)]
        public string Month { get; set; }

        /// <summary>来源表（月排程 / ORT看板 / 准备领用 / 黑名单，见 <see cref="PreIssueCodes"/>）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Source { get; set; }

        /// <summary>变动类型（新增 / 删除 / 修改 / 导入 / 入库 / 开单 / 黑名单）</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Kind { get; set; }

        /// <summary>变动对象（機種 / 工令 / 單號）</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string Target { get; set; }

        /// <summary>变动的列 / 字段名（新增删除时为空）</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Field { get; set; }

        /// <summary>变动前</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string OldValue { get; set; }

        /// <summary>变动后</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string NewValue { get; set; }

        /// <summary>补充说明（来源文件等）</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Note { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string Operator { get; set; }

        public DateTime? CreatedAt { get; set; }
    }

    /// <summary>
    /// 预领模块用到的固定文案/常量（存库值不随界面语言变化）
    /// </summary>
    public static class PreIssueCodes
    {
        /// <summary>准备领用：待开单</summary>
        public const string PendingState = "待开单";

        /// <summary>准备领用：已开单</summary>
        public const string OpenedState = "已开单";

        /// <summary>变动记录来源：月排程</summary>
        public const string SourceSchedule = "月排程";

        /// <summary>变动记录来源：ORT看板</summary>
        public const string SourceBoard = "ORT看板";

        /// <summary>变动记录来源：准备领用</summary>
        public const string SourceRequisition = "准备领用";

        /// <summary>变动记录来源：黑名单</summary>
        public const string SourceBlacklist = "黑名单";

        /// <summary>变动类型：新增</summary>
        public const string KindAdded = "新增";

        /// <summary>变动类型：删除</summary>
        public const string KindRemoved = "删除";

        /// <summary>变动类型：修改</summary>
        public const string KindModified = "修改";

        /// <summary>变动类型：导入（一次导入的摘要）</summary>
        public const string KindImported = "导入";

        /// <summary>变动类型：入库（SIM87入庫否变成 Y）</summary>
        public const string KindStocked = "入库";

        /// <summary>变动类型：开单（待开单 → 已开单）</summary>
        public const string KindOpened = "开单";

        /// <summary>变动类型：黑名单</summary>
        public const string KindBlacklist = "黑名单";

        /// <summary>按字面判断「是」（Y / y / 是 / Y是）</summary>
        public static bool IsYes(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            string v = value.Trim();
            return v.Equals("Y", StringComparison.OrdinalIgnoreCase)
                || v.Equals("YES", StringComparison.OrdinalIgnoreCase)
                || v.Equals("是", StringComparison.Ordinal);
        }

        /// <summary>是否待开单阶段</summary>
        public static bool IsPendingState(string state) =>
            string.IsNullOrWhiteSpace(state) || state.Trim() == PendingState;
    }
}
