using CommunityToolkit.Mvvm.ComponentModel;
using FreeSql.DataAnnotations;
using System;
using System.Collections.Generic;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 计划阶段取值：量产（MP）与新机种（NPI）。
    /// 报告 Cover 里的 Product Stage 若是 EVT/DVT/PVT 等试产阶段，索引时统一归并为 NPI。
    /// </summary>
    public static class PlanStage
    {
        /// <summary>量产</summary>
        public const string MP = "MP";

        /// <summary>新机种</summary>
        public const string NPI = "NPI";

        /// <summary>把报告里的阶段文本归并为 MP / NPI（认不出时按量产处理）</summary>
        public static string Normalize(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return MP;
            }
            string text = raw.Trim().ToUpperInvariant();
            if (text.Contains("NPI") || text.Contains("EVT") || text.Contains("DVT") || text.Contains("PVT")
                || text.Contains("PROTO") || text.Contains("新機種") || text.Contains("新机种"))
            {
                return NPI;
            }
            return MP;
        }

        /// <summary>阶段显示名</summary>
        public static string Display(string stage)
            => string.Equals(stage, NPI, StringComparison.OrdinalIgnoreCase) ? "新机种" : "量产";
    }

    /// <summary>
    /// 测试项模板（plan_item_templates 表）：按测试项目名保存"多个机种重复出现"的公共文本，
    /// 各机种计划只在差异字段上另存（见 <see cref="TestPlanItem"/>），实现"重复内容只保存一次"。
    /// </summary>
    [Table(Name = "plan_item_templates")]
    [Index("uk_plan_item_template", nameof(TestItemName), true)]
    public class PlanItemTemplate : ObservableObject
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        private string _testItemName;
        /// <summary>测试项目名（唯一；与测试项目字典登记的名称对应）</summary>
        [Column(StringLength = 128, IsNullable = false)]
        public string TestItemName { get => _testItemName; set => SetProperty(ref _testItemName, value); }

        private string _category;
        /// <summary>分类（如 RELIABILITY TEST / EMC），用于 ORT Plan 与 TestStatus 的分组行</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string Category { get => _category; set => SetProperty(ref _category, value); }

        private string _samplingPlan;
        /// <summary>抽样计划（公共文本）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string SamplingPlan { get => _samplingPlan; set => SetProperty(ref _samplingPlan, value); }

        private string _testCondition;
        /// <summary>测试条件（公共文本）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string TestCondition { get => _testCondition; set => SetProperty(ref _testCondition, value); }

        private string _passCriterion;
        /// <summary>通过判定（公共文本）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string PassCriterion { get => _passCriterion; set => SetProperty(ref _passCriterion, value); }

        private string _remark;
        /// <summary>备注（公共文本）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string Remark { get => _remark; set => SetProperty(ref _remark, value); }

        private string _period;
        /// <summary>试验周期（如 7天 / 168H），供 Waterfall 排测试安排使用</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Period { get => _period; set => SetProperty(ref _period, value); }

        private int _usageCount;
        /// <summary>采用该模板的报告份数（索引统计，仅供界面参考）</summary>
        public int UsageCount { get => _usageCount; set => SetProperty(ref _usageCount, value); }

        private bool _isManual;
        /// <summary>是否由用户手工维护（索引重建时不覆盖其文本，只更新统计）</summary>
        public bool IsManual { get => _isManual; set => SetProperty(ref _isManual, value); }

        private string _createdBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get => _createdBy; set => SetProperty(ref _createdBy, value); }

        private DateTime? _createdAt;
        [Column(IsNullable = true)]
        public DateTime? CreatedAt { get => _createdAt; set => SetProperty(ref _createdAt, value); }

        private string _updatedBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get => _updatedBy; set => SetProperty(ref _updatedBy, value); }

        private DateTime? _updatedAt;
        [Column(IsNullable = true)]
        public DateTime? UpdatedAt { get => _updatedAt; set => SetProperty(ref _updatedAt, value); }
    }

    /// <summary>
    /// 测试计划（test_plans 表）：按「机种 + 阶段」规划要做哪些测试。
    /// 普通机种一般只有 MP 计划；新机种才有 NPI 计划。
    /// </summary>
    [Table(Name = "test_plans")]
    [Index("uk_test_plan", nameof(ModelName) + "," + nameof(Stage), true)]
    public class TestPlan : ObservableObject
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        private string _modelName;
        /// <summary>机种名称</summary>
        [Column(StringLength = 128, IsNullable = false)]
        public string ModelName { get => _modelName; set => SetProperty(ref _modelName, value); }

        private string _stage;
        /// <summary>阶段：MP=量产，NPI=新机种</summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Stage { get => _stage; set => SetProperty(ref _stage, value); }

        private string _remark;
        /// <summary>计划备注（如 ORT Plan 表末的 Note 说明）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string Remark { get => _remark; set => SetProperty(ref _remark, value); }

        private string _source;
        /// <summary>来源：Index=计划索引生成，Manual=手工建立</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Source { get => _source; set => SetProperty(ref _source, value); }

        private string _createdBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get => _createdBy; set => SetProperty(ref _createdBy, value); }

        private DateTime? _createdAt;
        [Column(IsNullable = true)]
        public DateTime? CreatedAt { get => _createdAt; set => SetProperty(ref _createdAt, value); }

        private string _updatedBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get => _updatedBy; set => SetProperty(ref _updatedBy, value); }

        private DateTime? _updatedAt;
        [Column(IsNullable = true)]
        public DateTime? UpdatedAt { get => _updatedAt; set => SetProperty(ref _updatedAt, value); }

        /// <summary>阶段显示名（不落库）</summary>
        [Column(IsIgnore = true)]
        public string StageDisplay => PlanStage.Display(Stage);

        /// <summary>明细条数（不落库，由服务填充）</summary>
        [Column(IsIgnore = true)]
        public int ItemCount { get; set; }

        /// <summary>差异条数（不落库，由服务填充）</summary>
        [Column(IsIgnore = true)]
        public int OverrideCount { get; set; }
    }

    /// <summary>
    /// 测试计划明细（test_plan_items 表）：计划里的一项测试。
    /// 抽样计划/测试条件/通过判定/备注/周期 只在「与模板不同」时保存（为空表示沿用模板），
    /// <see cref="OverriddenFields"/> 记录哪些字段是差异，便于界面提示与人工确认。
    /// </summary>
    [Table(Name = "test_plan_items")]
    public class TestPlanItem : ObservableObject
    {
        /// <summary>五个可差异字段的名称（与 <see cref="OverriddenFields"/> 中使用的一致）</summary>
        public const string FieldSamplingPlan = "SamplingPlan";
        public const string FieldTestCondition = "TestCondition";
        public const string FieldPassCriterion = "PassCriterion";
        public const string FieldRemark = "Remark";
        public const string FieldPeriod = "Period";

        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>所属计划Id</summary>
        public long PlanId { get; set; }

        private int _orderNo;
        /// <summary>顺序号（计划内从 1 开始）</summary>
        public int OrderNo { get => _orderNo; set => SetProperty(ref _orderNo, value); }

        private string _category;
        /// <summary>分类（RELIABILITY TEST / EMC …）</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string Category { get => _category; set => SetProperty(ref _category, value); }

        private string _testItemName;
        /// <summary>测试项目名（必须在测试项目字典中登记）</summary>
        [Column(StringLength = 128, IsNullable = false)]
        public string TestItemName { get => _testItemName; set => SetProperty(ref _testItemName, value); }

        private long? _templateId;
        /// <summary>引用的测试项模板Id（为空表示未关联模板，五个字段全是自有值）</summary>
        [Column(IsNullable = true)]
        public long? TemplateId { get => _templateId; set => SetProperty(ref _templateId, value); }

        private string _samplingPlan;
        /// <summary>抽样计划差异（为空=沿用模板）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string SamplingPlan { get => _samplingPlan; set => SetProperty(ref _samplingPlan, value); }

        private string _testCondition;
        /// <summary>测试条件差异（为空=沿用模板）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string TestCondition { get => _testCondition; set => SetProperty(ref _testCondition, value); }

        private string _passCriterion;
        /// <summary>通过判定差异（为空=沿用模板）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string PassCriterion { get => _passCriterion; set => SetProperty(ref _passCriterion, value); }

        private string _remark;
        /// <summary>备注差异（为空=沿用模板）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string Remark { get => _remark; set => SetProperty(ref _remark, value); }

        private string _period;
        /// <summary>试验周期差异（为空=沿用模板）</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Period { get => _period; set => SetProperty(ref _period, value); }

        private string _overriddenFields;
        /// <summary>差异字段列表（逗号分隔），如 "TestCondition,PassCriterion"</summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string OverriddenFields { get => _overriddenFields; set => SetProperty(ref _overriddenFields, value); }

        private bool _confirmed;
        /// <summary>差异是否已人工确认（计划索引会把未确认的差异列为待确认清单）</summary>
        public bool Confirmed { get => _confirmed; set => SetProperty(ref _confirmed, value); }

        private string _sourceVariants;
        /// <summary>索引发现的其他写法（换行分隔，仅作人工确认时的参考）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string SourceVariants { get => _sourceVariants; set => SetProperty(ref _sourceVariants, value); }

        private bool _fromIndex;
        /// <summary>
        /// 是否由计划索引生成（重建索引时会被最新报告刷新或清理；
        /// 用户手工新增的明细为 false，重建时保留）
        /// </summary>
        public bool FromIndex { get => _fromIndex; set => SetProperty(ref _fromIndex, value); }

        private string _updatedBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get => _updatedBy; set => SetProperty(ref _updatedBy, value); }

        private DateTime? _updatedAt;
        [Column(IsNullable = true)]
        public DateTime? UpdatedAt { get => _updatedAt; set => SetProperty(ref _updatedAt, value); }

        /* ------------------ 展示用（不落库） ------------------ */

        /// <summary>引用的模板（不落库，由服务装配）</summary>
        [Column(IsIgnore = true)]
        public PlanItemTemplate Template { get; set; }

        /// <summary>是否有差异（不落库）</summary>
        [Column(IsIgnore = true)]
        public bool HasOverride => !string.IsNullOrWhiteSpace(OverriddenFields);

        /// <summary>差异字段的显示文本（不落库）</summary>
        [Column(IsIgnore = true)]
        public string OverriddenFieldsDisplay
        {
            get
            {
                if (!HasOverride)
                {
                    return "";
                }
                List<string> names = [];
                foreach (string fieldName in OverriddenFields.Split([','], StringSplitOptions.RemoveEmptyEntries))
                {
                    names.Add(fieldName switch
                    {
                        FieldSamplingPlan => "抽样计划",
                        FieldTestCondition => "测试条件",
                        FieldPassCriterion => "通过判定",
                        FieldRemark => "备注",
                        FieldPeriod => "周期",
                        _ => fieldName.Trim()
                    });
                }
                return string.Join("、", names);
            }
        }

        /// <summary>实际生效的抽样计划（差异优先，否则用模板）</summary>
        [Column(IsIgnore = true)]
        public string EffectiveSamplingPlan => string.IsNullOrEmpty(SamplingPlan) ? Template?.SamplingPlan : SamplingPlan;

        /// <summary>实际生效的测试条件</summary>
        [Column(IsIgnore = true)]
        public string EffectiveTestCondition => string.IsNullOrEmpty(TestCondition) ? Template?.TestCondition : TestCondition;

        /// <summary>实际生效的通过判定</summary>
        [Column(IsIgnore = true)]
        public string EffectivePassCriterion => string.IsNullOrEmpty(PassCriterion) ? Template?.PassCriterion : PassCriterion;

        /// <summary>实际生效的备注</summary>
        [Column(IsIgnore = true)]
        public string EffectiveRemark => string.IsNullOrEmpty(Remark) ? Template?.Remark : Remark;

        /// <summary>实际生效的试验周期</summary>
        [Column(IsIgnore = true)]
        public string EffectivePeriod => string.IsNullOrEmpty(Period) ? Template?.Period : Period;

        /// <summary>
        /// 按差异字段列表与取值刷新 <see cref="OverriddenFields"/>（值为空表示不再差异）
        /// </summary>
        public void RefreshOverriddenFields()
        {
            List<string> fields = [];
            if (!string.IsNullOrEmpty(SamplingPlan)) fields.Add(FieldSamplingPlan);
            if (!string.IsNullOrEmpty(TestCondition)) fields.Add(FieldTestCondition);
            if (!string.IsNullOrEmpty(PassCriterion)) fields.Add(FieldPassCriterion);
            if (!string.IsNullOrEmpty(Remark)) fields.Add(FieldRemark);
            if (!string.IsNullOrEmpty(Period)) fields.Add(FieldPeriod);
            OverriddenFields = fields.Count == 0 ? null : string.Join(",", fields);
            OnPropertyChanged(nameof(HasOverride));
            OnPropertyChanged(nameof(OverriddenFieldsDisplay));
            OnPropertyChanged(nameof(EffectiveSamplingPlan));
            OnPropertyChanged(nameof(EffectiveTestCondition));
            OnPropertyChanged(nameof(EffectivePassCriterion));
            OnPropertyChanged(nameof(EffectiveRemark));
            OnPropertyChanged(nameof(EffectivePeriod));
        }
    }

    /// <summary>
    /// 计划索引任务（plan_index_jobs 表）：一次「建立计划索引」的执行记录。
    /// 进度与待处理明细都落库，因此同一数据库上的任一客户端都能接着跑（断点继续）。
    /// </summary>
    [Table(Name = "plan_index_jobs")]
    public class PlanIndexJob : ObservableObject
    {
        /// <summary>待执行（含已建好明细、等待认领）</summary>
        public const string StatusPending = "Pending";

        /// <summary>执行中</summary>
        public const string StatusRunning = "Running";

        /// <summary>已暂停（用户主动暂停）</summary>
        public const string StatusPaused = "Paused";

        /// <summary>全部完成并已归并</summary>
        public const string StatusDone = "Done";

        /// <summary>失败</summary>
        public const string StatusFailed = "Failed";

        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        private string _status;
        /// <summary>状态：Pending/Running/Paused/Done/Failed</summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Status { get => _status; set => SetProperty(ref _status, value); }

        private string _rootPath;
        /// <summary>扫描的报告根目录（同一目录复用同一个任务，便于断点继续）</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string RootPath { get => _rootPath; set => SetProperty(ref _rootPath, value); }

        private int _total;
        /// <summary>报告明细总数</summary>
        public int Total { get => _total; set => SetProperty(ref _total, value); }

        private int _processed;
        /// <summary>已处理数</summary>
        public int Processed { get => _processed; set => SetProperty(ref _processed, value); }

        private int _failed;
        /// <summary>失败数</summary>
        public int Failed { get => _failed; set => SetProperty(ref _failed, value); }

        private string _message;
        /// <summary>最近一次状态说明</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Message { get => _message; set => SetProperty(ref _message, value); }

        private string _startedBy;
        /// <summary>发起人</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string StartedBy { get => _startedBy; set => SetProperty(ref _startedBy, value); }

        private string _claimedBy;
        /// <summary>当前认领该任务的客户端标识（多客户端互斥）</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ClaimedBy { get => _claimedBy; set => SetProperty(ref _claimedBy, value); }

        private DateTime? _claimedAt;
        /// <summary>认领时间（超时后可被其他客户端接管）</summary>
        [Column(IsNullable = true)]
        public DateTime? ClaimedAt { get => _claimedAt; set => SetProperty(ref _claimedAt, value); }

        private DateTime? _startedAt;
        [Column(IsNullable = true)]
        public DateTime? StartedAt { get => _startedAt; set => SetProperty(ref _startedAt, value); }

        private DateTime? _finishedAt;
        [Column(IsNullable = true)]
        public DateTime? FinishedAt { get => _finishedAt; set => SetProperty(ref _finishedAt, value); }

        private DateTime? _updatedAt;
        [Column(IsNullable = true)]
        public DateTime? UpdatedAt { get => _updatedAt; set => SetProperty(ref _updatedAt, value); }

        /// <summary>进度文本（不落库）</summary>
        [Column(IsIgnore = true)]
        public string ProgressText => Total <= 0 ? "尚未建立索引" : $"{Processed}/{Total}" + (Failed > 0 ? $"（失败 {Failed}）" : "");
    }

    /// <summary>
    /// 计划索引明细（plan_index_entries 表）：一份报告 = 一条待处理明细。
    /// 每条独立记录状态与认领信息，客户端中途退出后其他客户端可从 Pending/超时认领的记录继续。
    /// </summary>
    [Table(Name = "plan_index_entries")]
    public class PlanIndexEntry
    {
        /// <summary>待处理</summary>
        public const string StatusPending = "Pending";

        /// <summary>处理中（被某客户端认领）</summary>
        public const string StatusRunning = "Running";

        /// <summary>已完成</summary>
        public const string StatusDone = "Done";

        /// <summary>失败</summary>
        public const string StatusFailed = "Failed";

        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>所属任务Id</summary>
        public long JobId { get; set; }

        /// <summary>报告夹名称（用于界面显示）</summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string FolderName { get; set; }

        /// <summary>解析出的机种名称</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get; set; }

        /// <summary>解析出的阶段（MP/NPI）</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Stage { get; set; }

        /// <summary>报告概览 Excel 文件完整路径</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string OverviewFile { get; set; }

        /// <summary>处理状态：Pending/Running/Done/Failed</summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Status { get; set; }

        /// <summary>抽取到的测试项条数</summary>
        public int ItemCount { get; set; }

        /// <summary>ORT Plan 表末的 Note 说明（Sampling principle 等），归并时作为计划备注</summary>
        [Column(StringLength = 1024, IsNullable = true)]
        public string Note { get; set; }

        /// <summary>失败原因</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Error { get; set; }

        /// <summary>认领该明细的客户端标识</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ClaimedBy { get; set; }

        /// <summary>认领时间</summary>
        [Column(IsNullable = true)]
        public DateTime? ClaimedAt { get; set; }

        /// <summary>更新时间</summary>
        [Column(IsNullable = true)]
        public DateTime? UpdatedAt { get; set; }
    }

    /// <summary>
    /// 计划索引原始抽取结果（plan_index_raw_items 表）：从报告 ORT Plan 表里原样抽出的每项测试。
    /// 先落库再统一归并，归并过程可重复执行（重跑不会丢数据，也便于人工核对来源）。
    /// </summary>
    [Table(Name = "plan_index_raw_items")]
    public class PlanIndexRawItem
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>所属任务Id</summary>
        public long JobId { get; set; }

        /// <summary>来源明细Id</summary>
        public long EntryId { get; set; }

        /// <summary>机种名称</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get; set; }

        /// <summary>阶段（MP/NPI）</summary>
        [Column(StringLength = 16, IsNullable = true)]
        public string Stage { get; set; }

        /// <summary>分类</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string Category { get; set; }

        /// <summary>测试项目名</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string TestItemName { get; set; }

        /// <summary>顺序号</summary>
        public int OrderNo { get; set; }

        /// <summary>抽样计划（原文）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string SamplingPlan { get; set; }

        /// <summary>测试条件（原文）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string TestCondition { get; set; }

        /// <summary>通过判定（原文）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string PassCriterion { get; set; }

        /// <summary>备注（原文）</summary>
        [Column(DbType = "text", IsNullable = true)]
        public string Remark { get; set; }
    }

    /// <summary>
    /// 测试项目配图（plan_item_images 表）：计划索引时从历史报告的 ORT Plan 表里抽出图片，
    /// 按锚点所在行归到对应的测试项目上；生成新报告模板时把图片一起放进 ORT Plan 表
    /// （历史报告里每个测试项目旁边都有设备/测试现场照片）。
    /// 图片文件落在数据库目录下的 PlanImages 文件夹，这里只存相对文件名。
    /// </summary>
    [Table(Name = "plan_item_images")]
    public class PlanItemImage
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>测试项目名的归一化键（见 PlanIndexService.NameKey）</summary>
        [Column(StringLength = 128, IsNullable = false)]
        public string NameKey { get; set; }

        /// <summary>测试项目名（原样，便于人工核对）</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string TestItemName { get; set; }

        /// <summary>来源机种</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get; set; }

        /// <summary>来源报告概览文件</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string SourceFile { get; set; }

        /// <summary>图片文件名（相对 PlanImages 目录）</summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string FileName { get; set; }

        /// <summary>图片宽度（像素）</summary>
        public int WidthPx { get; set; }

        /// <summary>图片高度（像素）</summary>
        public int HeightPx { get; set; }

        /// <summary>同一测试项目内的顺序</summary>
        public int OrderNo { get; set; }

        [Column(IsNullable = true)]
        public DateTime? UpdatedAt { get; set; }
    }
}
