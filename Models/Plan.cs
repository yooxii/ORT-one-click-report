using CommunityToolkit.Mvvm.ComponentModel;
using FreeSql.DataAnnotations;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 计划表实体（plans 表）：ORT Test Schedule 数据。
    /// 与领退表（requisitions 表）分表存储。
    /// 工作編號唯一（QRT 前缀为非领用计划，RT 前缀为正常领用计划）。
    /// 实现属性通知以便表格内编辑/机种联动自动带出时单元格即时刷新。
    /// </summary>
    [Table(Name = "plans")]
    [Index("uk_job_no", nameof(JobNo), true)]
    public class Plan : ObservableObject
    {
        /// <summary>
        /// 自增主键索引
        /// </summary>
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        private string _modelName;
        /// <summary>
        /// 機種名/Part No
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get => _modelName; set => SetProperty(ref _modelName, value); }

        private string _testItem;
        /// <summary>
        /// 測試項目/Test Item
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string TestItem { get => _testItem; set => SetProperty(ref _testItem, value); }

        private string _remark;
        /// <summary>
        /// 備 考/Remark
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Remark { get => _remark; set => SetProperty(ref _remark, value); }

        private string _jobNo;
        /// <summary>
        /// 工作編號/Job No（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string JobNo { get => _jobNo; set => SetProperty(ref _jobNo, value); }

        private string _product;
        /// <summary>
        /// 產品別/Product
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Product { get => _product; set => SetProperty(ref _product, value); }

        private string _customer;
        /// <summary>
        /// 客戶別/Customer
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Customer { get => _customer; set => SetProperty(ref _customer, value); }

        private string _stage;
        /// <summary>
        /// 階 段/Stage
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Stage { get => _stage; set => SetProperty(ref _stage, value); }

        private string _sampleSize;
        /// <summary>
        /// 樣品數/Sample Size
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string SampleSize { get => _sampleSize; set => SetProperty(ref _sampleSize, value); }

        private string _testPeriod;
        /// <summary>
        /// 試驗時間/Test Period
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string TestPeriod { get => _testPeriod; set => SetProperty(ref _testPeriod, value); }

        private string _owner;
        /// <summary>
        /// 負責人/Owner
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Owner { get => _owner; set => SetProperty(ref _owner, value); }

        private System.DateTime? _startDate;
        /// <summary>
        /// 開始日期/Start Date（日期类型）
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? StartDate { get => _startDate; set => SetProperty(ref _startDate, value); }

        private System.DateTime? _endDate;
        /// <summary>
        /// 結束日期/End Date（日期类型）
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? EndDate { get => _endDate; set => SetProperty(ref _endDate, value); }

        private string _status;
        /// <summary>
        /// 完成狀況/Status（Close/Ongoing/Pending）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Status { get => _status; set => SetProperty(ref _status, value); }

        private string _uploadELab;
        /// <summary>
        /// 上傳系統/Upload e-lab
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string UploadELab { get => _uploadELab; set => SetProperty(ref _uploadELab, value); }

        /* ------------------ 审计字段 ------------------ */

        private string _createdBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get => _createdBy; set => SetProperty(ref _createdBy, value); }

        private System.DateTime? _createdAt;
        [Column(IsNullable = true)]
        public System.DateTime? CreatedAt { get => _createdAt; set => SetProperty(ref _createdAt, value); }

        private string _updatedBy;
        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get => _updatedBy; set => SetProperty(ref _updatedBy, value); }

        private System.DateTime? _updatedAt;
        [Column(IsNullable = true)]
        public System.DateTime? UpdatedAt { get => _updatedAt; set => SetProperty(ref _updatedAt, value); }

        private bool _hasReportLink;
        /// <summary>
        /// 工作编号是否存在对应的报告文件夹（仅用于界面颜色区分，不参与数据库存储与快照对比）
        /// </summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public bool HasReportLink { get => _hasReportLink; set => SetProperty(ref _hasReportLink, value); }
    }
}
