using CommunityToolkit.Mvvm.ComponentModel;
using FreeSql.DataAnnotations;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 计划管理实体，以计划表为基底、领用表辅助，两表数据合并到同一行（一般通过工令与机种名称关联）。
    /// 三个业务唯一键：领料单据号 / 回线RT工令 / 工作編號。
    /// 空值一律存 NULL（SQLite 唯一索引允许多个 NULL，避免空串冲突）。
    /// 实现属性通知以便表格内编辑/机种联动自动带出时单元格即时刷新。
    /// </summary>
    [Table(Name = "plans")]
    [Index("uk_requisition_no", nameof(RequisitionNo), true)]
    [Index("uk_return_rt_order", nameof(ReturnRtOrder), true)]
    [Index("uk_job_no", nameof(JobNo), true)]
    public class Plan : ObservableObject
    {
        /// <summary>
        /// 自增主键索引
        /// </summary>
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /* ------------------ 两表共用字段 ------------------ */

        private string _modelName;
        /// <summary>
        /// 機種名稱 / 機種名/Part No
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get => _modelName; set => SetProperty(ref _modelName, value); }

        private string _testItem;
        /// <summary>
        /// 測試項目 / 測試項目/Test Item
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string TestItem { get => _testItem; set => SetProperty(ref _testItem, value); }

        private string _remark;
        /// <summary>
        /// 备注 / 備 考/Remark
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Remark { get => _remark; set => SetProperty(ref _remark, value); }

        private string _sn;
        /// <summary>
        /// 序列号文本（可能包含以"/"分隔的多个SN）
        /// </summary>
        [Column(StringLength = 2048, IsNullable = true)]
        public string SN { get => _sn; set => SetProperty(ref _sn, value); }

        private string _snFilePath;
        /// <summary>
        /// 序列号文件路径（OLE对象提取或上传的SN清单文件），相对于数据库Data目录
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string SnFilePath { get => _snFilePath; set => SetProperty(ref _snFilePath, value); }

        /* ------------------ 领用表字段 ------------------ */

        private string _requisitionDate;
        /// <summary>
        /// 領用/日期（原文如"1月9日"，保留以便原样导出）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string RequisitionDate { get => _requisitionDate; set => SetProperty(ref _requisitionDate, value); }

        private System.DateTime? _requisitionDateValue;
        /// <summary>
        /// 領用日期解析值（用于筛选/排序；年份从导入文件名推断）
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? RequisitionDateValue { get => _requisitionDateValue; set => SetProperty(ref _requisitionDateValue, value); }

        private string _requisitionNo;
        /// <summary>
        /// 領料單据號（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string RequisitionNo { get => _requisitionNo; set => SetProperty(ref _requisitionNo, value); }

        private string _outQty;
        /// <summary>
        /// 領出/數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string OutQty { get => _outQty; set => SetProperty(ref _outQty, value); }

        private string _dc;
        /// <summary>
        /// D/C
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string DC { get => _dc; set => SetProperty(ref _dc, value); }

        private string _rev;
        /// <summary>
        /// REV.
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Rev { get => _rev; set => SetProperty(ref _rev, value); }

        private string _workOrder;
        /// <summary>
        /// Work Order
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string WorkOrder { get => _workOrder; set => SetProperty(ref _workOrder, value); }

        private string _returnRtOrder;
        /// <summary>
        /// 回綫 RT 工令（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string ReturnRtOrder { get => _returnRtOrder; set => SetProperty(ref _returnRtOrder, value); }

        private string _returnQty;
        /// <summary>
        /// 回線/數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReturnQty { get => _returnQty; set => SetProperty(ref _returnQty, value); }

        private string _lineNo;
        /// <summary>
        /// 線別
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string LineNo { get => _lineNo; set => SetProperty(ref _lineNo, value); }

        private string _returnDate;
        /// <summary>
        /// 回線/日期（原文如"1月20日"）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReturnDate { get => _returnDate; set => SetProperty(ref _returnDate, value); }

        private System.DateTime? _returnDateValue;
        /// <summary>
        /// 回線日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? ReturnDateValue { get => _returnDateValue; set => SetProperty(ref _returnDateValue, value); }

        private string _stockInNo;
        /// <summary>
        /// 入庫退料/單据號
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string StockInNo { get => _stockInNo; set => SetProperty(ref _stockInNo, value); }

        private string _stockInQty;
        /// <summary>
        /// 入庫/數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StockInQty { get => _stockInQty; set => SetProperty(ref _stockInQty, value); }

        private string _stockInDate;
        /// <summary>
        /// 入庫日期（原文如"1月21日"）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StockInDate { get => _stockInDate; set => SetProperty(ref _stockInDate, value); }

        private System.DateTime? _stockInDateValue;
        /// <summary>
        /// 入庫日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? StockInDateValue { get => _stockInDateValue; set => SetProperty(ref _stockInDateValue, value); }

        /* ------------------ 计划表字段 ------------------ */

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

        private string _startDate;
        /// <summary>
        /// 開始日期/Start Date（保留原文格式）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StartDate { get => _startDate; set => SetProperty(ref _startDate, value); }

        private System.DateTime? _startDateValue;
        /// <summary>
        /// 開始日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? StartDateValue { get => _startDateValue; set => SetProperty(ref _startDateValue, value); }

        private string _endDate;
        /// <summary>
        /// 結束日期/End Date（保留原文格式）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string EndDate { get => _endDate; set => SetProperty(ref _endDate, value); }

        private System.DateTime? _endDateValue;
        /// <summary>
        /// 結束日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? EndDateValue { get => _endDateValue; set => SetProperty(ref _endDateValue, value); }

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

        /* ------------------ 审计字段（预留权限/用户管理） ------------------ */

        private string _createdBy;
        /// <summary>
        /// 创建人（下一次任务接入用户管理后由登录用户填充）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get => _createdBy; set => SetProperty(ref _createdBy, value); }

        private System.DateTime? _createdAt;
        /// <summary>
        /// 创建时间
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? CreatedAt { get => _createdAt; set => SetProperty(ref _createdAt, value); }

        private string _updatedBy;
        /// <summary>
        /// 最后修改人
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get => _updatedBy; set => SetProperty(ref _updatedBy, value); }

        private System.DateTime? _updatedAt;
        /// <summary>
        /// 最后修改时间
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? UpdatedAt { get => _updatedAt; set => SetProperty(ref _updatedAt, value); }
    }
}
