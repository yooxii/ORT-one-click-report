using FreeSql.DataAnnotations;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 计划管理实体，以计划表为基底、领用表辅助，两表数据合并到同一行（一般通过工令与机种名称关联）。
    /// 三个业务唯一键：领料单据号 / 回线RT工令 / 工作編號。
    /// 空值一律存 NULL（SQLite 唯一索引允许多个 NULL，避免空串冲突）。
    /// </summary>
    [Table(Name = "plans")]
    [Index("uk_requisition_no", nameof(RequisitionNo), true)]
    [Index("uk_return_rt_order", nameof(ReturnRtOrder), true)]
    [Index("uk_job_no", nameof(JobNo), true)]
    public class Plan
    {
        /// <summary>
        /// 自增主键索引
        /// </summary>
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /* ------------------ 两表共用字段 ------------------ */

        /// <summary>
        /// 機種名稱 / 機種名/Part No
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get; set; }

        /// <summary>
        /// 測試項目 / 測試項目/Test Item
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string TestItem { get; set; }

        /// <summary>
        /// 备注 / 備 考/Remark
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Remark { get; set; }

        /// <summary>
        /// 序列号文本（可能包含以"/"分隔的多个SN）
        /// </summary>
        [Column(StringLength = 2048, IsNullable = true)]
        public string SN { get; set; }

        /// <summary>
        /// 序列号文件路径（OLE对象提取或上传的SN清单文件），相对于数据库Data目录
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string SnFilePath { get; set; }

        /* ------------------ 领用表字段 ------------------ */

        /// <summary>
        /// 領用/日期（原文如"1月9日"，保留以便原样导出）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string RequisitionDate { get; set; }

        /// <summary>
        /// 領用日期解析值（用于筛选/排序；年份从导入文件名推断）
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? RequisitionDateValue { get; set; }

        /// <summary>
        /// 領料單据號（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string RequisitionNo { get; set; }

        /// <summary>
        /// 領出/數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string OutQty { get; set; }

        /// <summary>
        /// D/C
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string DC { get; set; }

        /// <summary>
        /// REV.
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Rev { get; set; }

        /// <summary>
        /// Work Order
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string WorkOrder { get; set; }

        /// <summary>
        /// 回綫 RT 工令（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string ReturnRtOrder { get; set; }

        /// <summary>
        /// 回線/數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReturnQty { get; set; }

        /// <summary>
        /// 線別
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string LineNo { get; set; }

        /// <summary>
        /// 回線/日期（原文如"1月20日"）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReturnDate { get; set; }

        /// <summary>
        /// 回線日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? ReturnDateValue { get; set; }

        /// <summary>
        /// 入庫退料/單据號
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string StockInNo { get; set; }

        /// <summary>
        /// 入庫/數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StockInQty { get; set; }

        /// <summary>
        /// 入庫日期（原文如"1月21日"）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StockInDate { get; set; }

        /// <summary>
        /// 入庫日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? StockInDateValue { get; set; }

        /* ------------------ 计划表字段 ------------------ */

        /// <summary>
        /// 工作編號/Job No（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string JobNo { get; set; }

        /// <summary>
        /// 產品別/Product
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Product { get; set; }

        /// <summary>
        /// 客戶別/Customer
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Customer { get; set; }

        /// <summary>
        /// 階 段/Stage
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Stage { get; set; }

        /// <summary>
        /// 樣品數/Sample Size
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string SampleSize { get; set; }

        /// <summary>
        /// 試驗時間/Test Period
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string TestPeriod { get; set; }

        /// <summary>
        /// 負責人/Owner
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Owner { get; set; }

        /// <summary>
        /// 開始日期/Start Date（保留原文格式）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StartDate { get; set; }

        /// <summary>
        /// 開始日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? StartDateValue { get; set; }

        /// <summary>
        /// 結束日期/End Date（保留原文格式）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string EndDate { get; set; }

        /// <summary>
        /// 結束日期解析值
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? EndDateValue { get; set; }

        /// <summary>
        /// 完成狀況/Status（Close/Ongoing/Pending）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Status { get; set; }

        /// <summary>
        /// 上傳系統/Upload e-lab
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string UploadELab { get; set; }

        /* ------------------ 审计字段（预留权限/用户管理） ------------------ */

        /// <summary>
        /// 创建人（下一次任务接入用户管理后由登录用户填充）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string CreatedBy { get; set; }

        /// <summary>
        /// 创建时间
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? CreatedAt { get; set; }

        /// <summary>
        /// 最后修改人
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string UpdatedBy { get; set; }

        /// <summary>
        /// 最后修改时间
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? UpdatedAt { get; set; }
    }
}
