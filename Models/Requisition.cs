using CommunityToolkit.Mvvm.ComponentModel;
using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 领退表实体（requisitions 表）：记录成品領用/回线信息。
    /// 与计划表（plans 表）分表存储，通过领料单据号/WorkOrder/回线RT工令关联。
    /// </summary>
    [Table(Name = "requisitions")]
    [Index("uk_req_requisition_no", nameof(RequisitionNo), true)]
    public class Requisition : ObservableObject
    {
        /// <summary>
        /// 自增主键索引
        /// </summary>
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        private System.DateTime? _requisitionDate;
        /// <summary>
        /// 領用日期（必填，日期类型）
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? RequisitionDate { get => _requisitionDate; set => SetProperty(ref _requisitionDate, value); }

        private string _requisitionNo;
        /// <summary>
        /// 領料單据號（必填，唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string RequisitionNo { get => _requisitionNo; set => SetProperty(ref _requisitionNo, value); }

        private string _modelName;
        /// <summary>
        /// 機種名稱（必填）
        /// </summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string ModelName { get => _modelName; set => SetProperty(ref _modelName, value); }

        private string _outQty;
        /// <summary>
        /// 領出數量（必填）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string OutQty { get => _outQty; set => SetProperty(ref _outQty, value); }

        private string _sn;
        /// <summary>
        /// S/N（必填，字符串或附件形式）
        /// </summary>
        [Column(StringLength = 2048, IsNullable = true)]
        public string SN { get => _sn; set => SetProperty(ref _sn, value); }

        private string _snFilePath;
        /// <summary>
        /// SN附件文件路径（S/N 为附件形式时）
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string SnFilePath { get => _snFilePath; set => SetProperty(ref _snFilePath, value); }

        private string _rev;
        /// <summary>
        /// REV.（必填）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Rev { get => _rev; set => SetProperty(ref _rev, value); }

        private string _workOrder;
        /// <summary>
        /// Work Order（必填）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string WorkOrder { get => _workOrder; set => SetProperty(ref _workOrder, value); }

        private string _dc;
        /// <summary>
        /// D/C（自动补全：WorkOrder 倒数第三位起的两位表示第多少周）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string DC { get => _dc; set => SetProperty(ref _dc, value); }

        private string _lineNo;
        /// <summary>
        /// 線別（自动补全：WorkOrder 倒数第六位起的三位字符串）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string LineNo { get => _lineNo; set => SetProperty(ref _lineNo, value); }

        private string _returnRtOrder;
        /// <summary>
        /// 回线RT工令（可选自动生成：RTAH{当前年月}{编号}）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string ReturnRtOrder { get => _returnRtOrder; set => SetProperty(ref _returnRtOrder, value); }

        private string _returnQty;
        /// <summary>
        /// 回線數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string ReturnQty { get => _returnQty; set => SetProperty(ref _returnQty, value); }

        private System.DateTime? _returnDate;
        /// <summary>
        /// 回線日期
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? ReturnDate { get => _returnDate; set => SetProperty(ref _returnDate, value); }

        private string _stockInNo;
        /// <summary>
        /// 入庫退料單据號
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string StockInNo { get => _stockInNo; set => SetProperty(ref _stockInNo, value); }

        private string _stockInQty;
        /// <summary>
        /// 入庫數量
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string StockInQty { get => _stockInQty; set => SetProperty(ref _stockInQty, value); }

        private System.DateTime? _stockInDate;
        /// <summary>
        /// 入庫日期
        /// </summary>
        [Column(IsNullable = true)]
        public System.DateTime? StockInDate { get => _stockInDate; set => SetProperty(ref _stockInDate, value); }

        private string _remark;
        /// <summary>
        /// 备注
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string Remark { get => _remark; set => SetProperty(ref _remark, value); }

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
    }
}
