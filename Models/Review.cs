using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 审核请求实体（review_requests 表），工作流形式：
    /// 请求方提交（如计划表单更改），审核员通过（应用更改）或驳回。
    /// </summary>
    [Table(Name = "review_requests")]
    public class ReviewRequest
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 请求类型（当前仅"计划表单"，预留"报告"类型）
        /// </summary>
        [Column(StringLength = 32, IsNullable = false)]
        public string Type { get; set; }

        /// <summary>
        /// 操作类型：新增 / 编辑 / 删除
        /// </summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Action { get; set; }

        /// <summary>
        /// 目标记录Id（编辑/删除时有值）
        /// </summary>
        [Column(IsNullable = true)]
        public long? TargetId { get; set; }

        /// <summary>
        /// 请求摘要（列表展示用）
        /// </summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string Summary { get; set; }

        /// <summary>
        /// 更改内容（Plan 序列化 JSON）
        /// </summary>
        [Column(DbType = "text", IsNullable = true)]
        public string PayloadJson { get; set; }

        /// <summary>
        /// 请求人
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string RequesterName { get; set; }

        /// <summary>
        /// 当前待审核人（提交时自动指派给待办最少的审核员；审核完成后由 ReviewerName 记录实际审核人）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string AssigneeName { get; set; }

        /// <summary>
        /// 状态：待审核 / 已通过 / 已驳回
        /// </summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Status { get; set; } = "待审核";

        /// <summary>
        /// 审核人
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string ReviewerName { get; set; }

        /// <summary>
        /// 审核意见
        /// </summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string ReviewComment { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        [Column(IsNullable = true)]
        public DateTime? ReviewedAt { get; set; }
    }

    /// <summary>
    /// 客户实体（customers 表），数据源为计划表 Cust. Code 工作表的 B、C 列（Cust. Code, ENDCUSTOMER）
    /// </summary>
    [Table(Name = "customers")]
    [Index("uk_customer_name", nameof(Name), true)]
    public class Customer
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 客户名称（ENDCUSTOMER）
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Name { get; set; }

        /// <summary>
        /// 客户代码（Cust. Code）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Code { get; set; }

        [Column(StringLength = 256, IsNullable = true)]
        public string Remark { get; set; }
    }

    /// <summary>
    /// 测试项目字典（test_items_catalog 表），数据源为计划表的"Test Items"工作表
    /// </summary>
    [Table(Name = "test_items_catalog")]
    [Index("uk_testitem_name", nameof(Name), true)]
    public class TestItemCatalog
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 试验项目名
        /// </summary>
        [Column(StringLength = 128, IsNullable = false)]
        public string Name { get; set; }

        /// <summary>
        /// 试验时间（小时）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Period { get; set; }

        /// <summary>
        /// 负责人（显示用文本，多个以"/"分隔；由 OwnerIds 对应的显示名同步维护，导入时存原始文本）
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string Owner { get; set; }

        /// <summary>
        /// 负责人账号Id（多个以"/"分隔）——负责人以 uid 锚定技术员，
        /// 显示名修改后负责人栏展示随之更新，不会因改名而失去关联
        /// </summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string OwnerIds { get; set; }

        /// <summary>
        /// 负责人显示名（按 OwnerIds 解析，不落库，仅供界面展示）
        /// </summary>
        [Column(IsIgnore = true)]
        public string OwnerDisplay { get; set; }

        [Column(StringLength = 256, IsNullable = true)]
        public string Remark { get; set; }
    }
}
