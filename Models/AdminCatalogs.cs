using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 代码映射（code_mappings 表）：Cust. Code 工作表的两位代码 → 名称。
    /// CodeType：C=客户别（B、C 列：Cust. Code→ENDCUSTOMER），P=产品别（G、H 列：Code→Product Type）。
    /// 查询规则：客户别 = 机种名称第 8 位起的 2 位；产品别 = 机种名称开始的 2 位。
    /// </summary>
    [Table(Name = "code_mappings")]
    [Index("uk_code_mapping", nameof(CodeType) + "," + nameof(Code), true)]
    public class CodeMapping
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 代码类型：C=客户别，P=产品别
        /// </summary>
        [Column(StringLength = 1, IsNullable = false)]
        public string CodeType { get; set; }

        /// <summary>
        /// 两位代码
        /// </summary>
        [Column(StringLength = 8, IsNullable = false)]
        public string Code { get; set; }

        /// <summary>
        /// 对应名称（客户名/产品类型名）
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Name { get; set; }
    }
    /// <summary>
    /// 阶段字典（stages 表）：阶段名 + 描述。初始值 MP/EVT/DVT/PVT/RMA
    /// </summary>
    [Table(Name = "stages")]
    [Index("uk_stage_name", nameof(Name), true)]
    public class Stage
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 阶段名（如 MP/EVT/DVT/PVT/RMA）
        /// </summary>
        [Column(StringLength = 32, IsNullable = false)]
        public string Name { get; set; }

        /// <summary>
        /// 描述
        /// </summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string Description { get; set; }
    }

    /// <summary>
    /// 产品别字典（products 表），数据源为计划表 Cust. Code 工作表的 G、H 列（Code, Product Type）
    /// </summary>
    [Table(Name = "products")]
    [Index("uk_product_name", nameof(Name), true)]
    public class Product
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        [Column(StringLength = 64, IsNullable = false)]
        public string Name { get; set; }

        /// <summary>
        /// 产品代码（Cust. Code 表 G 列）
        /// </summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Code { get; set; }

        [Column(StringLength = 256, IsNullable = true)]
        public string Remark { get; set; }
    }

    /// <summary>
    /// 机种映射（model_mappings 表）：还原计划表公式关系——
    /// 输入机种名称即可带出对应的产品别与客户别
    /// </summary>
    [Table(Name = "model_mappings")]
    [Index("uk_model_mapping_name", nameof(ModelName), true)]
    public class ModelMapping
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        [Column(StringLength = 128, IsNullable = false)]
        public string ModelName { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string Product { get; set; }

        [Column(StringLength = 64, IsNullable = true)]
        public string Customer { get; set; }
    }

    /// <summary>
    /// 计划数据变更日志（plan_change_logs 表）：记录每次提交的更改前后快照，便于追溯与回滚
    /// </summary>
    [Table(Name = "plan_change_logs")]
    public class PlanChangeLog
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 操作：新增 / 编辑 / 删除
        /// </summary>
        [Column(StringLength = 16, IsNullable = false)]
        public string Action { get; set; }

        /// <summary>
        /// 计划记录Id（新增时为提交后的Id）
        /// </summary>
        public long PlanId { get; set; }

        /// <summary>
        /// 变更摘要
        /// </summary>
        [Column(StringLength = 256, IsNullable = true)]
        public string Summary { get; set; }

        /// <summary>
        /// 变更前快照（JSON，新增时为null）
        /// </summary>
        [Column(DbType = "text", IsNullable = true)]
        public string BeforeJson { get; set; }

        /// <summary>
        /// 变更后快照（JSON，删除时为null）
        /// </summary>
        [Column(DbType = "text", IsNullable = true)]
        public string AfterJson { get; set; }

        /// <summary>
        /// 操作人
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Operator { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}
