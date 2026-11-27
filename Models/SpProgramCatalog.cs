using System.Collections.Generic;

namespace ORT一键报告.Models
{
    /// <summary>SP程序集里一个条目的类型</summary>
    public enum SpProgramKind
    {
        /// <summary>本程序自己的窗口（例如回线工令登记）</summary>
        Internal,

        /// <summary>外部的公司程式，直接启动它自己的 exe</summary>
        External,

        /// <summary>先占位、还没接入</summary>
        Placeholder,
    }

    /// <summary>
    /// SP程序集里的一个条目：只描述「是什么」，具体怎么打开由界面层决定。
    /// 名称/说明用本地化键，路径有默认值、可在本机设置里覆盖（不同厂区目录不一样）。
    /// </summary>
    public sealed class SpProgramInfo
    {
        /// <summary>稳定 id（也用作本机设置里路径覆盖的键）</summary>
        public string Id { get; set; }

        /// <summary>名称本地化键</summary>
        public string NameKey { get; set; }

        /// <summary>说明本地化键</summary>
        public string DescriptionKey { get; set; }

        /// <summary>条目类型</summary>
        public SpProgramKind Kind { get; set; }

        /// <summary>外部程式的默认完整路径（内部窗口为空）</summary>
        public string DefaultPath { get; set; }
    }

    /// <summary>
    /// SP程序集清单。这里集中登记公司 MES 的常用程式，
    /// 界面上点一下就能打开，不用再去翻 C:\actelap、C:\acbelap 里的 exe。
    /// 以后新接入的程式加一条即可（外部程式填 exe 路径；我们自己做的窗口填 Internal）。
    /// </summary>
    public static class SpProgramCatalog
    {
        /// <summary>回线工令登记（本程序窗口）</summary>
        public const string Sim44Id = "sim44";

        /// <summary>各条目的稳定 id</summary>
        public const string OqcId = "oqc";
        public const string InventoryId = "inventory";
        public const string MoWorkOrderId = "moworkorder";
        public const string PreIssueId = "preissue";

        /// <summary>全部条目（顺序即界面顺序）</summary>
        public static IReadOnlyList<SpProgramInfo> All { get; } =
        [
            new SpProgramInfo
            {
                Id = Sim44Id,
                NameKey = "SpPrograms_Sim44_Name",
                DescriptionKey = "SpPrograms_Sim44_Desc",
                Kind = SpProgramKind.Internal,
            },
            new SpProgramInfo
            {
                Id = OqcId,
                NameKey = "SpPrograms_Oqc_Name",
                DescriptionKey = "SpPrograms_Oqc_Desc",
                Kind = SpProgramKind.External,
                DefaultPath = @"C:\actelap\C_SII51.EXE",
            },
            new SpProgramInfo
            {
                Id = InventoryId,
                NameKey = "SpPrograms_Inventory_Name",
                DescriptionKey = "SpPrograms_Inventory_Desc",
                Kind = SpProgramKind.Placeholder,
                DefaultPath = @"C:\acbelap\C_IVI33.EXE",
            },
            new SpProgramInfo
            {
                Id = MoWorkOrderId,
                NameKey = "SpPrograms_MoWorkOrder_Name",
                DescriptionKey = "SpPrograms_MoWorkOrder_Desc",
                Kind = SpProgramKind.Placeholder,
                DefaultPath = @"C:\acbelap\C_IVI540_BL.EXE",
            },
            new SpProgramInfo
            {
                Id = PreIssueId,
                NameKey = "SpPrograms_PreIssue_Name",
                DescriptionKey = "SpPrograms_PreIssue_Desc",
                Kind = SpProgramKind.Placeholder,
                DefaultPath = @"C:\acbelap\C_IVR621_BL.EXE",
            },
        ];
    }
}
