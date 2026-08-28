using ORT一键报告.Models;
using System;
using System.Collections.Generic;

namespace ORT一键报告.Reports.Models
{
    /// <summary>
    /// 测试结果状态（与原 WindowMainReport.ReportStatus 保持一致，统一放在模型层）
    /// </summary>
    public enum ReportStatus { Pass, Fail }


    /// <summary>
    /// 一键报告输入模型基类。每种报告一个具体派生类。
    /// </summary>
    public abstract class ReportInputModel
    {
        /// <summary>
        /// 报告类型名称（与模板/Tab 对应）：Thermal Shock / Burn In / EMI
        /// </summary>
        public abstract string ReportType { get; }

        /// <summary>
        /// 测试周期（天），用于根据 TestStart 推算 TestEnd 默认值
        /// </summary>
        public abstract int TestTimeDays { get; }

        /// <summary>
        /// 表头信息（所有报告共有）
        /// </summary>
        public ReportHeaderData Header { get; set; } = new();

        /// <summary>
        /// 报告 Excel 模板文件路径（必填，生成时使用此模板作为基础）
        /// </summary>
        public string TemplatePath { get; set; }

        /// <summary>
        /// 报告概览所在根目录（用于解析相对路径/默认保存位置）
        /// </summary>
        public string RootReportPath { get; set; }

        /// <summary>
        /// 临时目录（处理中间文件时使用）
        /// </summary>
        public string TempPath { get; set; }
    }

    /// <summary>
    /// Thermal Shock 报告输入模型
    /// </summary>
    public class ThermalShockReportModel : ReportInputModel
    {
        public override string ReportType => "Thermal Shock";
        public override int TestTimeDays => 1;

        /// <summary>
        /// 来源 UUT 数据（读取自报告概览）
        /// </summary>
        public UUTSourceData UUTSource { get; set; } = new();

        /// <summary>
        /// 每个 SN 的测试结果明细
        /// </summary>
        public List<ResultDetailItem> Details { get; set; } = [];

        /// <summary>
        /// ATE 数据文件路径（将作为 OLE 嵌入到报告）
        /// </summary>
        public string AteDataFilePath { get; set; }
    }

    /// <summary>
    /// Burn In 报告输入模型
    /// </summary>
    public class BurnInReportModel : ReportInputModel
    {
        public override string ReportType => "Burn In";
        public override int TestTimeDays => 7;

        public UUTSourceData UUTSource { get; set; } = new();
        public List<ResultDetailItem> Details { get; set; } = [];
        public string AteDataFilePath { get; set; }
    }

    /// <summary>
    /// EMI 报告输入模型（Conducted EMI Measurement）
    /// </summary>
    public class EMIReportModel : ReportInputModel
    {
        public override string ReportType => "EMI";
        public override int TestTimeDays => 1;

        /// <summary>
        /// 来源 UUT 数据（WorkOrder/Revision/DC 等基本信息）
        /// </summary>
        public UUTSourceData UUTSource { get; set; } = new();

        /// <summary>
        /// EMI 数据目录（含 PDF/DOCX 源数据文件）
        /// </summary>
        public string EmiDataDir { get; set; }

        /// <summary>
        /// 参与报告的 DOCX 源文件列表（从 EmiDataDir 解析得到）
        /// </summary>
        public List<string> EmiDocxFiles { get; set; } = [];

        /// <summary>
        /// EMI 被测单元的组合选择（SN/Voltage/Load/LISN）
        /// </summary>
        public EMIUnitSelection UnitSelection { get; set; } = new();
    }

    /// <summary>
    /// 报告表头共享字段
    /// </summary>
    public class ReportHeaderData
    {
        public string TestedBy { get; set; }
        public string ApprovedBy { get; set; }
        public string ProjectName { get; set; }
        public string TestStage { get; set; }
        public string TestDescription { get; set; }
        public DateTime TestStart { get; set; }
        public DateTime TestEnd { get; set; }
        public bool TestPass { get; set; } = true;

        /// <summary>
        /// 问题现象图片（最多 3 张）
        /// </summary>
        public List<ImageAttachment> IssuePhotos { get; set; } = [];

        /// <summary>
        /// 测试布置图片（最多 3 张）
        /// </summary>
        public List<ImageAttachment> TestSetupPhotos { get; set; } = [];

        /// <summary>
        /// 测试描述附图
        /// </summary>
        public ImageAttachment TestDescriptionPhoto { get; set; }

        /// <summary>
        /// 测试 ATE 数据附图（如有）
        /// </summary>
        public ImageAttachment AteDataPhoto { get; set; }
    }

    /// <summary>
    /// 被测单元来源数据（读取自报告概览 Excel 的 UUT 部分）
    /// </summary>
    public class UUTSourceData
    {
        /// <summary>
        /// 序列号列表
        /// </summary>
        public List<string> SNs { get; set; } = [];

        public string WorkOrder { get; set; }
        public string Revision { get; set; }
        public string DC { get; set; }

        /// <summary>
        /// 测试项目列表（名称 + 日期）
        /// </summary>
        public List<TestItemInfo> TestItems { get; set; } = [];
    }

    /// <summary>
    /// ThermalShock/BurnIn 报告每个 SN 的结果明细
    /// </summary>
    public class ResultDetailItem
    {
        public string BIroom { get; set; } = "";
        public string BIarea { get; set; } = "";
        public string BIplace { get; set; } = "";
        public string SN { get; set; } = "";
        public string WorkOrder { get; set; } = "";
        public string Version { get; set; } = "";
        public string DC { get; set; } = "";
        public ReportStatus InspectionPrev { get; set; }
        public ReportStatus FunPrev { get; set; }
        public ReportStatus InspectionAfter { get; set; }
        public ReportStatus FunAfter { get; set; }
        public ReportStatus HiPot { get; set; }
        public string Comments { get; set; } = "";
    }

    /// <summary>
    /// EMI 被测单元的组合选择
    /// </summary>
    public class EMIUnitSelection
    {
        public List<string> SNs { get; set; } = [];
        public List<string> Voltages { get; set; } = [];
        public List<string> Loads { get; set; } = [];
        public List<string> LISNs { get; set; } = [];
    }

    /// <summary>
    /// 图片附件（纯数据：字节数组 + 名称/路径）
    /// </summary>
    public class ImageAttachment
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public byte[] Bytes { get; set; }
    }
}
