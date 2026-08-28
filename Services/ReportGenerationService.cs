using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ORT一键报告.Models;
using ORT一键报告.Reports.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 一键报告生成服务：以 ReportInputModel 派生类作为唯一输入，
    /// 在指定输出路径生成对应 Excel 报告。
    /// 与 UI 解耦，之后只要提供该模型的实例即可直接生成报告。
    ///
    /// 用法：
    ///   var model = new BurnInReportModel { Header = ..., Details = ..., AteDataFilePath = ... };
    ///   service.Generate(model, @"C:\Reports\output.xlsx");
    /// </summary>
    public class ReportGenerationService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ReportGenerationService() { }

        /// <summary>
        /// 统一入口：根据模型类型派发到对应生成器
        /// </summary>
        public void Generate(ReportInputModel model, string outputPath)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("输出路径不能为空", nameof(outputPath));
            if (string.IsNullOrWhiteSpace(model.TemplatePath) || !File.Exists(model.TemplatePath))
            {
                throw new FileNotFoundException($"报告模板不存在: {model.TemplatePath}");
            }

            switch (model)
            {
                case ThermalShockReportModel ts:
                    GenerateBaseReport(ts, outputPath);
                    break;
                case BurnInReportModel bi:
                    GenerateBaseReport(bi, outputPath);
                    break;
                case EMIReportModel emi:
                    GenerateEMIReport(emi, outputPath);
                    break;
                default:
                    throw new NotSupportedException($"不支持的报告类型: {model.ReportType}");
            }
        }

        /* ###############################  ThermalShock / Burn In 生成  ################################ */

        /// <summary>
        /// 生成 ThermalShock 或 BurnIn 报告（基于模板写入表头/单体数据/图片/ATE OLE）
        /// </summary>
        private void GenerateBaseReport(ReportInputModel model, string outputPath)
        {
            _logger.Info($"{model.ReportType} 报告生成中...");

            // 1. 打开模板
            FileInfo templateFile = new(model.TemplatePath);
            using ExcelPackage package = new(templateFile);
            ExcelWorkbook wb = package.Workbook;
            ExcelWorksheet ws = wb.Worksheets[0];
            ExcelWorksheet ws_setup = wb.Worksheets[1];

            // 2. 写入表头信息（按 setup 表定义的 8 个字段地址映射）
            _logger.Info("处理表头");
            string[] headerValues =
            [
                model.Header.TestedBy,
                model.Header.ApprovedBy,
                model.Header.ProjectName,
                model.Header.TestStage,
                model.Header.TestStart.ToString("d"),
                model.Header.TestEnd.ToString("d"),
                model.Header.TestPass ? "Pass" : "Fail",
                model.Header.TestDescription,
            ];
            for (int r = 1; r <= 8; r++)
            {
                ws.Cells[ws_setup.Cells[r, 1].Text].Value = headerValues[r - 1];
            }

            // 3. 写入单体数据（Burn In 包含 BIroom/area/place，ThermalShock 跳过）
            _logger.Info("处理单体数据");
            List<object> detailInfoList = [];
            bool isBurnIn = model.ReportType.ToLower().Contains("burn");

            if (isBurnIn)
            {
                IEnumerable<ResultDetailItem> details = GetDetails(model);
                detailInfoList.Add(details.Select(d => d.BIroom).ToList());
                detailInfoList.Add(details.Select(d => d.BIarea).ToList());
                detailInfoList.Add(details.Select(d => d.BIplace).ToList());
            }

            IEnumerable<ResultDetailItem> allDetails = GetDetails(model);
            detailInfoList.Add(allDetails.Select(d => d.SN).ToList());
            detailInfoList.Add(allDetails.Select(d => d.WorkOrder).ToList());
            detailInfoList.Add(allDetails.Select(d => d.Version).ToList());
            detailInfoList.Add(allDetails.Select(d => d.DC).ToList());
            detailInfoList.Add(allDetails.Select(d => d.InspectionPrev).ToList());
            detailInfoList.Add(allDetails.Select(d => d.InspectionAfter).ToList());
            detailInfoList.Add(allDetails.Select(d => d.FunPrev).ToList());
            detailInfoList.Add(allDetails.Select(d => d.FunAfter).ToList());
            detailInfoList.Add(allDetails.Select(d => d.HiPot).ToList());

            int detailStartRow = 13; // setup 表 detail 起始行
            for (int r = detailStartRow; r < ws_setup.Dimension.End.Row; r++)
            {
                ExcelAddress address = new(ws_setup.Cells[r, 1].Text);
                int targetRow = address.Start.Row;
                int targetCol = address.Start.Column;
                int idx = r - detailStartRow;
                if (idx >= detailInfoList.Count) continue;

                if (detailInfoList[idx] is List<string> strs)
                {
                    for (int i = 0; i < strs.Count; i++)
                    {
                        ws.Cells[targetRow + i, targetCol].Value = strs[i];
                    }
                }
                else if (detailInfoList[idx] is List<ReportStatus> statuses)
                {
                    for (int i = 0; i < statuses.Count; i++)
                    {
                        ws.Cells[targetRow + i, targetCol].Value = statuses[i].ToString();
                    }
                }
            }

            // 4. 写入图片和 OLE 对象
            _logger.Info("处理图片和OLE对象");
            string tempPath = model.TempPath ?? Path.Combine(Path.GetTempPath(), "ORTTemp");
            ExcelAddPicture(ws, "Issue_Photos", ToLegacyDataCell(model.Header.IssuePhotos), ws_setup.Cells["A11"].Text, model.ReportType, tempPath);
            ExcelAddPicture(ws, "Test_Setup", ToLegacyDataCell(model.Header.TestSetupPhotos), ws_setup.Cells["A12"].Text, model.ReportType, tempPath);

            string ateAddr = ws_setup.Cells["A9"].Text;
            string atePath = GetAtePath(model);
            if (!string.IsNullOrWhiteSpace(atePath))
            {
                EmbedOleObjectWithEpplus(ws, atePath, ateAddr);
            }

            // 5. 删除 setup 表并保存
            wb.Worksheets.Delete(ws_setup);
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
            package.SaveAs(outputPath);

            _logger.Info($"{model.ReportType} 报告生成完成: {outputPath}");
        }

        /// <summary>
        /// 从 ThermalShock/BurnIn 模型提取 Details
        /// </summary>
        private static IEnumerable<ResultDetailItem> GetDetails(ReportInputModel model) => model switch
        {
            ThermalShockReportModel ts => ts.Details,
            BurnInReportModel bi => bi.Details,
            _ => []
        };

        /// <summary>
        /// 提取 ATE 数据文件路径
        /// </summary>
        private static string GetAtePath(ReportInputModel model) => model switch
        {
            ThermalShockReportModel ts => ts.AteDataFilePath,
            BurnInReportModel bi => bi.AteDataFilePath,
            _ => null
        };

        /// <summary>
        /// 将纯数据图片列表转换为旧版 ExcelPictureInfo 承载的 DataCell
        /// </summary>
        private static DataCell ToLegacyDataCell(List<ImageAttachment> images)
        {
            if (images == null || images.Count == 0) return null;
            DataCell cell = new();
            cell.Images = images.Select(a => new ExcelPictureInfo
            {
                Name = a.Name,
                ImageBytes = a.Bytes
            }).ToList();
            return cell;
        }

        /* ###############################  EMI 报告生成  ################################ */

        /// <summary>
        /// 生成 EMI (Conducted EMI Measurement) 报告。
        /// TODO: EMI 生成逻辑较复杂（电压/负载/LISN 组合解析、图表绘制、ZIP 嵌入）。
        /// 当前保留现有 UI 入口（EMIReportViewModel.DoReport）生成，待后续把核心逻辑迁入本服务。
        /// </summary>
        private void GenerateEMIReport(EMIReportModel model, string outputPath)
        {
            _logger.Info($"{model.ReportType} 报告生成中...");

            // TODO: 将 EMIReportViewModel.DoReport 的核心逻辑迁入此方法，以支持模型化生成。
            // 目前仍由 UI 层 EMIReportViewModel 处理，本服务仅作为占位。
            throw new NotImplementedException(
                $"EMI 报告模型化生成尚未接入，请通过 UI 的 EMI Tab 生成报告。\n"
                + $"（待迁移：EMIReportViewModel.DoReport → ReportGenerationService.GenerateEMIReport）");
        }
    }
}
