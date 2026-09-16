using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ORT一键报告.Models;
using ORT一键报告.Reports.Views;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using ORT一键报告.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Reports.ViewModels
{
    //2.1 Conducted EMI Measurement
    public partial class EMIReportViewModel : ObservableObject
    {
        private readonly IPathService _emiService;
        private readonly ReportService _reportService;
        private readonly AppSettingsService _appSettings;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly EMIUUTdataInfo emiUUTdatasInfo = new();
        private readonly List<string> emiDocxFiles = [];
        private readonly List<string> emiPdfFiles = [];

        public ReportHeaderViewModel ReportHeaderVM { get; set; }
        public EMISetupViewModel EMISetupVM { get; set; }

        // Word 表格解析后的数值列下标。实测行布局为：
        // [0]标记号 [1]频率 [2]QP实测 [3]QP限值 [4]QP余量 [5]AVG实测 [6]AVG限值 [7]AVG余量 [8]…
        // 其中"余量"= 限值 − 实测；挑最严苛的一行就是取余量最小的那行。
        /// <summary>标记号列</summary>
        public int MarkNo_Col { set; get; } = 0;
        /// <summary>频率列</summary>
        public int Freq_Col { set; get; } = 1;
        /// <summary>QP 实测值列</summary>
        public int PK_Col { set; get; } = 2;
        /// <summary>QP 限值列</summary>
        public int PK_Limit_Col { set; get; } = 3;
        /// <summary>QP 余量列（原命名 TolerableLimit，实际是"限值−实测"的余量）</summary>
        public int PK_Margin_Col { set; get; } = 4;
        /// <summary>AVG 实测值列</summary>
        public int AVG_Col { set; get; } = 5;
        /// <summary>AVG 限值列</summary>
        public int AVG_Limit_Col { set; get; } = 6;
        /// <summary>AVG 余量列（同上）</summary>
        public int AVG_Margin_Col { set; get; } = 7;

        private string _templatePath = string.Empty;

        public string TemplatePath
        {
            get => _templatePath;
            set => SetProperty(ref _templatePath, value);
        }

        private string _dataPath = string.Empty;

        public string DataPath
        {
            get => _dataPath;
            set
            {
                if (SetProperty(ref _dataPath, value))
                {
                    // 命令可能还没被界面访问过（懒加载），这里必须空安全，否则赋值早于绑定时会抛空引用
                    toPDFCommand?.RaiseCanExecuteChanged();
                    alertTimeCommand?.RaiseCanExecuteChanged();
                }
            }
        }

        private string _dc;
        public string DC
        {
            get => _dc;
            set => SetProperty(ref _dc, value);
        }

        private string _workOrder;
        public string WorkOrder
        {
            get => _workOrder;
            set => SetProperty(ref _workOrder, value);
        }

        private string _version;
        public string Version
        {
            get => _version;
            set => SetProperty(ref _version, value);
        }


        private string _remark = "Remark: \n1.  Q.P. and AV. are abbreviations of quasi-peak and average individually.\n2.  “-”This value have no tested, according to standard GB 9254-2008 Annex B, If the peak value under average limit, then not need to measure the QP and AV value.\n3.  Margin value= Read Value – Limit value";
        public string Remark
        {
            get => _remark;
            set => SetProperty(ref _remark, value);
        }

        public EMIReportViewModel(IPathService service, ReportService reportService, AppSettingsService appSettings)
        {
            _emiService = service;
            _reportService = reportService;
            _appSettings = appSettings;
            ReportHeaderVM = new();
            EMISetupVM = new(service, reportService);
            EMISetupVM.TemplatePathChanged += (newPath) => TemplatePath = newPath;
        }

        /* ###############################  功能函数  ################################ */

        private EMIUUTdataInfo ReadPath(string dataDir)
        {
            if (string.IsNullOrEmpty(dataDir))
            {
                return null;
            }
            emiDocxFiles.Clear();
            emiPdfFiles.Clear();
            foreach (string datafile in Directory.GetFiles(dataDir))
            {
                string[] infos = System.IO.Path.GetFileNameWithoutExtension(datafile).Split('-');
                if (System.IO.Path.GetExtension(datafile).ToLower().Contains("docx"))
                {
                    if (!emiUUTdatasInfo.SN.Contains(infos[0])) emiUUTdatasInfo.SN.Add(infos[0]);
                    else if (!emiUUTdatasInfo.Voltage.Contains(infos[1])) emiUUTdatasInfo.Voltage.Add(infos[1]);
                    else if (!emiUUTdatasInfo.Load.Contains(infos[2])) emiUUTdatasInfo.Load.Add(infos[2]);
                    else if (!emiUUTdatasInfo.LISN.Contains(infos[3])) emiUUTdatasInfo.LISN.Add(infos[3]);
                    emiDocxFiles.Add(datafile);
                }
                else if (System.IO.Path.GetExtension(datafile).ToLower().Contains("pdf"))
                {
                    emiPdfFiles.Add(datafile);
                }
            }
            return emiUUTdatasInfo;
        }

        private List<EMIUUTData> ReadDatas(List<string> emiDocxPaths)
        {
            static string Last(string[] strs)
            {
                return strs == null || strs.Length == 0 ? "" : strs[strs.Length - 1];
            }

            List<EMIUUTData> emiUUTDatas = [];
            foreach (string emiDocxPath in emiDocxPaths)
            {
                EMIUUTData emiData = new();
                string dataCsv = SmartTableExtractor.ConvertWordTablesToCsv(emiDocxPath);
                string[] dataCsvLines = dataCsv.Split(['\r', '\n']);
                for (int i = 0; i < dataCsvLines.Length; i++)
                {
                    string[] line = dataCsvLines[i].Split(',');
                    if (line.Length is > 1 and <= 4)
                    {
                        if (dataCsvLines[i].Contains("Model")) emiData.Model = Last(line);
                        else if (dataCsvLines[i].Contains("Serial")) emiData.SN = Last(line);
                        else if (dataCsvLines[i].Contains("Power")) emiData.Voltage = Last(line);
                        else if (dataCsvLines[i].Contains("Load")) emiData.Load = Last(line);
                    }
                    else if (line.Length >= 10)
                    {
                        if (string.IsNullOrEmpty(emiData.LISN))
                        {
                            if (line.Contains("L1")) emiData.LISN = "L";
                            else if (line.Contains("N")) emiData.LISN = "N";
                        }
                        List<float> tmp = [];
                        int offest = 0;
                        for (int j = 0; j < 9; j++)
                        {
                            if (offest > 3)
                            {
                                _logger.Error($"{emiData.Name}中找不到结果表格");
                                break;
                            }
                            try
                            {
                                tmp.Add(float.Parse(line[j + offest]));
                            }
                            catch
                            {
                                j--;
                                offest++;
                            }
                        }
                        if (offest <= 3)
                        {
                            tmp[0] = emiData.Datas.Count + 1;
                            emiData.Datas.Add(tmp);
                            if (emiData.MinDatas.Count == 0)
                            {
                                emiData.MinDatas = tmp;
                            }
                            else
                            {
                                if (tmp[PK_Margin_Col] < emiData.MinDatas[PK_Margin_Col] || tmp[AVG_Margin_Col] < emiData.MinDatas[AVG_Margin_Col])
                                {
                                    emiData.MinDatas = tmp;
                                }
                            }
                        }
                    }
                }
                emiUUTDatas.Add(emiData);
            }
            return emiUUTDatas;
        }

        private string GetEMITemplatePath(EMIUUTdataInfo emiUUTdatasInfo)
        {
            string[] excelExtensions = [".xlsx", ".xls", ".xlsm"];
            string[] excelFiles = Directory.GetFiles(_reportService.TemplateDir, "*.*", SearchOption.AllDirectories).Where(file => excelExtensions.Contains(Path.GetExtension(file))).ToArray();
            string tmp = emiUUTdatasInfo.GetUUTFileName(2);
            foreach (string excelfile in excelFiles)
            {
                if (excelfile.Contains(tmp))
                {
                    return excelfile;
                }
            }
            return excelFiles[0];
        }

        private OleEmbedRequest WriteDatas(ISheet ws, Dictionary<string, object> setups, List<EMIUUTData> datas, UUTInfoFromExcel uutInfos)
        {
            int rowStart = 44;
            int colSN = 4;
            int colWorkOrder = 6;
            int colVersion = 8;
            int colDC = 9;
            int colVoltage = 11;
            int colLoad = 14;
            int colLisn = 13;
            int colNo = 15;
            int colFreq = 16;
            int colQP_Limit = 17;
            int colAVG_Limit = 18;
            int colQP_Max = 20;
            int colAVG = 21;
            int colComments = 26;
            string addressTESTED_BY = "F4";
            string addressAPPROVED_BY = "M4";
            string addressPROJECT_NAME = "F5";
            string addressTEST_STAGE = "M5";
            string addressTEST_PERIOD = "F6";
            string addressTEST_CONCLUSION = "M6";

            try
            {
                if (setups["Data"] is Dictionary<string, object> setup_datas && setup_datas["Row"] is Dictionary<string, object> rowSetup && setup_datas["Col"] is Dictionary<string, object> colSetup)
                {
                    rowStart = ToInt(rowSetup["Start"], rowStart);
                    colSN = ToInt(colSetup["SN"], colSN);
                    colWorkOrder = ToInt(colSetup["WorkOrder"], colWorkOrder);
                    colVersion = ToInt(colSetup["Version"], colVersion);
                    colDC = ToInt(colSetup["DC"], colDC);
                    colVoltage = ToInt(colSetup["Voltage"], colVoltage);
                    colLoad = ToInt(colSetup["Load"], colLoad);
                    colLisn = ToInt(colSetup["Phase"], colLisn);
                    colNo = ToInt(colSetup["Mark No"], colNo);
                    colFreq = ToInt(colSetup["Mark Freq"], colFreq);
                    colQP_Limit = ToInt(colSetup["QP Limit"], colQP_Limit);
                    colAVG_Limit = ToInt(colSetup["AVG Limit"], colAVG_Limit);
                    colQP_Max = ToInt(colSetup["QP Max"], colQP_Max);
                    colAVG = ToInt(colSetup["AVG"], colAVG);
                    colComments = ToInt(colSetup["Comments"], colComments);
                }

                if (setups["Header"] is Dictionary<string, object> setup_header)
                {
                    addressTESTED_BY = To_String(setup_header["TESTED_BY"], addressTESTED_BY);
                    addressAPPROVED_BY = To_String(setup_header["APPROVED_BY"], addressAPPROVED_BY);
                    addressPROJECT_NAME = To_String(setup_header["PROJECT_NAME"], addressPROJECT_NAME);
                    addressTEST_STAGE = To_String(setup_header["TEST_STAGE"], addressTEST_STAGE);
                    addressTEST_PERIOD = To_String(setup_header["TEST_PERIOD"], addressTEST_PERIOD);
                    addressTEST_CONCLUSION = To_String(setup_header["TEST_CONCLUSION"], addressTEST_CONCLUSION);
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"{ex.Message}, 从Setup工作表解析行列信息失败，使用默认行列设置");
            }

            // 表头单元格可能读不到（报告/模板里缺对应标题），这里全部空安全，避免生成报告时崩在空引用上
            ExcelNpoi.SetCell(ws, ExcelNpoi.RowOf(addressTESTED_BY), ExcelNpoi.ColumnOf(addressTESTED_BY), ReportHeaderVM?.TESTED_BY?.Data);
            ExcelNpoi.SetCell(ws, ExcelNpoi.RowOf(addressAPPROVED_BY), ExcelNpoi.ColumnOf(addressAPPROVED_BY), ReportHeaderVM?.APPROVED_BY?.Data);
            ExcelNpoi.SetCell(ws, ExcelNpoi.RowOf(addressPROJECT_NAME), ExcelNpoi.ColumnOf(addressPROJECT_NAME), ReportHeaderVM?.PROJECT_NAME?.Data);
            ExcelNpoi.SetCell(ws, ExcelNpoi.RowOf(addressTEST_STAGE), ExcelNpoi.ColumnOf(addressTEST_STAGE), ReportHeaderVM?.TEST_STAGE?.Data);
            ExcelNpoi.SetCell(ws, ExcelNpoi.RowOf(addressTEST_PERIOD), ExcelNpoi.ColumnOf(addressTEST_PERIOD), ReportHeaderVM?.TestStart);
            ExcelNpoi.SetCell(ws, ExcelNpoi.RowOf(addressTEST_CONCLUSION), ExcelNpoi.ColumnOf(addressTEST_CONCLUSION), ReportHeaderVM?.TestPass is true ? "Pass" : "Fail");
            ExcelNpoi.SetCell(ws, rowStart, colWorkOrder, uutInfos?.WorkOrder);
            ExcelNpoi.SetCell(ws, rowStart, colVersion, uutInfos?.Revision);
            ExcelNpoi.SetCell(ws, rowStart, colDC, uutInfos?.DC);

            int sn_rows = 0;
            List<DataCell> SN_cells = [];
            ExcelNpoi.SetCell(ws, rowStart, colSN, emiUUTdatasInfo.SN[0]);
            SN_cells.Add(new DataCell(rowStart, colSN) { Data = emiUUTdatasInfo.SN[0] });

            int sn_written_count = 1;
            for (int _sn_row = rowStart; _sn_row < rowStart + datas.Count; _sn_row++)
            {
                if (sn_written_count >= emiUUTdatasInfo.SN.Count)
                {
                    break;
                }
                if (ExcelNpoi.MergeRegionId(ws, _sn_row, colSN) != ExcelNpoi.MergeRegionId(ws, _sn_row + 1, colSN))
                {
                    SN_cells.Add(new DataCell(_sn_row, colSN) { Data = emiUUTdatasInfo.SN[sn_written_count] });
                    ExcelNpoi.SetCell(ws, _sn_row + 1, colSN, emiUUTdatasInfo.SN[sn_written_count++]);
                    if (sn_rows == 0) sn_rows = _sn_row - rowStart + 1;
                }
            }

            var _datas = datas.GroupBy(d => d.SN).ToDictionary(
                sn => sn.Key,
                sn => sn.GroupBy(d => d.Voltage).ToDictionary(
                    v => v.Key,
                    v => v.GroupBy(d => d.Load)
                        .OrderBy(l => int.Parse(l.Key.TrimEnd('%')))
                        .ToDictionary(
                            l => l.Key,
                            l => l.ToList()
            )));

            int row_cursor = rowStart;
            int uutNo = 1;
            foreach (var sn in _datas)
            {
                int row_snStart = row_cursor;
                ExcelNpoi.SetCell(ws, row_snStart, colSN, sn.Key);
                ExcelNpoi.SetCell(ws, row_snStart, colSN - 2, uutNo);
                ExcelNpoi.SetCell(ws, row_snStart, colDC + 1, uutNo++);
                if (row_snStart != rowStart)
                {
                    ExcelNpoi.SetFormula(ws, row_snStart, colWorkOrder, $"={GetCellColumn(colWorkOrder)}{rowStart}");
                    ExcelNpoi.SetFormula(ws, row_snStart, colVersion, $"={GetCellColumn(colVersion)}{rowStart}");
                    ExcelNpoi.SetFormula(ws, row_snStart, colDC, $"={GetCellColumn(colDC)}{rowStart}");
                }
                foreach (var vol in sn.Value)
                {
                    int row_voltageStart = row_cursor;
                    ExcelNpoi.SetCell(ws, row_voltageStart, colVoltage, vol.Key);
                    ExcelNpoi.SetCell(ws, row_voltageStart, colVoltage + 1, vol.Key.Contains("110") ? "60Hz" : "50Hz");
                    foreach (var load in vol.Value)
                    {
                        int row_loadStart = row_cursor;
                        ExcelNpoi.SetCell(ws, row_loadStart, colLoad, load.Key);
                        foreach (var lisn in load.Value)
                        {
                            ExcelNpoi.SetCell(ws, row_cursor, colLisn, lisn.LISN == "L" ? "Line" : "Neutral");
                            ExcelNpoi.SetCell(ws, row_cursor, colNo, lisn.MinDatas[MarkNo_Col]);
                            ExcelNpoi.SetCell(ws, row_cursor, colFreq, lisn.MinDatas[Freq_Col]);
                            ExcelNpoi.SetCell(ws, row_cursor, colQP_Limit, lisn.MinDatas[PK_Limit_Col]);
                            ExcelNpoi.SetCell(ws, row_cursor, colAVG_Limit, lisn.MinDatas[AVG_Limit_Col]);
                            ExcelNpoi.SetCell(ws, row_cursor, colQP_Max, lisn.MinDatas[PK_Col]);
                            ExcelNpoi.SetCell(ws, row_cursor, colAVG, lisn.MinDatas[AVG_Col]);

                            ExcelNpoi.SetFormula(ws, row_cursor, colAVG + 2, $"T{row_cursor}-Q{row_cursor}");
                            ExcelNpoi.SetFormula(ws, row_cursor, colAVG + 3, $"U{row_cursor}-R{row_cursor}");

                            ExcelNpoi.SetRowHeight(ws, row_cursor, 21.75);
                            row_cursor++;
                        }
                        ExcelNpoi.Merge(ws, row_loadStart, colLoad, row_cursor - 1, colLoad); // 合并负载列
                    }
                    ExcelNpoi.Merge(ws, row_voltageStart, colVoltage, row_cursor - 1, colVoltage); // 合并电压列
                    ExcelNpoi.Merge(ws, row_voltageStart, colVoltage + 1, row_cursor - 1, colVoltage + 1); // 合并频率列
                    ExcelNpoi.Merge(ws, row_voltageStart, colComments - 1, row_cursor - 1, colComments - 1); // 合并Appendix列
                    ExcelNpoi.Merge(ws, row_voltageStart, colComments, row_cursor - 1, colComments); // 合并Comments列
                }
                ExcelNpoi.Merge(ws, row_snStart, colSN - 2, row_cursor - 1, colSN - 1); // 合并No列
                ExcelNpoi.Merge(ws, row_snStart, colSN, row_cursor - 1, colSN + 1); // 合并SN列
                ExcelNpoi.Merge(ws, row_snStart, colWorkOrder, row_cursor - 1, colWorkOrder + 1); // 合并WorlerNo列
                ExcelNpoi.Merge(ws, row_snStart, colVersion, row_cursor - 1, colVersion); // 合并Rev列
                ExcelNpoi.Merge(ws, row_snStart, colDC, row_cursor - 1, colDC); // 合并DC列
                ExcelNpoi.Merge(ws, row_snStart, colDC + 1, row_cursor - 1, colDC + 1); // 合并No.列
            }

            int rowEnd = row_cursor - 1;

            ExcelNpoi.FillRange(ws, rowStart, colAVG_Limit + 1, rowEnd, colAVG_Limit + 1, "-"); //设置Peak Max列的值为"-"
            ExcelNpoi.FillRange(ws, rowStart, colAVG + 1, rowEnd, colAVG + 1, "-"); //设置Margin Peak列的值为"-"

            const string FMT_3_DECIMALS = "0.000";
            const string FMT_2_DECIMALS = "0.00";

            // 写注脚（先写值，样式在下面统一按「每个单元格一次成型」的方式套用：
            // NPOI 一个单元格只有一个样式，不能像 EPPlus 那样分多次叠加，所以这里合并成一份 spec）
            var remarkLines = Remark.Split('\n');
            foreach (var line in remarkLines)
            {
                // 备注行落在模板网格之上：先清掉该行原有的条件标签/占位值，避免备注右侧残留 Line/百分比/0
                ExcelNpoi.ClearCells(ws, row_cursor, 1, colComments + 1);
                ExcelNpoi.SetCell(ws, row_cursor, colSN - 2, line);
                row_cursor++;
            }
            int remarkEnd = row_cursor - 1;

            // 删除模板里本次没用到的多余行：备注块之下（模板预写的条件网格 + 模板自带备注页脚）整段删掉，
            // 生成结果只保留"数据行 + 本次备注"。模板结构：标题/表头(41-43) + 条件网格(44 起，每块 16 行)
            int gridLastRow = 0;
            int scanLast = Math.Max(ExcelNpoi.LastRow(ws), remarkEnd);
            int lastCol = ExcelNpoi.LastColumn(ws);
            for (int r = remarkEnd + 1; r <= scanLast; r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    if (!string.IsNullOrWhiteSpace(ExcelNpoi.CellText(ws, r, c)))
                    {
                        gridLastRow = r;
                        break;
                    }
                }
            }
            if (gridLastRow > remarkEnd)
            {
                ExcelNpoi.DeleteRows(ws, remarkEnd + 1, gridLastRow - remarkEnd);
                _logger.Info($"已删除 EMI 模板中备注块之下的 {gridLastRow - remarkEnd} 行（本次未用到的条件网格与模板自带备注）");
            }

            int[] grayBgColumns = [colLisn, colQP_Limit, colAVG_Limit, colAVG + 1, colAVG + 2, colAVG + 3];
            for (int r = rowStart - 2; r <= Math.Max(rowEnd + 1, remarkEnd); r++)
            {
                for (int c = 1; c <= colComments + 1; c++)
                {
                    ExcelNpoi.CellStyleSpec spec = new();
                    bool inBlock = r >= rowStart && r <= rowEnd;
                    bool inStyled = r >= rowStart - 2 && r <= rowEnd && c >= colSN - 2 && c <= colComments;
                    bool inRemark = r >= rowEnd + 1 && r <= remarkEnd && c == colSN - 2;
                    bool inRemarkBlock = r >= rowEnd + 1 && r <= remarkEnd && c >= 1 && c <= colComments + 1;

                    if (inBlock)
                    {
                        if (c == colFreq)
                        {
                            spec.NumberFormat = FMT_3_DECIMALS;
                        }
                        else if (c == colQP_Limit || c == colAVG_Limit || c == colQP_Max || c == colAVG || c == colAVG + 2 || c == colAVG + 3)
                        {
                            spec.NumberFormat = FMT_2_DECIMALS;
                        }
                    }
                    if (inStyled)
                    {
                        spec.Border = BorderStyle.Thin;
                        spec.Horizontal = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        spec.Vertical = NPOI.SS.UserModel.VerticalAlignment.Center;
                        spec.WrapText = true;
                        spec.FontSize = 10;
                        if (r == rowStart - 2 || r == rowEnd || c == colSN - 2 || c == colComments)
                        {
                            spec.OuterBorder = BorderStyle.Medium;
                        }
                    }
                    if (inRemark)
                    {
                        spec.Horizontal = NPOI.SS.UserModel.HorizontalAlignment.Left;
                        spec.FontSize = 11;
                    }
                    if (inBlock || inRemarkBlock)
                    {
                        spec.FillRgb = [255, 255, 255]; // 白底
                    }
                    if (inBlock && grayBgColumns.Contains(c))
                    {
                        spec.FillRgb = [242, 242, 242]; // 灰底
                    }
                    if (spec.NumberFormat != null || spec.Border.HasValue || spec.OuterBorder.HasValue || spec.Horizontal.HasValue
                        || spec.Vertical.HasValue || spec.WrapText || spec.FontSize.HasValue || spec.FillRgb != null)
                    {
                        ExcelNpoi.ApplyStyle(ws, r, c, r, c, spec);
                    }
                }
            }

            // 插入数据的压缩包（OLE 由调用方在保存后用 Excel COM 嵌入）
            string zipPath = System.IO.Path.Combine(DataPath, $"{datas[0].Model}.zip");
            int zipped = FileService.CreateFilteredZip(DataPath, zipPath, @"\.pdf$");
            if (zipped == 0)
            {
                _logger.Warn("EMI 数据压缩包里没有 PDF（Word 转 PDF 可能未完成），报告将缺少数据附件");
            }
            string iconDir = System.IO.Path.Combine(_reportService.TemplateDir, "ZipEMF");
            if (!Directory.Exists(iconDir))
                Directory.CreateDirectory(iconDir);
            string iconPath = System.IO.Path.Combine(iconDir, $"{datas[0].Model}.zip.emf");
            if (!File.Exists(iconPath))
                Image.GenerateCenteredEmf(iconPath, Resources._7z_Icon, $"{datas[0].Model}.zip");
            return new OleEmbedRequest
            {
                ObjectPath = zipPath,
                TopLeftAddress = new DataCell(rowStart, colComments - 2).TopLeftAddress,
                IconPath = iconPath,
                WidthPx = 120,
                HeightPx = 60,
                OffsetXPx = 0,
                OffsetYPx = 0
            };
        }

        private async void ConvertToPdfAsync(string sourcePath)
        {
            PopupWindow popup = PopupWindow.ShowBusy(LanguageService.Get("Msg_PleaseWait"));
            (int Converted, int Failed) result = (0, 0);
            await Task.Run(() =>
            {
                if (sourcePath.ToLower().EndsWith("docx"))
                {
                    bool ok = Docx2Pdf.ConvertToPdf(sourcePath, Path.ChangeExtension(sourcePath, ".pdf"));
                    result = ok ? (1, 0) : (0, 1);
                }
                else
                {
                    result = Docx2Pdf.ConvertToPdf(sourcePath);
                }
            });
            popup.Close();
            _logger.Info($"Word 转 PDF 完成：成功 {result.Converted} 个，失败 {result.Failed} 个");
            // 转换失败不再静默：后面压缩进报告的数据包会缺内容，必须让用户知道
            if (result.Failed > 0)
            {
                _ = MessageBox.Show(
                    $"有 {result.Failed} 个 Word 文档未能转换为 PDF（成功 {result.Converted} 个）。\n"
                    + "报告里嵌入的数据包会缺少这部分内容，请确认 Word 可用、文件未被占用后重试。",
                    LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void DoReport()
        {
            List<EMIUUTData> emiDatas = await Task.Run(() => ReadDatas(emiDocxFiles));
            if (string.IsNullOrEmpty(TemplatePath))
                TemplatePath = GetEMITemplatePath(emiUUTdatasInfo);

            if (!File.Exists(TemplatePath))
            {
                _logger.Error("EMI报告模板不存在");
                return;
            }
            XSSFWorkbook wb = ExcelNpoi.OpenRead(TemplatePath);
            OleEmbedRequest oleRequest;
            string savePath;
            try
            {
                ISheet ws = ExcelNpoi.SheetByName(wb, "Conducted EMI");
                ISheet ws_setup = ExcelNpoi.SheetByName(wb, "Setup");
                // 模板可能选错：缺工作表时明确报错，避免后面写单元格时空引用崩溃
                if (ws == null || ws_setup == null)
                {
                    _logger.Error($"EMI 模板缺少工作表（需要 \"Conducted EMI\" 与 \"Setup\"）：{TemplatePath}");
                    _ = MessageBox.Show(
                        $"EMI 模板格式不符：缺少 \"Conducted EMI\" 或 \"Setup\" 工作表。\n模板：{TemplatePath}",
                        LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                var setups = SettingsViewModel.ParseJson(ExcelNpoi.CellText(ws_setup, 1, 1));
                int setupIndex = wb.GetSheetIndex(ws_setup);
                if (setupIndex >= 0)
                {
                    wb.RemoveSheetAt(setupIndex);
                }

                oleRequest = await Task.Run(() => WriteDatas(ws, setups, emiDatas, _reportService.UUTInfos));

                savePath = _emiService.SavePathDialog("选择保存路径", "2.1 Conducted EMI Measurement", "EMI报告|*.xlsx", _reportService.RootPath) ?? Directory.GetCurrentDirectory() + "2.1 Conducted EMI Measurement.xlsx";
                ExcelNpoi.Save(wb, savePath);
            }
            finally
            {
                wb.Close();
            }

            // OLE 附件（数据压缩包）在文件保存后用 Excel COM 嵌入
            if (oleRequest != null && File.Exists(oleRequest.ObjectPath))
            {
                oleRequest.SheetName = "Conducted EMI";
                ExcelOleEmbedder.Embed(savePath, [oleRequest]);
            }
            MessageBox.Show($"报告已保存到{savePath}", LanguageService.Get("Cap_SaveSuccess"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /* ###############################  EMIReportPage Command  ################################ */

        private RelayCommand fileSelectCommand;
        public ICommand FileSelectCommand => fileSelectCommand ??= new RelayCommand(FileSelect);

        private void FileSelect()
        {
            TemplatePath = _emiService.OpenPathDialog(LanguageService.Get("Dlg_SelectEMITemplate"));
        }

        private RelayCommand dirSelectCommand;
        public ICommand DirSelectCommand => dirSelectCommand ??= new RelayCommand(DirSelect);

        private void DirSelect()
        {
            DataPath = _emiService.OpenPathDialog(LanguageService.Get("Dlg_SelectEMIData"), filter: "EMI数据文件|*.pdf;*.docx|所有文件|*.*", initPath: _appSettings.EmiDataDir, isDir: true);
            ReadPath(DataPath);
        }

        private RelayCommand toPDFCommand;
        public ICommand ToPDFCommand => toPDFCommand ??= new RelayCommand(ToPDF, CanToPDF);
        private void ToPDF()
        {
            ConvertToPdfAsync(DataPath);
        }

        private bool CanToPDF()
        {
            return !string.IsNullOrEmpty(DataPath) && (Directory.Exists(DataPath) || File.Exists(DataPath));
        }

        private RelayCommand alertTimeCommand;
        public ICommand AlertTimeCommand => alertTimeCommand ??= new RelayCommand(AlertTime, CanToPDF);

        /// <summary>
        /// Alert：打开"EMI 数据处理"对话框，由用户主动填写文件时间参数，并可批量替换 Word 字符串
        /// （原来在这里直接按随机规则改文件时间，现已改为界面操作）
        /// </summary>
        private void AlertTime()
        {
            if (string.IsNullOrWhiteSpace(DataPath) || !Directory.Exists(DataPath))
            {
                _ = MessageBox.Show("请先选择 EMI 测试数据文件夹。", LanguageService.Get("Cap_Warning"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            WindowEmiTools dialog = new(_emiService, DataPath);
            dialog.ShowDialog();
        }

        private RelayCommand emiFinishCommand;
        public ICommand EMIFinishCommand => emiFinishCommand ??= new RelayCommand(EMIFinish, CanEMIFinish);

        private void EMIFinish()
        {
            DoReport();
        }

        private bool CanEMIFinish()
        {
            return true;
        }

    }
}