using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告.ViewModels
{
    //2.1 Conducted EMI Measurement
    public partial class EMIReportViewModel : ObservableObject
    {
        private readonly IService _emiService;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly EMIUUTdataInfo emiUUTdatasInfo = new();
        private readonly List<string> emiDocxFiles = [];
        private readonly List<string> emiPdfFiles = [];

        public ReportHeaderViewModel ReportHeaderVM { get; set; }
        public EMISetupViewModel EMISetupVM { get; set; }

        public int PK_Col { set; get; } = 2;
        public int PK_Limit_Col { set; get; } = 3;
        public int PK_TolerableLimit_Col { set; get; } = 4;
        public int AVG_Col { set; get; } = 5;
        public int AVG_Limit_Col { set; get; } = 6;
        public int AVG_TolerableLimit_Col { set; get; } = 7;

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
                    toPDFCommand.RaiseCanExecuteChanged();
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

        public EMIReportViewModel(IService service)
        {
            _emiService = service;
            ReportHeaderVM = new();
            EMISetupVM = new(service);
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
                string[] infos = Path.GetFileNameWithoutExtension(datafile).Split('-');
                if (Path.GetExtension(datafile).ToLower().Contains("docx"))
                {
                    if (!emiUUTdatasInfo.SN.Contains(infos[0])) emiUUTdatasInfo.SN.Add(infos[0]);
                    else if (!emiUUTdatasInfo.Voltage.Contains(infos[1])) emiUUTdatasInfo.Voltage.Add(infos[1]);
                    else if (!emiUUTdatasInfo.Load.Contains(infos[2])) emiUUTdatasInfo.Load.Add(infos[2]);
                    else if (!emiUUTdatasInfo.LISN.Contains(infos[3])) emiUUTdatasInfo.LISN.Add(infos[3]);
                    emiDocxFiles.Add(datafile);
                }
                else if (Path.GetExtension(datafile).ToLower().Contains("pdf"))
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
                                if (tmp[PK_TolerableLimit_Col] < emiData.MinDatas[PK_TolerableLimit_Col] || tmp[AVG_TolerableLimit_Col] < emiData.MinDatas[AVG_TolerableLimit_Col])
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
            string[] excelFiles = Directory.GetFiles(MainWindow.TemplateDir, "*.*", SearchOption.AllDirectories).Where(file => excelExtensions.Contains(Path.GetExtension(file))).ToArray();
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

        private void WriteDatas(ExcelWorksheet ws, Dictionary<string, object> setups, List<EMIUUTData> datas, UUTInfoFromExcel uutInfos)
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

            ws.Cells[addressTESTED_BY].Value = ReportHeaderVM?.TESTED_BY.Data ?? null;
            ws.Cells[addressAPPROVED_BY].Value = ReportHeaderVM?.APPROVED_BY.Data ?? null;
            ws.Cells[addressPROJECT_NAME].Value = ReportHeaderVM?.PROJECT_NAME.Data ?? null;
            ws.Cells[addressTEST_STAGE].Value = ReportHeaderVM?.TEST_STAGE.Data ?? null;
            ws.Cells[addressTEST_PERIOD].Value = ReportHeaderVM?.TestStart ?? null;
            ws.Cells[addressTEST_CONCLUSION].Value = ReportHeaderVM?.TestPass is true ? "Pass" : "Fail";
            ws.Cells[rowStart, colWorkOrder].Value = uutInfos?.WorkOrder ?? null;
            ws.Cells[rowStart, colVersion].Value = uutInfos?.Revision ?? null;
            ws.Cells[rowStart, colDC].Value = uutInfos?.DC ?? null;

            int sn_rows = 0;
            List<DataCell> SN_cells = [];
            ws.Cells[rowStart, colSN].Value = emiUUTdatasInfo.SN[0];
            SN_cells.Add(new DataCell(rowStart, colSN) { Data = emiUUTdatasInfo.SN[0] });

            int sn_written_count = 1;
            for (int _sn_row = rowStart; _sn_row < rowStart + datas.Count; _sn_row++)
            {
                if (sn_written_count >= emiUUTdatasInfo.SN.Count)
                {
                    break;
                }
                if (ws.GetMergeCellId(_sn_row, colSN) != ws.GetMergeCellId(_sn_row + 1, colSN))
                {
                    SN_cells.Add(new DataCell(_sn_row, colSN) { Data = emiUUTdatasInfo.SN[sn_written_count] });
                    ws.Cells[_sn_row + 1, colSN].Value = emiUUTdatasInfo.SN[sn_written_count++];
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
                ws.Cells[row_snStart, colSN].Value = sn.Key;
                ws.Cells[row_snStart, colSN - 2].Value = uutNo;
                ws.Cells[row_snStart, colDC + 1].Value = uutNo++;
                if (row_snStart != rowStart)
                {
                    ws.Cells[row_snStart, colWorkOrder].Formula = $"={GetCellColumn(colWorkOrder)}{rowStart}";
                    ws.Cells[row_snStart, colVersion].Formula = $"={GetCellColumn(colVersion)}{rowStart}";
                    ws.Cells[row_snStart, colDC].Formula = $"={GetCellColumn(colDC)}{rowStart}";
                }
                foreach (var vol in sn.Value)
                {
                    int row_voltageStart = row_cursor;
                    ws.Cells[row_voltageStart, colVoltage].Value = vol.Key;
                    ws.Cells[row_voltageStart, colVoltage + 1].Value = vol.Key.Contains("110") ? "60Hz" : "50Hz";
                    foreach (var load in vol.Value)
                    {
                        int row_loadStart = row_cursor;
                        ws.Cells[row_loadStart, colLoad].Value = load.Key;
                        foreach (var lisn in load.Value)
                        {
                            ws.Cells[row_cursor, colLisn].Value = lisn.LISN == "L" ? "Line" : "Neutral";
                            ws.Cells[row_cursor, colNo].Value = lisn.MinDatas[0];
                            ws.Cells[row_cursor, colFreq].Value = lisn.MinDatas[1];
                            ws.Cells[row_cursor, colQP_Limit].Value = lisn.MinDatas[3];
                            ws.Cells[row_cursor, colAVG_Limit].Value = lisn.MinDatas[6];
                            ws.Cells[row_cursor, colQP_Max].Value = lisn.MinDatas[2];
                            ws.Cells[row_cursor, colAVG].Value = lisn.MinDatas[5];

                            ws.Cells[row_cursor, colAVG + 2].Formula = $"T{row_cursor}-Q{row_cursor}";
                            ws.Cells[row_cursor, colAVG + 3].Formula = $"U{row_cursor}-R{row_cursor}";

                            ws.Row(row_cursor).Height = 21.75;
                            row_cursor++;
                        }
                        ws.Cells[row_loadStart, colLoad, row_cursor - 1, colLoad].Merge = true; // 合并负载列
                    }
                    ws.Cells[row_voltageStart, colVoltage, row_cursor - 1, colVoltage].Merge = true; // 合并电压列
                    ws.Cells[row_voltageStart, colVoltage + 1, row_cursor - 1, colVoltage + 1].Merge = true; // 合并频率列
                    ws.Cells[row_voltageStart, colComments - 1, row_cursor - 1, colComments - 1].Merge = true; // 合并Appendix列
                    ws.Cells[row_voltageStart, colComments, row_cursor - 1, colComments].Merge = true; // 合并Comments列
                }
                ws.Cells[row_snStart, colSN - 2, row_cursor - 1, colSN - 1].Merge = true; // 合并No列
                ws.Cells[row_snStart, colSN, row_cursor - 1, colSN + 1].Merge = true; // 合并SN列
                ws.Cells[row_snStart, colWorkOrder, row_cursor - 1, colWorkOrder + 1].Merge = true; // 合并WorlerNo列
                ws.Cells[row_snStart, colVersion, row_cursor - 1, colVersion].Merge = true; // 合并Rev列
                ws.Cells[row_snStart, colDC, row_cursor - 1, colDC].Merge = true; // 合并DC列
                ws.Cells[row_snStart, colDC + 1, row_cursor - 1, colDC + 1].Merge = true; // 合并No.列
            }

            int rowEnd = row_cursor - 1;

            ws.Cells[rowStart, colAVG_Limit + 1, rowEnd, colAVG_Limit + 1].Value = "-"; //设置Peak Max列的值为"-"
            ws.Cells[rowStart, colAVG + 1, rowEnd, colAVG + 1].Value = "-"; //设置Margin Peak列的值为"-"

            const string FMT_3_DECIMALS = "0.000";
            const string FMT_2_DECIMALS = "0.00";

            // 设置数字格式
            ws.Cells[rowStart, colFreq, rowEnd, colFreq].Style.Numberformat.Format = FMT_3_DECIMALS;
            ws.Cells[rowStart, colQP_Limit, rowEnd, colQP_Limit].Style.Numberformat.Format = FMT_2_DECIMALS;
            ws.Cells[rowStart, colAVG_Limit, rowEnd, colAVG_Limit].Style.Numberformat.Format = FMT_2_DECIMALS;
            ws.Cells[rowStart, colQP_Max, rowEnd, colQP_Max].Style.Numberformat.Format = FMT_2_DECIMALS;
            ws.Cells[rowStart, colAVG, rowEnd, colAVG].Style.Numberformat.Format = FMT_2_DECIMALS;
            ws.Cells[rowStart, colAVG + 2, rowEnd, colAVG + 2].Style.Numberformat.Format = FMT_2_DECIMALS;
            ws.Cells[rowStart, colAVG + 3, rowEnd, colAVG + 3].Style.Numberformat.Format = FMT_2_DECIMALS;

            // 设置边框和样式
            ws.Cells[rowStart, colSN - 2, rowEnd, colComments].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[rowStart, colSN - 2, rowEnd, colComments].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells[rowStart, colSN - 2, rowEnd, colComments].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[rowStart, colSN - 2, rowEnd, colComments].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[rowStart - 2, colSN - 2, rowEnd, colComments].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            ws.Cells[rowStart - 2, colSN - 2, rowEnd, colComments].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[rowStart - 2, colSN - 2, rowEnd, colComments].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells[rowStart - 2, colSN - 2, rowEnd, colComments].Style.WrapText = true;
            ws.Cells[rowStart - 2, colSN - 2, rowEnd, colComments].Style.Font.Size = 10;
            ws.Cells[rowStart, 1, rowEnd + 1, colComments + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[rowStart, 1, rowEnd + 1, colComments + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            // 设置一些列为灰色背景
            var grayBgColumns = new[] { colLisn, colQP_Limit, colAVG_Limit, colAVG + 1, colAVG + 2, colAVG + 3 };
            foreach (var col in grayBgColumns)
            {
                ws.Cells[rowStart, col, rowEnd, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[rowStart, col, rowEnd, col].Style.Fill.BackgroundColor.SetColor(255, 242, 242, 242);
            }

            // 写入注脚
            var remarkLines = Remark.Split('\n');
            foreach (var line in remarkLines)
            {
                ws.Cells[row_cursor, colSN - 2].Value = line;
                ws.Cells[row_cursor, colSN - 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ws.Cells[row_cursor, colSN - 2].Style.Font.Size = 11;
                row_cursor++;
            }
            // 设置注脚区域的边框和背景
            ws.Cells[rowEnd + 1, 1, row_cursor, colComments + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[rowEnd + 1, 1, row_cursor, colComments + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            // 插入数据的压缩包
            string zipPath = Path.Combine(DataPath, $"{datas[0].Model}.zip");
            CreateFilteredZip(DataPath, zipPath, @"\.pdf$");
            string iconDir = Path.Combine(MainWindow.TemplateDir, "ZipEMF");
            if (!Directory.Exists(iconDir))
                Directory.CreateDirectory(iconDir);
            string iconPath = Path.Combine(iconDir, $"{datas[0].Model}.zip.emf");
            if (!File.Exists(iconPath))
                ImageUtils.GenerateCenteredEmf(iconPath, Resources._7z_Icon, $"{datas[0].Model}.zip");
            EmbedOleObjectWithEpplus(ws, zipPath, new DataCell(rowStart, colComments - 2).TopLeftAddress, iconPath, 0, 0, 120, 60);
        }

        private async Task ConvertToPdfAsync(string sourcePath)
        {
            PopupWindow popup = new() { Title = "处理中", Message = "请耐心等待..." };
            popup.Show();
            await Task.Run(() =>
            {
                if (sourcePath.ToLower().EndsWith("docx"))
                    Docx2Pdf.ConvertToPdf(sourcePath, sourcePath.Split('.')[0] + "pdf");
                else
                    Docx2Pdf.ConvertToPdf(sourcePath);
            });
            popup.Close();
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
            using ExcelPackage package = new(new FileInfo(TemplatePath));
            ExcelWorkbook wb = package.Workbook;
            ExcelWorksheet ws = wb.Worksheets["Conducted EMI"];

            ExcelWorksheet ws_setup = wb.Worksheets["Setup"];
            var setups = SettingsViewModel.ParseJson(ws_setup.Cells["A1"].Text);
            wb.Worksheets.Delete(ws_setup);

            await Task.Run(() => WriteDatas(ws, setups, emiDatas, MainWindow.UUTInfos));

            string savePath = _emiService.SavePathDialog("选择保存路径", "2.1 Conducted EMI Measurement", "EMI报告|*.xlsx", MainWindow.RootPath) ?? Directory.GetCurrentDirectory() + "2.1 Conducted EMI Measurement.xlsx";
            package.SaveAs(savePath);
            MessageBox.Show($"报告已保存到{savePath}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /* ###############################  EMIReportPage Command  ################################ */

        private RelayCommand fileSelectCommand;
        public ICommand FileSelectCommand => fileSelectCommand ??= new RelayCommand(FileSelect);

        private void FileSelect()
        {
            TemplatePath = _emiService.OpenPathDialog("请选择EMI模板");
        }

        private RelayCommand dirSelectCommand;
        public ICommand DirSelectCommand => dirSelectCommand ??= new RelayCommand(DirSelect);

        private void DirSelect()
        {
            DataPath = _emiService.OpenPathDialog("请选择EMI数据", filter: "EMI数据文件|*.pdf;*.docx|所有文件|*.*", isDir: true);
            ReadPath(DataPath);
        }

        private RelayCommand toPDFCommand;
        public ICommand ToPDFCommand => toPDFCommand ??= new RelayCommand(ToPDF, CanToPDF);

        private void ToPDF()
        {
            _ = ConvertToPdfAsync(DataPath);
        }

        private bool CanToPDF()
        {
            return !string.IsNullOrEmpty(DataPath) && (Directory.Exists(DataPath) || File.Exists(DataPath));
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