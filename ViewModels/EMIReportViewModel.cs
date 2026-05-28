using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using OfficeOpenXml;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ORT一键报告.ViewModels
{
    //2.1 Conducted EMI Measurement
    public class EMIUUTdataInfo
    {
        public List<string> SN { get; set; } = new List<string>();
        public List<string> Voltage { get; set; } = new List<string>();
        public List<string> Load { get; set; } = new List<string>();
        public List<string> LISN { get; set; } = new List<string>();

        /// <summary>
        /// 返回该机种的信息，字符串形式
        /// </summary>
        /// <param name="n">控制返回信息种类的个数，最大3</param>
        /// <returns></returns>
        public string GetName(int n)
        {
            string res = "";
            if (n >= 1)
                res += $"-{string.Join("_", Voltage.ToArray())}";
            if (n >= 2)
                res += $"-{string.Join("_", Load.ToArray())}".Replace("%", "");
            if (n >= 3)
                res += $"-{string.Join("_", LISN.ToArray())}";
            return res;
        }
    }

    public class EMIUUTData
    {
        public string Name => $"{SN}-{Voltage}-{Load}-{LISN}";
        public string SN { get; set; }
        public string Voltage { get; set; }
        public string Load { get; set; }
        public string LISN { get; set; }
        public string Model { get; set; }

        public List<List<float>> Datas { get; set; }
        public List<float> MinDatas { get; set; }

        public EMIUUTData()
        {
            Datas = [];
            MinDatas = [];
        }
    }

    public partial class EMIReportViewModel : ObservableObject
    {
        private readonly IService _emiService;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly EMIUUTdataInfo emiUUTdatasInfo = new();
        private readonly List<string> emiDocxFiles = [];
        private readonly List<string> emiPdfFiles = [];

        public ReportHeaderViewModel ReportHeaderVM { get; set; }
        public EMISetupViewModel EMISetupVM { get; set; }

        public int PK_TolerableLimit_Col { set; get; } = 4;
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
                        List<float> tmp = new();
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
            string tmp = emiUUTdatasInfo.GetName(2);
            foreach (string excelfile in excelFiles)
            {
                if (excelfile.Contains(tmp))
                {
                    return excelfile;
                }
            }
            return excelFiles[0];
        }

        private void WriteDatas(string templatePath, List<EMIUUTData> data, UUTInfoFromExcel uutInfos, string savePath = "")
        {
            if (!File.Exists(templatePath))
            {
                _logger.Error("EMI报告模板不存在");
                return;
            }
            ExcelPackage package = new(new FileInfo(templatePath));
            ExcelWorkbook wb = package.Workbook;
            ExcelWorksheet ws = wb.Worksheets["Conducted EMI"];
            ExcelWorksheet ws_setup = wb.Worksheets["Setup"];


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
            await Task.Run(() => WriteDatas(TemplatePath, emiDatas, MainWindow.UUTInfos));
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

        private RelayCommand emiSetupCommand;
        public ICommand EMISetupCommand => emiSetupCommand ??= new RelayCommand(EMISetup);

        private void EMISetup()
        {
            EMIReportSetup emisetup = new();
            emisetup.Show();
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