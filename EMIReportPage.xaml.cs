using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using NLog;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System;
using OfficeOpenXml;
using System.Linq;

namespace ORT一键报告
{
    //2.1 Conducted EMI Measurement
    public class EMIUUTSetup
    {
        public List<string> SN { get; set; } = new List<string>();
        public List<string> Voltage { get; set; } = new List<string>();
        public List<string> Load { get; set; } = new List<string>();
        public List<string> LISN { get; set; } = new List<string>();
        public string Name => $"EMI-{string.Join("_", Voltage.ToArray())}-{string.Join("_", Load.ToArray())}";
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
            Datas = new List<List<float>>();
            MinDatas = new List<float>();
        }
    }

    /// <summary>
    /// EMIReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class EMIReportPage : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly EMIUUTSetup emiUUTsetup = new EMIUUTSetup();
        private readonly List<string> emiDocxFiles = new List<string>();
        private readonly List<string> emiPdfFiles = new List<string>();

        public int PK_TolerableLimit_Col { set; get; } = 4;
        public int AVG_TolerableLimit_Col { set; get; } = 7;

        public EMIReportPage()
        {
            InitializeComponent();
        }

        /* ###############################  功能函数  ################################ */
        private EMIUUTSetup ReadPath(string dataDir)
        {
            emiDocxFiles.Clear();
            emiPdfFiles.Clear();
            foreach (string datafile in Directory.GetFiles(dataDir))
            {
                string[] infos = Path.GetFileNameWithoutExtension(datafile).Split('-');
                if (Path.GetExtension(datafile).ToLower().Contains("docx"))
                {
                    if (!emiUUTsetup.SN.Contains(infos[0])) emiUUTsetup.SN.Add(infos[0]);
                    else if (!emiUUTsetup.Voltage.Contains(infos[1])) emiUUTsetup.Voltage.Add(infos[1]);
                    else if (!emiUUTsetup.Load.Contains(infos[2])) emiUUTsetup.Load.Add(infos[2]);
                    else if (!emiUUTsetup.LISN.Contains(infos[3])) emiUUTsetup.LISN.Add(infos[3]);
                    emiDocxFiles.Add(datafile);
                }
                else if (Path.GetExtension(datafile).ToLower().Contains("pdf"))
                {
                    emiPdfFiles.Add(datafile);
                }
            }
            return emiUUTsetup;
        }

        private List<EMIUUTData> ReadDatas(List<string> emiDocxPaths)
        {
            string Last(string[] strs)
            {
                return strs == null || strs.Length == 0 ? "" : strs[strs.Length - 1];
            }

            List<EMIUUTData> emiUUTDatas = new List<EMIUUTData>();
            foreach (string emiDocxPath in emiDocxPaths)
            {
                EMIUUTData emiData = new EMIUUTData();
                string dataCsv = SmartTableExtractor.ConvertWordTablesToCsv(emiDocxPath);
                var dataCsvLines = dataCsv.Split(new[] { '\r', '\n' });
                for (int i = 0; i < dataCsvLines.Length; i++)
                {
                    var line = dataCsvLines[i].Split(',');
                    if (line.Length > 1 && line.Length <= 4)
                    {
                        if (dataCsvLines[i].Contains("Model")) emiData.Model = Last(line);
                        else if (dataCsvLines[i].Contains("Serial")) emiData.SN = Last(line);
                        else if (dataCsvLines[i].Contains("Power")) emiData.Voltage = Last(line);
                        else if (dataCsvLines[i].Contains("Load")) emiData.Load = Last(line);
                    }
                    else if (line.Length >= 10)
                    {
                        List<float> tmp = new List<float>();
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
                            if (i == 0)
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

        private string GetEMITemplatePath(EMIUUTSetup emiUUTSetup)
        {
            string[] excelExtensions = new[] { ".xlsx", ".xls", ".xlsm" };
            string[] excelFiles = Directory.GetFiles(MainWindow.TemplatePath, "*.*", SearchOption.AllDirectories).Where(file => excelExtensions.Contains(Path.GetExtension(file))).ToArray();
        }

        private void WriteDatas(List<EMIUUTData> datas, string templatePath, string savePath = null)
        {
            if (!File.Exists(templatePath))
            {
                _logger.Error("EMI报告模板不存在");
                return;
            }
            var package = new ExcelPackage(new FileInfo(templatePath));
            var wb = package.Workbook;
            var ws = wb.Worksheets[0];
            
        }

        private async Task ConvertToPdfAsync(string sourcePath)
        {
            PopupWindow popup = new PopupWindow() { Title = "处理中", Message = "请耐心等待..." };
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

        private async Task DoReport()
        {
            List<EMIUUTData> emiDatas = await Task.Run(() => ReadDatas(emiDocxFiles));
        }

        /* ###############################  事件函数  ################################ */

        private void btn_Dir_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string textbox)
            {
                if (FindName(textbox) is TextBox textBox)
                {
                    FileDialog fd = new OpenFileDialog()
                    {
                        Filter = "EMI数据文件|*.pdf;*.docx|所有文件|*.*"
                    };
                    _ = fd.ShowDialog();
                    if (fd.FileName is string fileName && fileName != "")
                    {
                        textBox.Text = Path.GetDirectoryName(fd.FileName);
                    }
                    ReadPath(textBox.Text);
                }
            }
        }

        private void btn_File_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string textbox)
            {
                if (FindName(textbox) is TextBox textBox)
                {
                    FileDialog fd = new OpenFileDialog()
                    {
                        Filter = "EMI模版文件|*.xlsx;*.xlsx|所有文件|*.*"
                    };
                    _ = fd.ShowDialog();
                    if (fd.FileName is string fileName && fileName != "")
                    {
                        textBox.Text = fd.FileName;
                    }
                }
            }
        }

        private void btnReadPath_Click(object sender, RoutedEventArgs e)
        {
            string dataDir = DataPath_text.Text;
            if (!Directory.Exists(dataDir))
            {
                return;
            }
            ReadPath(dataDir);
        }

        private void btnToPDF_Click(object sender, RoutedEventArgs e)
        {
            _ = ConvertToPdfAsync(DataPath_text.Text);
        }

        private void btn_Default_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
