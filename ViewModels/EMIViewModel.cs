using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告.ViewModels
{
    //2.1 Conducted EMI Measurement
    public class EMIUUTSetup
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
            List<string> res = new();
            if (n >= 1)
                res.Add(string.Join("_", Voltage.ToArray()));
            if (n >= 2)
                res.Add(string.Join("_", Load.ToArray()));
            if (n >= 3)
                res.Add(string.Join("_", LISN.ToArray()));
            return string.Join("-", res.ToArray());
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
            Datas = new List<List<float>>();
            MinDatas = new List<float>();
        }
    }

    public partial class EMIViewModel : ObservableObject
    {
        private readonly IEMIService _emiService;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly EMIUUTSetup emiUUTsetup = new();
        private readonly List<string> emiDocxFiles = new();
        private readonly List<string> emiPdfFiles = new();

        private ExpandoObject settingsData;
        private readonly string defaultJson = @"{""结果开始行"":44,""结束行"":103,""SN列"":4,""工令列"":6,""版本列"":7,""周期列"":8,""电压列"":11,""频率列"":12,""Phase列"":13,""负载列"":14,""Mark No列"":15,""Mark Freq列"":16,""QP Limit列"":17,""AVG Limit列"":18,""QP Max列"":20,""AVG列"":21,""备注列"":26}";

        public ObservableCollection<DynamicField> EMISetupFields { get; set; } = new();

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
                SetProperty(ref _dataPath, value);
                toPDFCommand.RaiseCanExecuteChanged();
            }
        }

        public EMIViewModel(IEMIService emiService)
        {
            _emiService = emiService;
        }

        /* ###############################  功能函数  ################################ */

        private EMIUUTSetup ReadPath(string dataDir)
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
            static string Last(string[] strs)
            {
                return strs == null || strs.Length == 0 ? "" : strs[strs.Length - 1];
            }

            List<EMIUUTData> emiUUTDatas = new();
            foreach (string emiDocxPath in emiDocxPaths)
            {
                EMIUUTData emiData = new();
                string dataCsv = SmartTableExtractor.ConvertWordTablesToCsv(emiDocxPath);
                string[] dataCsvLines = dataCsv.Split(new[] { '\r', '\n' });
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
            foreach (string excelfile in excelFiles)
            {
                if (excelFiles.Contains(emiUUTSetup.GetName(2)))
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

        private async Task DoReport()
        {
            List<EMIUUTData> emiDatas = await Task.Run(() => ReadDatas(emiDocxFiles));
            string templatePath = GetEMITemplatePath(emiUUTsetup);
            await Task.Run(() => WriteDatas(templatePath, emiDatas, MainWindow.UUTInfos));
        }

        public void CreateDynamicForm()
        {
            EMISetupFields.Clear();
            // 动态创建控件
            foreach (KeyValuePair<string, object> setting in settingsData)
            {
                EMISetupFields.Add(new TextField() { Label = setting.Key + ":", Value = setting.Value.ToString() });
            }
        }

        private void LoadAndParseJson(string json)
        {
            JsonSerializerOptions options = new()
            {
                PropertyNameCaseInsensitive = true
            };

            // 解析JSON为ExpandoObject以便动态操作
            settingsData = new ExpandoObject();
            IDictionary<string, object> settingsDict = settingsData;
            JsonDocument doc;
            try
            {
                doc = JsonDocument.Parse(json);
            }
            catch
            {
                MessageBox.Show($"JSON解析失败, 使用默认设置", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                doc = JsonDocument.Parse(defaultJson);
            }

            foreach (JsonProperty property in doc.RootElement.EnumerateObject())
            {
                settingsDict[property.Name] = property.Value.ValueKind == JsonValueKind.Number ? property.Value.GetInt32() : property.Value.GetString();
            }
        }

        private string ReadJsonFromExcel(string filePath)
        {
            FileInfo fileInfo = new(filePath);
            using (ExcelPackage package = new(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Setup"]; // 获取名为"setup"的工作表

                if (worksheet != null)
                {
                    string cellValue = worksheet.Cells[1, 1].Value?.ToString(); // 第一行第一列
                    return cellValue;
                }
            }

            return null;
        }

        private void SaveJsonToExcel(string filePath, string updatedJson)
        {
            FileInfo fileInfo = new(filePath);
            using ExcelPackage package = new(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Setup"]; // 获取名为"setup"的工作表

            if (worksheet != null)
            {
                worksheet.Cells[1, 1].Value = updatedJson;
            }
            package.Save();
        }

        public void LoadFromExcel()
        {
            string excelFilePath;

            if (string.IsNullOrEmpty(TemplatePath))
            {
                excelFilePath = _emiService.OpenPathDialog("请选择EMI模板", filter: "Excel Files|*.xlsx;*.xlsm");
                TemplatePath = excelFilePath;
            }
            else
            {
                excelFilePath = TemplatePath;
            }

            try
            {
                string jsonFromExcel = ReadJsonFromExcel(excelFilePath);
                if (!string.IsNullOrEmpty(jsonFromExcel))
                {
                    LoadAndParseJson(jsonFromExcel);
                    CreateDynamicForm(); // 重新创建表单
                }
                else
                {
                    MessageBox.Show("Excel文件中未找到有效的JSON数据！", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载Excel文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveSettings()
        {
            if (!ValidateSettings())
            {
                return;
            }

            try
            {
                // 更新settingsData
                IDictionary<string, object> settingsDict = settingsData;
                foreach (object fields in EMISetupFields)
                {
                    if (fields is TextField text)
                    {
                        settingsDict[text.Label.Trim(':')] = text.Value;
                    }
                }

                // 序列化回JSON
                JsonSerializerOptions options = new()
                {
                    WriteIndented = true
                };
                string updatedJson = JsonSerializer.Serialize(settingsDict, options);
                SaveJsonToExcel(TemplatePath, updatedJson);
                MessageBox.Show($"设置已保存!\nJSON:\n{TemplatePath}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ValidateSettings()
        {
            try
            {
                // 收集所有列值
                List<int> columnValues = new();
                Dictionary<int, List<string>> columnNameMap = new(); // 值 -> 对应的键列表

                foreach (object fields in EMISetupFields)
                {
                    if (fields is TextField text && text.Label.Contains("列")) // 只检查列相关的设置
                    {
                        if (int.TryParse(text.Value, out int value))
                        {
                            columnValues.Add(value);

                            if (!columnNameMap.ContainsKey(value))
                            {
                                columnNameMap[value] = new List<string>();
                            }
                            columnNameMap[value].Add(text.Label);
                        }
                    }
                }

                // 检查是否有重复值
                bool hasDuplicates = columnValues.Count != columnValues.Distinct().Count();

                if (hasDuplicates)
                {
                    List<string> duplicateColumns = new();
                    foreach (KeyValuePair<int, List<string>> kvp in columnNameMap)
                    {
                        if (kvp.Value.Count > 1)
                        {
                            duplicateColumns.Add($"{string.Join(", ", kvp.Value)} (值: {kvp.Key})");
                        }
                    }

                    MessageBox.Show($"验证失败！存在重复的列设置：\n{string.Join("\n", duplicateColumns)}",
                                  "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"验证过程中出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
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

        private RelayCommand eMISetupCommand;
        public ICommand EMISetupCommand => eMISetupCommand ??= new RelayCommand(EMISetup);

        private void EMISetup()
        {
            EMIReportSetup emisetup = new();
            emisetup.Show();
        }

        /* ###############################  EMIReportSetup Command  ################################ */

        private RelayCommand loadFromExcelCommand;
        public ICommand LoadFromExcelCommand => loadFromExcelCommand ??= new RelayCommand(LoadFromExcel);

        private RelayCommand saveSettingsCommand;
        public ICommand SaveSettingsCommand => saveSettingsCommand ??= new RelayCommand(SaveSettings);
    }
}