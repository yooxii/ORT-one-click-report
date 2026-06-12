using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;

namespace ORT一键报告.ViewModels
{
    public partial class EMISetupViewModel(IService service) : SettingsViewModel
    {
        private readonly IService _emiService = service;

        private readonly string defaultJson = @"{""Header"":{""TESTED BY"":""F4"",""APPROVED BY"":""M4"",""PROJECT NAME"":""F5"",""TEST STAGE"":""M5"",""TEST STAGE"":""M5"",""TEST PERIOD"":""F6"",""TEST CONCLUSION"":""M6"",},""Data"":{""Row"":{""Start"":44,""End"":103},""Col"":{""SN"":4,""WorkOrder"":6,""Version"":8,""DC"":9,""Voltage"":11,""Freq"":12,""Phase"":13,""Load"":14,""Mark No"":15,""Mark Freq"":16,""QP Limit"":17,""AVG Limit"":18,""QP Max"":20,""AVG"":21,""Comments"":26}}}";


        public event Action<string> TemplatePathChanged;
        private string _templatePath = string.Empty;

        public string TemplatePath
        {
            get => _templatePath;
            set
            {
                if (SetProperty(ref _templatePath, value))
                    TemplatePathChanged?.Invoke(value); // 通知外部
            }
        }

        private string ReadJsonFromExcel(string filePath)
        {
            FileInfo fileInfo = new(filePath);
            using ExcelPackage package = new(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Setup"]; // 获取名为"setup"的工作表

            if (worksheet != null)
            {
                string cellValue = worksheet.Cells[1, 1].Value?.ToString(); // 第一行第一列
                return cellValue;
            }

            return null;
        }

        private string SaveJsonToExcel(string updatedJson)
        {
            string filePath = _emiService.OpenPathDialog("保存设置到", initPath: Path.GetDirectoryName(TemplatePath));
            FileInfo fileInfo = new(filePath);
            using ExcelPackage package = new(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Setup"]; // 获取名为"setup"的工作表

            worksheet?.Cells[1, 1].Value = updatedJson;
            package.Save();
            return filePath;
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
                    LoadFromJson(jsonFromExcel);
                    return;
                }
                else
                {
                    MessageBox.Show("Excel文件中未找到有效的JSON数据！使用默认设置。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载Excel文件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            LoadFromJson(defaultJson);
        }

        private bool ValidateSettings()
        {
            try
            {
                // 收集所有列值
                List<int> columnValues = [];
                Dictionary<int, List<string>> columnNameMap = []; // 值 -> 对应的键列表

                foreach (var datas in SettingsData)
                {
                    if (datas.Key.Contains("列")) // 只检查列相关的设置
                    {
                        if (datas.Value is string val && int.TryParse(val, out int value))
                        {
                            columnValues.Add(value);

                            if (!columnNameMap.ContainsKey(value))
                            {
                                columnNameMap[value] = [];
                            }
                            columnNameMap[value].Add(datas.Key);
                        }
                    }
                }

                // 检查是否有重复值
                bool hasDuplicates = columnValues.Count != columnValues.Distinct().Count();

                if (hasDuplicates)
                {
                    List<string> duplicateColumns = [];
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

        private void SaveSettings()
        {
            if (!ValidateSettings())
            {
                return;
            }

            try
            {
                // 序列化回JSON
                JsonSerializerOptions options = new()
                {
                    WriteIndented = true
                };
                string updatedJson = JsonSerializer.Serialize(SettingItemViewModel.GetDictionary(RootItems), options);
                string savepath = SaveJsonToExcel(updatedJson);
                MessageBox.Show($"设置已保存!\nJSON:\n{savepath}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private RelayCommand loadFromExcelCommand;
        public ICommand LoadFromExcelCommand => loadFromExcelCommand ??= new RelayCommand(LoadFromExcel);

        private RelayCommand saveSettingsCommand;
        public ICommand SaveSettingsCommand => saveSettingsCommand ??= new RelayCommand(SaveSettings);
    }
}
