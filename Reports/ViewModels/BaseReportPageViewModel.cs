using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using NLog;
using OfficeOpenXml;
using ORT一键报告.Models;
using ORT一键报告.Reports.Models;
using ORT一键报告.Reports.Views;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Reports.ViewModels
{
    public partial class BaseReportPageViewModel(IPathService service, ReportService reportService, AppSettingsService appSettings) : ObservableObject
    {
        private readonly IPathService _Service = service;
        private readonly AppSettingsService _appSettings = appSettings;
        private readonly ReportService _reportService = reportService;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ObservableCollection<ResultDetails> DetailsList { get; set; } = [];
        public ReportHeaderViewModel ReportHeaderVM { get; set; } = new();

        public string RootReportPath { get; set; }

        private string _atePath = "请点击右侧按钮选择ATE数据文件";
        public string ATEPath { get => _atePath; set => SetProperty(ref _atePath, value); }

        private int _testTime;
        public int TestTime
        {
            get => _testTime;
            set => SetProperty(ref _testTime, value);
        }

        private string _reportType;
        public string ReportType
        {
            get => _reportType;
            set => SetProperty(ref _reportType, value);
        }

        /* ###############################  功能函数  ################################ */

        public void ReportFinish(ReportHeaderViewModel reportHeaderInfo)
        {
            _logger.Info($"{ReportType}报告生成中...");
            string saveReportPath;
            try
            {
                string currentPath = Directory.GetCurrentDirectory();
                string templatePath = GetTemplatePath(Path.Combine(currentPath, "Templates"), ReportType);
                if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                {
                    _logger.Warn($"未找到{ReportType}报告模板，无法生成报告");
                    MessageBox.Show($"未找到{ReportType}报告模板，无法生成报告", LanguageService.Get("Cap_Error"));
                    return;
                }
                FileInfo reportFI = new(templatePath);
                string templateInReport = GetTemplatePath(_reportService.RootPath, ReportType);
                string initDir = string.IsNullOrWhiteSpace(templateInReport)
                    ? Path.GetDirectoryName(templatePath)
                    : Path.GetDirectoryName(templateInReport);
                SaveFileDialog saveFileDialog = new()
                {
                    FileName = reportFI.Name,
                    Filter = "Excel文件|*.xlsx;*.xls",
                    InitialDirectory = initDir
                };
                ExcelPackage package = new(reportFI);
                ExcelWorkbook wb = package.Workbook;
                ExcelWorksheet ws = wb.Worksheets[0];
                ExcelWorksheet ws_setup = wb.Worksheets[1];


                // 1.表头信息
                _logger.Info("处理表头");
                for (int r = 1; r <= 8; r++)
                {
                    ws.Cells[ws_setup.Cells[r, 1].Text].Value = reportHeaderInfo.HeaderInfoList[r - 1];
                }

                // 2.单体数据
                _logger.Info("处理单体数据");
                List<object> detailInfoList =
                [
                    DetailsList.Select(r => r.BIroom).ToList(),
                    DetailsList.Select(r => r.BIarea).ToList(),
                    DetailsList.Select(r => r.BIplace).ToList(),
                    DetailsList.Select(r => r.SN).ToList(),
                    DetailsList.Select(r => r.WorkOrder).ToList(),
                    DetailsList.Select(r => r.Version).ToList(),
                    DetailsList.Select(r => r.DC).ToList(),
                    DetailsList.Select(r => r.InspectionPrev).ToList(),
                    DetailsList.Select(r => r.InspectionAfter).ToList(),
                    DetailsList.Select(r => r.FunPrev).ToList(),
                    DetailsList.Select(r => r.FunAfter).ToList(),
                    DetailsList.Select(r => r.HiPot).ToList(),
                ];
                if (!ReportType.ToLower().Contains("burn"))
                {
                    detailInfoList.RemoveRange(0, 3);
                }
                /* 根据模板setup表的定义来保存结果。
                 */
                int _detail_start_row = 13; //setup表detail的起始行
                for (int r = _detail_start_row; r < ws_setup.Dimension.End.Row; r++)
                {
                    ExcelAddress address = new(ws_setup.Cells[r, 1].Text);
                    int Rp_row = address.Start.Row;
                    int Rp_col = address.Start.Column;
                    if (detailInfoList[r - _detail_start_row] is List<string> detailInfo)
                    {
                        for (int i = 0; i < detailInfo.Count; i++)
                        {
                            ws.Cells[Rp_row + i, Rp_col].Value = detailInfo[i];
                        }
                    }
                    else if (detailInfoList[r - _detail_start_row] is List<ReportStatus> detailStatus)
                    {
                        for (int i = 0; i < detailStatus.Count; i++)
                        {
                            ws.Cells[Rp_row + i, Rp_col].Value = detailStatus[i].ToString();
                        }
                    }
                }

                // 3.图片和OLE对象
                _logger.Info("处理图片和OLE对象");
                ExcelAddPicture(ws, "Issue_Photos", reportHeaderInfo.Issue_Photos_Pics, ws_setup.Cells["A11"].Text, ReportType, _reportService.TempPath);
                ExcelAddPicture(ws, "Test_Setup", reportHeaderInfo.Test_Setup_Pics, ws_setup.Cells["A12"].Text, ReportType, _reportService.TempPath);

                string ate_Addr = ws_setup.Cells["A9"].Text;
                EmbedOleObjectWithEpplus(ws, ATEPath, ate_Addr);
                wb.Worksheets.Delete(ws_setup); // 删除设置表

                saveReportPath = saveFileDialog.ShowDialog() == true
                    ? saveFileDialog.FileName
                    : Path.Combine(Directory.GetCurrentDirectory(), reportFI.Name);
                saveReportPath = Path.GetFullPath(saveReportPath);
                package.SaveAs(saveReportPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{ReportType}模版生成失败");
                MessageBox.Show(ex + $"{ReportType}模版生成失败");
                return;
            }
            _logger.Info($"{ReportType}报告生成完成, 保存在{saveReportPath}");
            MessageBox.Show($"{ReportType}报告生成完成, 保存在{saveReportPath}", LanguageService.Get("Cap_Success"));
        }

        private RelayCommand finishCommand;
        public ICommand FinishCommand => finishCommand ??= new RelayCommand(Finish);

        private async void Finish()
        {
            PopupWindow popup = new() { Title = LanguageService.Get("Title_SaveReport"), Message = "处理中..." };
            try
            {
                popup.Show();
                await Task.Run(() => { ReportFinish(ReportHeaderVM); });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存报告时出现错误");
                popup.Message = "保存报告时出现错误" + ex.Message;
            }
            finally
            {
                popup.Close();
            }
        }

        private RelayCommand selectATEDatasCommand;
        public ICommand SelectATEDatasCommand => selectATEDatasCommand ??= new RelayCommand(SelectATEDatas);

        private void SelectATEDatas()
        {
            ATEPath = _Service.OpenPathDialog(LanguageService.Get("Dlg_SelectATEData"), initPath: _appSettings.AteDataDir);
        }
    }
}
