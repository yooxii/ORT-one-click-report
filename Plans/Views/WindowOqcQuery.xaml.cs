using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// OQC 测试记录查询（原生实现，不依赖 C_SII51.EXE）。
    /// SQL 从原程式里还原：主表是 SIM92（托盘主档）+ SIM96（检验结果）+ SIM94（条码明细）的联接，
    /// 按料号/日期区间/托盘号/工令/条码/结果/状态过滤；下方明细按托盘号查 SIM94 的条码级数据。
    /// 需要先登录 MES 数据库（与回线工令共用同一套连接与账号）。
    /// </summary>
    public partial class WindowOqcQuery : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Sim44Service _service;
        private bool _busy;

        public WindowOqcQuery()
        {
            InitializeComponent();
            _service = App.ServiceProvider.GetRequiredService<Sim44Service>();
            Loaded += async (_, _) => await InitAsync();
        }

        private async Task InitAsync()
        {
            dp_from.SelectedDate = DateTime.Today.AddDays(-30);
            dp_to.SelectedDate = DateTime.Today;
            if (!_service.IsLoggedIn)
            {
                bool ok = await Task.Run(() => _service.TryAutoLogin(out _));
                if (!ok)
                {
                    SetHint(LanguageService.Get("Oqc_Msg_NeedLogin"));
                }
            }
            if (_service.IsLoggedIn)
            {
                SetHint(string.IsNullOrWhiteSpace(_service.LastLoginWarning)
                    ? string.Format(LanguageService.Get("Oqc_Msg_ReadyFormat"), _service.CurrentUser)
                    : _service.LastLoginWarning);
            }
            UpdateLoginUi();
        }

        private void UpdateLoginUi()
        {
            bool loggedIn = _service.IsLoggedIn;
            txt_loginState.Text = loggedIn
                ? string.Format(LanguageService.Get("Sim44_Status_LoggedInFormat"), _service.CurrentUser)
                : LanguageService.Get("Sim44_Status_NotLoggedIn");
            txt_source.Text = LanguageService.Get("Oqc_SourceHint");
            btn_login.IsEnabled = !loggedIn && !_busy;
            btn_login.Visibility = loggedIn ? Visibility.Collapsed : Visibility.Visible;
            btn_query.IsEnabled = !_busy;
            btn_detail.IsEnabled = !_busy;
            btn_clear.IsEnabled = !_busy;
        }

        /// <summary>确保已登录：没登录先试本机凭据，再不行弹 MES 登录框</summary>
        private async Task<bool> EnsureLoginAsync()
        {
            if (_service.IsLoggedIn)
            {
                return true;
            }
            if (await Task.Run(() => _service.TryAutoLogin(out _)))
            {
                UpdateLoginUi();
                return true;
            }
            WindowSim44Login dialog = new(_service) { Owner = this };
            bool ok = dialog.ShowDialog() == true;
            UpdateLoginUi();
            return ok;
        }

        private async void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            WindowSim44Login dialog = new(_service) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                UpdateLoginUi();
                await QueryAsync();
            }
        }

        private void Btn_Settings_Click(object sender, RoutedEventArgs e)
        {
            WindowSim44Settings settings = new() { Owner = this };
            settings.ShowDialog();
        }

        private void Btn_Clear_Click(object sender, RoutedEventArgs e)
        {
            txt_part.Clear();
            txt_pallet.Clear();
            txt_mo.Clear();
            txt_barcode.Clear();
            txt_result.Clear();
            txt_status.Clear();
            dp_from.SelectedDate = DateTime.Today.AddDays(-30);
            dp_to.SelectedDate = DateTime.Today;
            grid_main.ItemsSource = null;
            grid_detail.ItemsSource = null;
            SetHint("");
        }

        private async void Btn_Query_Click(object sender, RoutedEventArgs e) => await QueryAsync();

        private async Task QueryAsync()
        {
            if (!await EnsureLoginAsync())
            {
                SetHint(LanguageService.Get("Oqc_Msg_NeedLogin"));
                return;
            }
            string part = txt_part.Text.Trim();
            string pallet = txt_pallet.Text.Trim();
            string mo = txt_mo.Text.Trim();
            string barcode = txt_barcode.Text.Trim();
            string result = txt_result.Text.Trim();
            string status = txt_status.Text.Trim();
            string sDate = dp_from.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string eDate = dp_to.SelectedDate?.ToString("yyyyMMdd") ?? "";

            await RunBusyAsync(LanguageService.Get("Oqc_Msg_Querying"), async () =>
            {
                DataTable table = await Task.Run(() => _service.QueryOqc(part, sDate, eDate, pallet, mo, barcode, result, status, 300));
                grid_main.ItemsSource = table.DefaultView;
                grid_detail.ItemsSource = null;
                SetHint(table.Rows.Count == 0
                    ? LanguageService.Get("Oqc_Msg_NoResult")
                    : string.Format(LanguageService.Get("Oqc_Msg_ResultFormat"), table.Rows.Count));
            });
        }

        private async void Btn_Detail_Click(object sender, RoutedEventArgs e) => await LoadDetailAsync();

        private async void Grid_Main_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
            => await LoadDetailAsync();

        /// <summary>按选中行的托盘号查 SIM94 条码明细</summary>
        private async Task LoadDetailAsync()
        {
            if (grid_main.SelectedItem is not DataRowView rowView)
            {
                SetHint(LanguageService.Get("Oqc_Msg_SelectRow"));
                return;
            }
            DataRow row = rowView.Row;
            if (!await EnsureLoginAsync())
            {
                return;
            }
            string pallet = Value(row, "PALLET_NO");
            string part = Value(row, "PART_NO");
            await RunBusyAsync(LanguageService.Get("Oqc_Msg_QueryingDetail"), async () =>
            {
                DataTable table = await Task.Run(() => _service.DetailOqc(part, pallet, "", "", 500));
                grid_detail.ItemsSource = table.DefaultView;
                SetHint(string.Format(LanguageService.Get("Oqc_Msg_DetailFormat"), table.Rows.Count, pallet));
            });
        }

        private static string Value(DataRow row, string column)
        {
            return row.Table.Columns.Contains(column) ? Convert.ToString(row[column]) ?? "" : "";
        }

        /// <summary>还想用原程式时，从这里打开 C_SII51（路径可在本机设置里覆盖）</summary>
        private void Btn_Legacy_Click(object sender, RoutedEventArgs e)
        {
            string path = ResolveLegacyPath();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_NotFoundFormat"), path),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = path,
                    WorkingDirectory = Path.GetDirectoryName(path) ?? "",
                    UseShellExecute = true,
                });
                _logger.Info("OQC：已打开原程式 " + path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, string.Format(LanguageService.Get("SpPrograms_Msg_LaunchFailedFormat"), ex.Message),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string ResolveLegacyPath()
        {
            string fromSettings = null;
            try
            {
                Dictionary<string, string> overrides = LocalSettingsStore.Read()?.SpPrograms;
                if (overrides != null && overrides.TryGetValue(SpProgramCatalog.OqcId, out string custom)
                    && !string.IsNullOrWhiteSpace(custom))
                {
                    fromSettings = custom.Trim();
                }
            }
            catch
            {
                // 读不到就用默认路径
            }
            return fromSettings ?? @"C:\actelap\C_SII51.EXE";
        }

        private void SetHint(string text) => txt_hint.Text = text ?? "";

        private async Task RunBusyAsync(string busyText, Func<Task> action)
        {
            if (_busy)
            {
                return;
            }
            _busy = true;
            UpdateLoginUi();
            SetHint(busyText);
            try
            {
                await action();
            }
            catch (Sim44Exception ex)
            {
                _logger.Warn($"OQC 查询失败（{ex.Code}）：{ex.Message}"
                             + (string.IsNullOrWhiteSpace(ex.Detail) ? "" : "｜SQL：" + ex.Detail));
                SetHint(ex.Message);
                MessageBox.Show(this, ex.Message, LanguageService.Get("Cap_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "OQC 查询异常");
                SetHint(ex.Message);
                MessageBox.Show(this, ex.Message, LanguageService.Get("Cap_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _busy = false;
                UpdateLoginUi();
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
