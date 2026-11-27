using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// MES 生产类查询（原生实现，不依赖原程式）：
    /// 1) 月排程 / 小工令——IVM540，还原自 C_IVI540_BL，是 ORT 预领「月排程」的数据来源；
    /// 2) 准备领用单——IVM530_MP，还原自 C_IVR621_BL，是预领「准备领用」导入的数据来源。
    /// 两页都只读，共用同一套 MES 登录与连接设置。
    /// </summary>
    public partial class WindowMesProductionQuery : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Sim44Service _service;
        private bool _busy;

        /// <summary>打开时默认选中的页（0=月排程，1=准备领用）</summary>
        private readonly int _initialTab;

        public WindowMesProductionQuery(int initialTab = 0)
        {
            InitializeComponent();
            _initialTab = initialTab;
            _service = App.ServiceProvider.GetRequiredService<Sim44Service>();
            Loaded += async (_, _) => await InitAsync();
        }

        private async Task InitAsync()
        {
            DateTime today = DateTime.Today;
            dp_s_from.SelectedDate = new DateTime(today.Year, today.Month, 1);
            dp_s_to.SelectedDate = new DateTime(today.Year, today.Month, 1).AddMonths(1).AddDays(-1);
            dp_p_from.SelectedDate = new DateTime(today.Year, today.Month, 1);
            dp_p_to.SelectedDate = today;
            dp_st_from.SelectedDate = new DateTime(today.Year, today.Month, 1).AddMonths(-6);
            dp_st_to.SelectedDate = today;
            tabs.SelectedIndex = _initialTab is 0 or 1 or 2 ? _initialTab : 0;

            if (!_service.IsLoggedIn)
            {
                await Task.Run(() => _service.TryAutoLogin(out _));
            }
            UpdateLoginUi();
            if (_service.IsLoggedIn)
            {
                SetHint(string.Format(LanguageService.Get("Oqc_Msg_ReadyFormat"), _service.CurrentUser));
            }
        }

        private void UpdateLoginUi()
        {
            bool loggedIn = _service.IsLoggedIn;
            txt_loginState.Text = loggedIn
                ? string.Format(LanguageService.Get("Sim44_Status_LoggedInFormat"), _service.CurrentUser)
                : LanguageService.Get("Sim44_Status_NotLoggedIn");
            txt_loginHint.Text = loggedIn
                ? (_service.LastLoginWarning ?? "")
                : LanguageService.Get("ProdQuery_Msg_NeedLogin");
            btn_login.Visibility = loggedIn ? Visibility.Collapsed : Visibility.Visible;
            btn_login.IsEnabled = !_busy;
            btn_s_query.IsEnabled = !_busy;
            btn_p_query.IsEnabled = !_busy;
            btn_st_query.IsEnabled = !_busy;
        }

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
            }
        }

        private void Btn_Settings_Click(object sender, RoutedEventArgs e)
        {
            WindowSim44Settings settings = new() { Owner = this };
            settings.ShowDialog();
        }

        /* ###############################  月排程  ################################ */

        private async void Btn_ScheduleQuery_Click(object sender, RoutedEventArgs e)
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            string from = dp_s_from.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string to = dp_s_to.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string pcb = txt_s_pcb.Text.Trim();
            await RunBusyAsync(LanguageService.Get("Oqc_Msg_Querying"), async () =>
            {
                DataTable table = await Task.Run(() => _service.QuerySchedule(from, to, pcb, 1000));
                grid_schedule.ItemsSource = table.DefaultView;
                txt_s_state.Text = string.Format(LanguageService.Get("ProdQuery_Msg_CountFormat"), table.Rows.Count);
                SetHint(table.Rows.Count == 0
                    ? LanguageService.Get("ProdQuery_Msg_ScheduleEmpty")
                    : string.Format(LanguageService.Get("ProdQuery_Msg_ScheduleFormat"), table.Rows.Count));
            });
        }

        /* ###############################  准备领用  ################################ */

        private async void Btn_PreIssueQuery_Click(object sender, RoutedEventArgs e)
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            string from = dp_p_from.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string to = dp_p_to.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string user = txt_p_user.Text.Trim();
            string v1 = txt_v_from.Text.Trim();
            string v2 = txt_v_to.Text.Trim();
            string p1 = txt_part_from.Text.Trim();
            string p2 = txt_part_to.Text.Trim();
            await RunBusyAsync(LanguageService.Get("Oqc_Msg_Querying"), async () =>
            {
                DataTable table = await Task.Run(() => _service.QueryPrepareIssue(from, to, user, v1, v2, p1, p2, 1000));
                grid_preissue.ItemsSource = table.DefaultView;
                txt_p_state.Text = string.Format(LanguageService.Get("ProdQuery_Msg_CountFormat"), table.Rows.Count);
                SetHint(table.Rows.Count == 0
                    ? LanguageService.Get("ProdQuery_Msg_PreIssueEmpty")
                    : string.Format(LanguageService.Get("ProdQuery_Msg_PreIssueFormat"), table.Rows.Count));
            });
        }

        /* ###############################  成品库存  ################################ */

        private async void Btn_StockQuery_Click(object sender, RoutedEventArgs e)
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            string part = txt_st_part.Text.Trim();
            if (string.IsNullOrEmpty(part))
            {
                MessageBox.Show(this, LanguageService.Get("ProdQuery_Msg_PartRequired"), LanguageService.Get("Cap_Info"));
                txt_st_part.Focus();
                return;
            }
            string from = dp_st_from.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string to = dp_st_to.SelectedDate?.ToString("yyyyMMdd") ?? "";
            string pallet = txt_st_pallet.Text.Trim();
            await RunBusyAsync(LanguageService.Get("Oqc_Msg_Querying"), async () =>
            {
                DataTable table = await Task.Run(() => _service.QueryStock(part, from, to, pallet, 500));
                grid_stock.ItemsSource = table.DefaultView;
                txt_st_state.Text = string.Format(LanguageService.Get("ProdQuery_Msg_CountFormat"), table.Rows.Count);
                SetHint(table.Rows.Count == 0
                    ? LanguageService.Get("ProdQuery_Msg_StockEmpty")
                    : string.Format(LanguageService.Get("ProdQuery_Msg_StockFormat"), table.Rows.Count, part));
            });
        }

        /* ###############################  公共  ################################ */

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
                _logger.Warn($"MES 生产查询失败（{ex.Code}）：{ex.Message}"
                             + (string.IsNullOrWhiteSpace(ex.Detail) ? "" : "｜SQL：" + ex.Detail));
                SetHint(ex.Message);
                MessageBox.Show(this, ex.Message, LanguageService.Get("Cap_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "MES 生产查询异常");
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
