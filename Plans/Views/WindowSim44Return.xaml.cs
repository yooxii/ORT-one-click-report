using NLog;
using ORT一键报告.Main.Views;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// MES 回线工令（SIM44）维护窗口：把原来要在 C_SIM19「重工工令輸入」里手工敲的
    /// 工令 / 机种（料号）/ 线别 / 日期 / 数量，直接在 ORT 回线界面里查询、新增、修改。
    /// 走 32 位桥接进程连 Oracle（见 Sim44Service），首次操作时按需提示登录。
    /// </summary>
    public partial class WindowSim44Return : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Sim44Service _service;
        private List<Sim44Row> _rows = [];
        private bool _busy;

        /// <summary>最近一次操作是否成功（回线窗口据此决定要不要提示）</summary>
        public bool AnyChangeSaved { get; private set; }

        /// <summary>
        /// 构造：可带一条「从回线记录预填」的数据（工令=回线RT工令、线别、日期、数量、机种名用于找料号）。
        /// </summary>
        public WindowSim44Return(Sim44Service service, Sim44Row prefill = null, string modelHint = null)
        {
            InitializeComponent();
            _service = service ?? throw new ArgumentNullException(nameof(service));
            Loaded += async (_, _) =>
            {
                await RefreshLoginStateAsync(attemptAutoLogin: true);
                if (prefill != null)
                {
                    ApplyToEditor(prefill, isFromSelection: true);
                    txt_q_mo.Text = prefill.MoNo;
                }
                if (!string.IsNullOrWhiteSpace(modelHint) && string.IsNullOrWhiteSpace(txt_part.Text))
                {
                    // 机种名只作为「查料号」的提示，不自动写进料号栏（避免把机种名当料号写进 MES）
                    SetStatus(string.Format(LanguageService.Get("Sim44_ModelHintFormat"), modelHint));
                }
                if (_service.IsLoggedIn)
                {
                    txt_mo.Focus();
                }
            };
        }

        /* ###############################  登录态  ################################ */

        private async Task RefreshLoginStateAsync(bool attemptAutoLogin)
        {
            if (!_service.IsLoggedIn && attemptAutoLogin)
            {
                SetStatus(LanguageService.Get("Sim44_Status_LoggingIn"));
                bool ok = await Task.Run(() => _service.TryAutoLogin(out _));
                if (ok)
                {
                    SetStatus(LanguageService.Get("Sim44_AutoLoginOk"));
                }
                else
                {
                    SetStatus(LanguageService.Get("Sim44_Status_NotLoggedIn"));
                }
            }
            UpdateLoginUi();
        }

        private void UpdateLoginUi()
        {
            bool loggedIn = _service.IsLoggedIn;
            txt_loginState.Text = loggedIn
                ? string.Format(LanguageService.Get("Sim44_Status_LoggedInFormat"), _service.CurrentUser)
                : LanguageService.Get("Sim44_Status_NotLoggedIn");
            btn_login.IsEnabled = !loggedIn && !_busy;
            btn_logout.IsEnabled = loggedIn && !_busy;
            foreach (UIElement element in new UIElement[] { btn_query, btn_query_clear, btn_insert, btn_update, btn_clear, btn_lookup_part, btn_lookup_line })
            {
                if (element != null)
                {
                    element.IsEnabled = !_busy;
                }
            }
            if (!loggedIn)
            {
                txt_expiry.Text = "";
                return;
            }
            Sim44CredentialStatus status = _service.GetCredentialStatus();
            if (status.DaysLeft is int days)
            {
                txt_expiry.Text = days < 0
                    ? string.Format(LanguageService.Get("Sim44_Status_ExpiredFormat"), -days)
                    : string.Format(LanguageService.Get("Sim44_Status_DaysLeftFormat"), days);
            }
            else
            {
                txt_expiry.Text = LanguageService.Get("Sim44_Status_ExpiryUnknown");
            }
        }

        /// <summary>确保已登录：没登录先试本机凭据，再不行弹登录框。返回 false 表示用户放弃。</summary>
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
            bool ok = await ShowLoginDialogAsync();
            UpdateLoginUi();
            return ok;
        }

        private Task<bool> ShowLoginDialogAsync()
        {
            WindowSim44Login dialog = new(_service) { Owner = this };
            return Task.FromResult(dialog.ShowDialog() == true);
        }

        private async void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            if (await ShowLoginDialogAsync())
            {
                UpdateLoginUi();
                SetStatus(LanguageService.Get("Sim44_AutoLoginOk"));
            }
        }

        private void Btn_Logout_Click(object sender, RoutedEventArgs e)
        {
            _service.Logout();
            UpdateLoginUi();
            SetStatus(LanguageService.Get("Sim44_Status_LoggedOut"));
        }

        private async void Btn_SelfTest_Click(object sender, RoutedEventArgs e)
        {
            btn_selftest.IsEnabled = false;
            try
            {
                (bool ok, string message) = await Task.Run(() => _service.SelfTest());
                string caption = LanguageService.Get(ok ? "Sim44_SelfTest_Ok" : "Sim44_SelfTest_Fail");
                MessageBox.Show(this, message, caption,
                    MessageBoxButton.OK, ok ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            finally
            {
                btn_selftest.IsEnabled = true;
            }
        }

        /* ###############################  查询  ################################ */

        private async void Btn_Query_Click(object sender, RoutedEventArgs e)
        {
            await QueryAsync();
        }

        private async Task QueryAsync()
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            await RunBusyAsync(LanguageService.Get("Sim44_Status_Querying"), QueryCoreAsync);
        }

        /// <summary>真正的查询动作（不负责登录与忙碌标记，供保存后刷新复用）</summary>
        private async Task QueryCoreAsync()
        {
            string mo = txt_q_mo.Text.Trim();
            string part = txt_q_part.Text.Trim();
            string line = txt_q_line.Text.Trim();
            DateTime? from = dp_q_from.SelectedDate;
            DateTime? to = dp_q_to.SelectedDate;

            var rows = await Task.Run(() => _service.Query(mo, part, line, from, to));
            _rows = [];
            if (rows != null)
            {
                foreach (var token in rows)
                {
                    _rows.Add(Sim44Row.FromJson(token));
                }
            }
            grid_result.ItemsSource = null;
            grid_result.ItemsSource = _rows;
            SetStatus(_rows.Count == 0
                ? LanguageService.Get("Sim44_Status_NoResult")
                : string.Format(LanguageService.Get("Sim44_Status_ResultFormat"), _rows.Count));
        }

        private void Btn_QueryClear_Click(object sender, RoutedEventArgs e)
        {
            txt_q_mo.Clear();
            txt_q_part.Clear();
            txt_q_line.Clear();
            dp_q_from.SelectedDate = null;
            dp_q_to.SelectedDate = null;
            txt_q_mo.Focus();
        }

        private void Grid_Result_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (grid_result.SelectedItem is Sim44Row row)
            {
                ApplyToEditor(row, isFromSelection: true);
                SetStatus(LanguageService.Get("Sim44_Status_RowSelected"));
            }
        }

        /* ###############################  编辑  ################################ */

        private void ApplyToEditor(Sim44Row row, bool isFromSelection)
        {
            txt_mo.Text = row.MoNo ?? "";
            txt_part.Text = row.PartNo ?? "";
            txt_line_no.Text = row.LineNo ?? "";
            txt_line_name.Text = row.LineName ?? "";
            txt_qty.Text = row.MoQty ?? "";
            txt_rma.Text = row.RmaNo ?? "";
            txt_rma_seq.Text = row.RmaSeq ?? "";
            if (DateTime.TryParse(row.StartDate, out DateTime date))
            {
                dp_date.SelectedDate = date;
            }
            else if (!isFromSelection)
            {
                dp_date.SelectedDate = DateTime.Today;
            }
        }

        private Sim44Row ReadEditor(out string error)
        {
            error = null;
            string mo = txt_mo.Text.Trim();
            if (string.IsNullOrEmpty(mo))
            {
                error = LanguageService.Get("Sim44_Msg_MoRequired");
                return null;
            }
            string qty = txt_qty.Text.Trim();
            if (string.IsNullOrEmpty(qty))
            {
                error = LanguageService.Get("Sim44_Msg_QtyRequired");
                return null;
            }
            if (!double.TryParse(qty, out _))
            {
                error = LanguageService.Get("Sim44_Msg_QtyInvalid");
                return null;
            }
            if (dp_date.SelectedDate == null)
            {
                error = LanguageService.Get("Sim44_Msg_DateRequired");
                return null;
            }
            return new Sim44Row
            {
                MoNo = mo,
                PartNo = txt_part.Text.Trim(),
                LineNo = txt_line_no.Text.Trim(),
                LineName = txt_line_name.Text.Trim(),
                MoQty = qty,
                RmaNo = txt_rma.Text.Trim(),
                RmaSeq = txt_rma_seq.Text.Trim(),
                StartDate = dp_date.SelectedDate.Value.ToString("yyyy/MM/dd"),
                UpdUser = _service.CurrentUser,
            };
        }

        private async void Btn_Insert_Click(object sender, RoutedEventArgs e) => await SaveAsync(isInsert: true);

        private async void Btn_Update_Click(object sender, RoutedEventArgs e) => await SaveAsync(isInsert: false);

        private async Task SaveAsync(bool isInsert)
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            Sim44Row row = ReadEditor(out string error);
            if (row == null)
            {
                MessageBox.Show(this, error, LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string confirmText = string.Format(
                LanguageService.Get(isInsert ? "Sim44_Confirm_Insert" : "Sim44_Confirm_Update"),
                row.MoNo, row.PartNo, row.LineNo, row.StartDate, row.MoQty);
            if (MessageBox.Show(this, confirmText, LanguageService.Get("Cap_Info"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }

            await RunBusyAsync(LanguageService.Get("Sim44_Status_Saving"), async () =>
            {
                if (isInsert)
                {
                    int affected = await Task.Run(() => _service.Insert(row));
                    SetStatus(string.Format(LanguageService.Get("Sim44_Status_InsertedFormat"), affected, row.MoNo));
                }
                else
                {
                    // 修改前先确认这笔工令真的存在（避免改了 0 行还以为成功）
                    var existing = await Task.Run(() => _service.GetByMoNo(row.MoNo));
                    if (existing == null || existing.Count == 0)
                    {
                        throw new Sim44Exception("notfound",
                            string.Format(LanguageService.Get("Sim44_Msg_NoSuchMo"), row.MoNo));
                    }
                    int affected = await Task.Run(() => _service.Update(row));
                    SetStatus(string.Format(LanguageService.Get("Sim44_Status_UpdatedFormat"), affected, row.MoNo));
                }
                AnyChangeSaved = true;
                txt_q_mo.Text = row.MoNo;
                await QueryCoreAsync();
            });
        }

        private void Btn_Clear_Click(object sender, RoutedEventArgs e)
        {
            ApplyToEditor(new Sim44Row(), isFromSelection: false);
            dp_date.SelectedDate = DateTime.Today;
            grid_result.SelectedItem = null;
            txt_mo.Focus();
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();

        /* ###############################  查找  ################################ */

        private async void Btn_LookupPart_Click(object sender, RoutedEventArgs e)
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            string text = txt_part.Text.Trim();
            await RunBusyAsync(LanguageService.Get("Sim44_Status_LookingUp"), async () =>
            {
                var rows = await Task.Run(() => _service.LookupPart(text, 100));
                List<WindowListPicker.ListItem> items = [];
                if (rows != null)
                {
                    foreach (var token in rows)
                    {
                        string partNo = token.Value<string>("partNo") ?? "";
                        string model = token.Value<string>("model") ?? "";
                        string spec = token.Value<string>("spec") ?? "";
                        items.Add(new WindowListPicker.ListItem
                        {
                            Value = partNo,
                            Display = partNo + "　" + model + (string.IsNullOrWhiteSpace(spec) ? "" : "　" + spec),
                        });
                    }
                }
                if (items.Count == 0)
                {
                    SetStatus(LanguageService.Get("Sim44_Status_PartNotFound"));
                    return;
                }
                string picked = WindowListPicker.Pick(this, LanguageService.Get("Sim44_LookupPart"),
                    LanguageService.Get("Sim44_LookupPartHint"), items);
                if (!string.IsNullOrEmpty(picked))
                {
                    txt_part.Text = picked;
                    SetStatus(string.Format(LanguageService.Get("Sim44_Status_PartPickedFormat"), picked));
                }
            });
        }

        private async void Btn_LookupLine_Click(object sender, RoutedEventArgs e)
        {
            if (!await EnsureLoginAsync())
            {
                return;
            }
            string text = string.IsNullOrWhiteSpace(txt_line_no.Text) ? txt_line_name.Text.Trim() : txt_line_no.Text.Trim();
            await RunBusyAsync(LanguageService.Get("Sim44_Status_LookingUp"), async () =>
            {
                var rows = await Task.Run(() => _service.LookupLine(text, 100));
                List<WindowListPicker.ListItem> items = [];
                if (rows != null)
                {
                    foreach (var token in rows)
                    {
                        string lineNo = token.Value<string>("lineNo") ?? "";
                        string lineName = token.Value<string>("lineName") ?? "";
                        items.Add(new WindowListPicker.ListItem
                        {
                            Value = lineNo,
                            Display = lineNo + "　" + lineName,
                        });
                    }
                }
                if (items.Count == 0)
                {
                    SetStatus(LanguageService.Get("Sim44_Status_LineNotFound"));
                    return;
                }
                string picked = WindowListPicker.Pick(this, LanguageService.Get("Sim44_LookupLine"),
                    LanguageService.Get("Sim44_LookupLineHint"), items);
                if (!string.IsNullOrEmpty(picked))
                {
                    txt_line_no.Text = picked;
                    // 顺手把线名也带出来
                    foreach (WindowListPicker.ListItem item in items)
                    {
                        if (item.Value == picked)
                        {
                            string[] parts = item.Display.Split('　');
                            if (parts.Length > 1)
                            {
                                txt_line_name.Text = parts[1];
                            }
                            break;
                        }
                    }
                    SetStatus(string.Format(LanguageService.Get("Sim44_Status_LinePickedFormat"), picked));
                }
            });
        }

        /* ###############################  公共  ################################ */

        internal void SetStatus(string text)
        {
            txt_status.Text = text ?? "";
        }

        /// <summary>统一处理「忙 / 异常 / 提示」</summary>
        private async Task RunBusyAsync(string busyText, Func<Task> action)
        {
            if (_busy)
            {
                return;
            }
            _busy = true;
            UpdateLoginUi();
            SetStatus(busyText);
            try
            {
                await action();
            }
            catch (Sim44Exception ex)
            {
                _logger.Warn($"SIM44 操作失败（{ex.Code}）：{ex.Message}"
                             + (string.IsNullOrWhiteSpace(ex.Detail) ? "" : "｜SQL：" + ex.Detail));
                SetStatus(ex.Message);
                if (ex.Code == "nologin")
                {
                    UpdateLoginUi();
                }
                MessageBox.Show(this, ex.Message, LanguageService.Get("Cap_Error"), MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "SIM44 操作异常");
                SetStatus(ex.Message);
                MessageBox.Show(this, ex.Message, LanguageService.Get("Cap_Error"), MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _busy = false;
                UpdateLoginUi();
            }
        }
    }
}

