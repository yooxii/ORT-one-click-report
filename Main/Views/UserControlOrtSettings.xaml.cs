using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// ORT 看板取数设置控件：普通选项（日期 / MoType）与高级选项（目标报表 + 全部链接配置）。
    /// 两个入口共用这一份界面：
    /// 1) 设置窗口「ORT 看板取数」（随设置一起保存）；
    /// 2) 「领退 → 工具 → 看板平台配置」（技术员及以上，独立窗口，自己带保存按钮）。
    /// 编辑在副本上进行，<see cref="Apply"/> 校验通过才写库。
    /// </summary>
    public partial class UserControlOrtSettings : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>取数设置服务（数据库 app_settings 的 ort.* 键）</summary>
        private OrtBoardSettingsService _store;

        /// <summary>界面编辑用的设置副本（保存成功后才替换）</summary>
        private OrtBoardSettings _settings;

        /// <summary>载入界面期间不触发整月开关的联动</summary>
        private bool _loading;

        public UserControlOrtSettings()
        {
            InitializeComponent();
        }

        /* ###############################  对外接口  ################################ */

        /// <summary>载入当前设置（设置窗口打开时、独立窗口打开时各调一次）</summary>
        public void Load()
        {
            _store ??= App.ServiceProvider.GetRequiredService<OrtBoardSettingsService>();
            _settings = _store.Get();
            LoadFrom(_settings);
        }

        /// <summary>把一份设置显示到界面上（恢复默认时也走这里，不再回头读库）</summary>
        private void LoadFrom(OrtBoardSettings data)
        {
            OrtBoardQuery query = data.Query;
            OrtBoardEndpoints ep = data.Endpoints;

            _loading = true;
            try
            {
                chk_ortWholeMonth.IsChecked = query.WholeMonth;
                dp_ortFrom.SelectedDate = ParseYmd(query.DateFrom);
                dp_ortTo.SelectedDate = ParseYmd(query.DateTo);
                UpdateDatePickers();
                txt_ortMoType.Text = query.MoType ?? "";

                txt_ortEntryId.Text = ep.TargetEntryId ?? "";
                txt_ortWidget.Text = ep.WidgetHint ?? "";
                txt_ortLoginUrl.Text = ep.LoginUrl ?? "";
                txt_ortRedirectUrl.Text = ep.RedirectUrl ?? "";
                txt_ortReportHost.Text = ep.ReportHost ?? "";
                txt_ortPortalOrigin.Text = ep.PortalOrigin ?? "";
                txt_ortEntryPath.Text = ep.EntryAccessPath ?? "";
                txt_ortConfigPath.Text = ep.ConfigPath ?? "";
                txt_ortLoadPath.Text = ep.LoadContentPath ?? "";
                txt_ortDeviceUuid.Text = ep.DeviceUuid ?? "";
                txt_ortUserRegion.Text = ep.UserRegion ?? "";
                txt_ortUserLanguage.Text = ep.UserLanguage ?? "";
                txt_ortWidth.Text = ep.DashboardWidth.ToString();
                txt_ortHeight.Text = ep.DashboardHeight.ToString();
                txt_ortRefererHeight.Text = ep.LoadRefererHeight.ToString();
                txt_ortPageIndex.Text = ep.PageIndex.ToString();
                txt_ortDateFromKey.Text = ep.DateFromKey ?? "";
                txt_ortDateToKey.Text = ep.DateToKey ?? "";
                txt_ortMoTypeKey.Text = ep.MoTypeKey ?? "";
                txt_ortLabelKeys.Text = $"{ep.DateLabelKey} / {ep.MoTypeLabelKey}";
                txt_ortLabelTexts.Text = $"{ep.DateLabelText} / {ep.MoTypeLabelText}";
            }
            finally
            {
                _loading = false;
            }
        }

        /// <summary>保存取数设置（先校验链接/日期，错了就不写）；成功返回 true</summary>
        public bool Apply()
        {
            _store ??= App.ServiceProvider.GetRequiredService<OrtBoardSettingsService>();
            OrtBoardSettings data = Collect();
            if (!data.IsQueryReady(out string queryError))
            {
                _ = MessageBox.Show(queryError, LanguageService.Get("Ort_Settings_Title"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            foreach ((string Label, string Value) item in new[]
            {
                (LanguageService.Get("Ort_Settings_LoginUrl"), data.Endpoints.LoginUrl),
                (LanguageService.Get("Ort_Settings_RedirectUrl"), data.Endpoints.RedirectUrl),
                (LanguageService.Get("Ort_Settings_ReportHost"), data.Endpoints.ReportHost)
            })
            {
                if (string.IsNullOrWhiteSpace(item.Value)
                    || !Uri.TryCreate(item.Value, UriKind.Absolute, out _))
                {
                    _ = MessageBox.Show(string.Format(LanguageService.Get("Ort_Settings_UrlInvalidFormat"), item.Label),
                        LanguageService.Get("Ort_Settings_Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
            _store.Save(data);
            _settings = data;
            return true;
        }

        /// <summary>切到「普通 / 高级」页签（设置窗口左侧目录点击时用）</summary>
        public void SelectTab(bool advanced) => tc_ort.SelectedIndex = advanced ? 1 : 0;

        /* ###############################  界面逻辑  ################################ */

        /// <summary>「整月取数」勾选状态 → 日期框可用性</summary>
        private void UpdateDatePickers()
        {
            bool wholeMonth = chk_ortWholeMonth.IsChecked == true;
            dp_ortFrom.IsEnabled = !wholeMonth;
            dp_ortTo.IsEnabled = !wholeMonth;
            if (wholeMonth)
            {
                DateTime now = DateTime.Now;
                dp_ortFrom.SelectedDate = new DateTime(now.Year, now.Month, 1);
                dp_ortTo.SelectedDate = now;
            }
        }

        /// <summary>整月开关变化：勾上就把区间拉到本月 1 日~今天，取消则允许手改</summary>
        private void Chk_OrtWholeMonth_Changed(object sender, RoutedEventArgs e)
        {
            if (_loading)
            {
                return;
            }
            UpdateDatePickers();
        }

        /// <summary>从一段看板链接里提取 entryId（避免手抄 GUID 抄错）</summary>
        private void Btn_OrtExtractEntry_Click(object sender, RoutedEventArgs e)
        {
            ORT一键报告.Admin.Views.WindowAdminInput dialog = new(
                LanguageService.Get("Ort_Settings_ExtractEntry"),
                (LanguageService.Get("Ort_Settings_ExtractEntryPrompt"), "", false))
            { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true)
            {
                return;
            }
            string input = dialog.Values.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(input))
            {
                return;
            }
            string entryId = OrtBoardEndpoints.ExtractEntryId(input);
            if (string.IsNullOrEmpty(entryId))
            {
                _ = MessageBox.Show(LanguageService.Get("Ort_Settings_ExtractEntryFailed"),
                    LanguageService.Get("Ort_Settings_Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            txt_ortEntryId.Text = entryId;
        }

        /// <summary>
        /// 读取控件列表（WidgetHint 的动态获取）：
        /// 用界面上当前的 TargetEntryId 登录一次，读回这张看板的控件名，让用户直接挑一个。
        /// </summary>
        private async void Btn_OrtProbeWidget_Click(object sender, RoutedEventArgs e)
        {
            btn_ortProbeWidget.IsEnabled = false;
            try
            {
                if (string.IsNullOrWhiteSpace(txt_ortEntryId.Text))
                {
                    _ = MessageBox.Show(LanguageService.Get("Ort_Settings_EntryRequired"),
                        LanguageService.Get("Ort_Settings_Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                _store ??= App.ServiceProvider.GetRequiredService<OrtBoardSettingsService>();
                // 先按界面值存一次配置：探测用的是同一套链接配置，避免"界面改了但探测还用旧值"
                _store.Save(Collect());

                (bool ok, string message, List<string> widgets) = await App.ServiceProvider
                    .GetRequiredService<OrtBoardService>().ProbeWidgetsAsync().ConfigureAwait(true);
                if (!ok)
                {
                    _ = MessageBox.Show(message ?? LanguageService.Get("Ort_Settings_ProbeFailed"),
                        LanguageService.Get("Ort_Settings_Title"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                List<WindowListPicker.ListItem> items = widgets
                    .Select(w => new WindowListPicker.ListItem { Display = w, Value = w }).ToList();
                string picked = WindowListPicker.Pick(Window.GetWindow(this),
                    LanguageService.Get("Ort_Settings_ProbeWidget"),
                    LanguageService.Get("Ort_Settings_ProbePickHint"), items);
                if (!string.IsNullOrEmpty(picked))
                {
                    txt_ortWidget.Text = picked;
                }
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "读取控件列表失败");
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Ort_Settings_Title"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                btn_ortProbeWidget.IsEnabled = true;
            }
        }

        /// <summary>恢复默认：链接配置回到默认，TargetEntryId/WidgetHint 清空（它们不是默认值）</summary>
        private void Btn_OrtRestoreDefaults_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(LanguageService.Get("Ort_Settings_RestoreConfirm"),
                LanguageService.Get("Ort_Settings_Title"), MessageBoxButton.YesNo, MessageBoxImage.Question)
                != MessageBoxResult.Yes)
            {
                return;
            }
            OrtBoardSettings reset = (_settings ?? new OrtBoardSettings()).Clone();
            reset.Endpoints = reset.Endpoints.ResetToDefault();
            _settings = reset;

            // 只刷新界面，不写库：设置窗口随「保存」一起落库，独立窗口由用户点自己的保存
            LoadFrom(reset);
            _ = MessageBox.Show(LanguageService.Get("Ort_Settings_RestoreDone"),
                LanguageService.Get("Ort_Settings_Title"), MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>把界面值收集成设置对象（不写库）</summary>
        private OrtBoardSettings Collect()
        {
            OrtBoardSettings data = _settings?.Clone() ?? new OrtBoardSettings();
            data.Query.WholeMonth = chk_ortWholeMonth.IsChecked == true;
            if (chk_ortWholeMonth.IsChecked == true)
            {
                data.Query.ApplyWholeMonth();
            }
            else
            {
                data.Query.DateFrom = dp_ortFrom.SelectedDate.HasValue
                    ? OrtBoardSettings.ToYmd(dp_ortFrom.SelectedDate.Value) : null;
                data.Query.DateTo = dp_ortTo.SelectedDate.HasValue
                    ? OrtBoardSettings.ToYmd(dp_ortTo.SelectedDate.Value) : null;
            }
            data.Query.MoType = (txt_ortMoType.Text ?? "").Trim();

            OrtBoardEndpoints ep = data.Endpoints;
            ep.TargetEntryId = (txt_ortEntryId.Text ?? "").Trim();
            ep.WidgetHint = (txt_ortWidget.Text ?? "").Trim();
            ep.LoginUrl = (txt_ortLoginUrl.Text ?? "").Trim();
            ep.RedirectUrl = (txt_ortRedirectUrl.Text ?? "").Trim();
            ep.ReportHost = (txt_ortReportHost.Text ?? "").Trim();
            ep.PortalOrigin = (txt_ortPortalOrigin.Text ?? "").Trim();
            ep.EntryAccessPath = (txt_ortEntryPath.Text ?? "").Trim();
            ep.ConfigPath = (txt_ortConfigPath.Text ?? "").Trim();
            ep.LoadContentPath = (txt_ortLoadPath.Text ?? "").Trim();
            ep.DeviceUuid = (txt_ortDeviceUuid.Text ?? "").Trim();
            ep.UserRegion = (txt_ortUserRegion.Text ?? "").Trim();
            ep.UserLanguage = (txt_ortUserLanguage.Text ?? "").Trim();
            ep.DashboardWidth = ToInt(txt_ortWidth.Text, ep.DashboardWidth);
            ep.DashboardHeight = ToInt(txt_ortHeight.Text, ep.DashboardHeight);
            ep.LoadRefererHeight = ToInt(txt_ortRefererHeight.Text, ep.LoadRefererHeight);
            ep.PageIndex = ToInt(txt_ortPageIndex.Text, ep.PageIndex);
            ep.DateFromKey = (txt_ortDateFromKey.Text ?? "").Trim();
            ep.DateToKey = (txt_ortDateToKey.Text ?? "").Trim();
            ep.MoTypeKey = (txt_ortMoTypeKey.Text ?? "").Trim();

            string[] labelKeys = SplitPair(txt_ortLabelKeys.Text);
            if (labelKeys.Length == 2)
            {
                ep.DateLabelKey = labelKeys[0];
                ep.MoTypeLabelKey = labelKeys[1];
            }
            string[] labelTexts = SplitPair(txt_ortLabelTexts.Text);
            if (labelTexts.Length == 2)
            {
                ep.DateLabelText = labelTexts[0];
                ep.MoTypeLabelText = labelTexts[1];
            }
            return data;
        }

        private static DateTime? ParseYmd(string text)
            => OrtBoardSettings.TryParseYmd(text, out DateTime value) ? value : null;

        private static int ToInt(string text, int fallback)
            => int.TryParse((text ?? "").Trim(), out int value) && value > 0 ? value : fallback;

        /// <summary>"A / B" 或 "A,B" → 两项（空项保持空）</summary>
        private static string[] SplitPair(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return [];
            }
            string[] parts = text.Split(['/', ',', ';', '|']);
            if (parts.Length < 2)
            {
                return [];
            }
            return [parts[0].Trim(), parts[1].Trim()];
        }
    }
}
