using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowAppSettings.xaml 的交互逻辑：参考 VSCode 的设置页面。
    /// 左侧树状目录 + 右侧设置详情，支持同步滚动与点击跳转；
    /// 修改后需点击“保存/应用”才生效（取消即时保存）；数据文件夹仅管理员可修改（改后需重启）。
    /// </summary>
    public partial class WindowAppSettings : Window
    {
        private readonly AppSettingsService _settings;
        private readonly IPathService _pathService;
        private readonly IPermissionService _permission;
        private readonly MailService _mail;
        private readonly MailNotifier _mailNotifier;
        private readonly DatabaseService _db;

        /// <summary>邮件模板编辑器：类型代码 → 主题/正文输入框</summary>
        private readonly Dictionary<string, TextBox> _mailSubjectBoxes = [];
        private readonly Dictionary<string, TextBox> _mailBodyBoxes = [];

        /// <summary>抄送管理员开关：类型代码 → 复选框</summary>
        private readonly Dictionary<string, CheckBox> _mailCcAdminBoxes = [];

        /// <summary>
        /// 防止树选择与滚动互相触发的同步标记
        /// </summary>
        private bool _syncing;

        /// <summary>
        /// 初始加载标记
        /// </summary>
        private bool _loading = true;

        /// <summary>
        /// 打开窗口时的数据文件夹（数据库/附件/配图所在目录，用于判断是否修改）
        /// </summary>
        private string _initialDataFolder;

        /// <summary>
        /// 打开窗口时的 ATE/EMI 路径（用于判断是否修改）
        /// </summary>
        private string _initialAtePath;
        private string _initialEmiPath;

        /// <summary>
        /// 当前用户是否为管理员（仅管理员可修改数据库路径）
        /// </summary>
        private readonly bool _isAdmin;

        /// <summary>
        /// 设置节 Tag → 右侧 Border 映射（按顺序）
        /// </summary>
        private readonly List<(string Tag, Border Section)> _sections = [];

        /// <summary>
        /// 设置节 Tag → 树节点映射
        /// </summary>
        private readonly Dictionary<string, TreeViewItem> _treeNodes = [];

        public WindowAppSettings()
        {
            InitializeComponent();
            _settings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            _pathService = App.ServiceProvider.GetRequiredService<IPathService>();
            _permission = App.ServiceProvider.GetRequiredService<IPermissionService>();
            _mail = App.ServiceProvider.GetRequiredService<MailService>();
            _mailNotifier = App.ServiceProvider.GetRequiredService<MailNotifier>();
            _db = App.ServiceProvider.GetRequiredService<DatabaseService>();
            _isAdmin = _permission.Can("admin.manage");

            CollectSections();
            CollectTreeNodes(tv_settings);

            LoadFontOptions();
            LoadLanguageOptions();
            LoadThemeOptions();
            if (_isAdmin)
            {
                // 模板/抄送管理员控件必须先于载入设置创建，否则复选框初值取不到
                BuildMailTemplateEditors();
            }
            else
            {
                // 邮件服务与数据库路径仅管理员可维护：直接隐藏，普通用户既看不到也不解密邮件密码
                HideAdminOnlySections();
            }
            LoadValues();

            _loading = false;
        }

        /// <summary>
        /// 隐藏仅管理员可见的设置节（邮件服务、邮件模板、数据库路径）：
        /// 同时从左侧目录树与滚动同步列表移除，避免同步到不可见的位置
        /// </summary>
        private void HideAdminOnlySections()
        {
            foreach (string tag in new[] { "sec_mail", "sec_mail_template", "sec_dbpath", "sec_planindex", "sec_reportscan" })
            {
                int index = _sections.FindIndex(section => section.Tag == tag);
                if (index >= 0)
                {
                    _sections[index].Section.Visibility = Visibility.Collapsed;
                    _sections.RemoveAt(index);
                }
                if (_treeNodes.TryGetValue(tag, out TreeViewItem node))
                {
                    node.Visibility = Visibility.Collapsed;
                }
            }
        }

        /* ###############################  收集  ################################ */

        private void CollectSections()
        {
            _sections.Add(("sec_ui", sec_ui));
            _sections.Add(("sec_font", sec_font));
            _sections.Add(("sec_paths", sec_paths));
            _sections.Add(("sec_ate", sec_ate));
            _sections.Add(("sec_emi", sec_emi));
            _sections.Add(("sec_dbpath", sec_dbpath));
            _sections.Add(("sec_schedule", sec_schedule));
            _sections.Add(("sec_requisition", sec_requisition));
            _sections.Add(("sec_report", sec_report));
            _sections.Add(("sec_mail", sec_mail));
            _sections.Add(("sec_mail_template", sec_mail_template));
            _sections.Add(("sec_planindex", sec_planindex));
            _sections.Add(("sec_reportscan", sec_reportscan));
        }

        /// <summary>
        /// 收集 设置节 Tag → 树节点 映射。
        /// 注意：不能在构造函数里遍历可视树——TreeView 的节点容器要到布局阶段才生成，
        /// 此处直接按 Items 递归，XAML 中内联的 TreeViewItem 在解析时即已存在。
        /// </summary>
        private void CollectTreeNodes(ItemsControl parent)
        {
            foreach (object entry in parent.Items)
            {
                if (entry is not TreeViewItem item)
                {
                    continue;
                }
                if (item.Tag is string tag)
                {
                    _treeNodes[tag] = item;
                }
                CollectTreeNodes(item);
            }
        }

        /* ###############################  加载  ################################ */

        private void LoadFontOptions()
        {
            List<string> families = Fonts.SystemFontFamilies
                .Select(f => f.Source)
                .Distinct()
                .OrderBy(n => n)
                .ToList();
            cb_fontFamily.ItemsSource = families;

            // 上限见 AppSettingsService.MaxUiFontSize：更大的字号界面无法完整显示，手输超限会被拒绝
            cb_fontSize.ItemsSource = new[] { 10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24 };
        }

        private void LoadLanguageOptions()
        {
            cb_language.ItemsSource = LanguageService.SupportedLanguages;
            cb_language.SelectedValuePath = "Code";
            string current = LanguageService.GetCurrentLanguage();
            foreach (var lang in LanguageService.SupportedLanguages)
            {
                if (current.StartsWith(lang.Code))
                {
                    cb_language.SelectedValue = lang.Code;
                    break;
                }
            }
        }

        private void LoadThemeOptions()
        {
            var list = new List<ThemeOption>(ThemeService.SupportedThemes);
            if (!string.IsNullOrEmpty(ThemeService.CustomThemePath))
            {
                list.Add(new ThemeOption(ThemeService.CustomCode, "自定义主题", null));
            }
            _loading = true;
            cb_theme.ItemsSource = list;
            cb_theme.SelectedValuePath = "Code";
            cb_theme.SelectedValue = ThemeService.CurrentTheme;
            _loading = false;
        }

        /// <summary>
        /// 填充 Toast 位置选项（默认右上角）
        /// </summary>
        private void LoadToastPositions()
        {
            var list = new List<ThemeOption>
            {
                new ThemeOption("TopRight", LanguageService.Get("Toast_Pos_TopRight"), null),
                new ThemeOption("TopLeft", LanguageService.Get("Toast_Pos_TopLeft"), null),
                new ThemeOption("BottomRight", LanguageService.Get("Toast_Pos_BottomRight"), null),
                new ThemeOption("BottomLeft", LanguageService.Get("Toast_Pos_BottomLeft"), null),
            };
            _loading = true;
            cb_toastPos.ItemsSource = list;
            cb_toastPos.SelectedValuePath = "Code";
            cb_toastPos.SelectedValue = _settings.Settings.UI.ToastPosition ?? "TopRight";
            _loading = false;
        }

        /// <summary>
        /// 填充字重选项（默认常规；中文字体多无 Medium 字面，选中等会被合成为粗体）
        /// </summary>
        private void LoadFontWeightOptions()
        {
            var list = new List<FontWeightOption>
            {
                new("Normal", LanguageService.Get("FontWeight_Normal")),
                new("Medium", LanguageService.Get("FontWeight_Medium")),
                new("SemiBold", LanguageService.Get("FontWeight_SemiBold")),
                new("Bold", LanguageService.Get("FontWeight_Bold")),
            };
            _loading = true;
            cb_fontWeight.ItemsSource = list;
            cb_fontWeight.SelectedValuePath = "Code";
            cb_fontWeight.SelectedValue = _settings.Settings.UI.FontWeight ?? "Normal";
            _loading = false;
        }

        /* ###############################  主界面背景图  ################################ */

        /// <summary>选择主界面背景图（点保存后应用）</summary>
        private void Btn_BrowseBackground_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new()
            {
                Title = LanguageService.Get("Settings_BackgroundImage"),
                Filter = "图片|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*",
                InitialDirectory = System.IO.File.Exists(txt_background.Text)
                    ? System.IO.Path.GetDirectoryName(txt_background.Text)
                    : null
            };
            if (dialog.ShowDialog() == true)
            {
                txt_background.Text = dialog.FileName;
            }
        }

        /// <summary>清除背景图（点保存后生效，主界面回到主题背景）</summary>
        private void Btn_ClearBackground_Click(object sender, RoutedEventArgs e)
        {
            txt_background.Text = "";
        }

        /* ###############################  邮件设置  ################################ */

        /// <summary>
        /// 载入邮件设置到界面
        /// </summary>
        private void LoadMailValues()
        {
            MailSettings mail = _settings.Settings.Mail;
            _loading = true;
            cb_mailSecurity.ItemsSource = new List<MailSecurityOption>
            {
                new("None", LanguageService.Get("Mail_Security_None")),
                new("StartTls", LanguageService.Get("Mail_Security_StartTls")),
                new("Ssl", LanguageService.Get("Mail_Security_Ssl")),
            };
            cb_mailSecurity.SelectedValuePath = "Code";

            chk_mailEnabled.IsChecked = mail.Enabled;
            chk_mailNoticeEnabled.IsChecked = mail.NoticeEnabled;
            chk_mailWarningEnabled.IsChecked = mail.WarningEnabled;
            txt_mailHost.Text = mail.Host;
            txt_mailPort.Text = mail.Port.ToString();
            cb_mailSecurity.SelectedValue = mail.Security ?? "None";
            chk_mailIgnoreCert.IsChecked = mail.IgnoreCertErrors;
            chk_mailDefaultCred.IsChecked = mail.UseDefaultCredentials;
            txt_mailUser.Text = mail.Username;
            txt_mailPassword.Password = mail.Password ?? "";
            txt_mailFrom.Text = mail.FromAddress;
            txt_mailFromName.Text = mail.FromName;
            txt_mailCc.Text = mail.CcList;
            txt_mailTimeout.Text = mail.TimeoutSeconds.ToString();
            chk_mailHtml.IsChecked = mail.BodyIsHtml;
            txt_mailDaysBefore.Text = mail.WarningDaysBefore.ToString();
            txt_mailDedupe.Text = mail.DedupeDays.ToString();
            chk_mailIncludeOverdue.IsChecked = mail.WarningIncludeOverdue;
            foreach (MailTypeDefinition type in MailKind.All)
            {
                if (_mailCcAdminBoxes.TryGetValue(type.Code, out CheckBox ccBox))
                {
                    ccBox.IsChecked = mail.ShouldCcAdmins(type);
                }
            }
            txt_mailTestTo.Text = string.IsNullOrWhiteSpace(txt_mailTestTo.Text) ? mail.FromAddress : txt_mailTestTo.Text;

            // 模板：未自定义时显示内置默认模板，管理员可直接修改
            foreach (MailTypeDefinition type in MailKind.All)
            {
                if (_mailSubjectBoxes.TryGetValue(type.Code, out TextBox subjectBox))
                {
                    subjectBox.Text = mail.GetTemplate(type, true) ?? LanguageService.Get(type.DefaultSubjectKey);
                }
                if (_mailBodyBoxes.TryGetValue(type.Code, out TextBox bodyBox))
                {
                    bodyBox.Text = mail.GetTemplate(type, false) ?? LanguageService.Get(type.DefaultBodyKey);
                }
            }
            _loading = false;
            RefreshMailLogs();
        }

        /// <summary>
        /// 界面值写回邮件设置
        /// </summary>
        private void ApplyMailValues(MailSettings mail)
        {
            mail.Enabled = chk_mailEnabled.IsChecked == true;
            mail.NoticeEnabled = chk_mailNoticeEnabled.IsChecked == true;
            mail.WarningEnabled = chk_mailWarningEnabled.IsChecked == true;
            mail.Host = TrimOrNull(txt_mailHost.Text);
            if (int.TryParse(txt_mailPort.Text?.Trim(), out int port) && port > 0 && port <= 65535)
            {
                mail.Port = port;
            }
            mail.Security = cb_mailSecurity.SelectedValue as string ?? "None";
            mail.IgnoreCertErrors = chk_mailIgnoreCert.IsChecked == true;
            mail.UseDefaultCredentials = chk_mailDefaultCred.IsChecked == true;
            mail.Username = TrimOrNull(txt_mailUser.Text);
            mail.Password = string.IsNullOrEmpty(txt_mailPassword.Password) ? null : txt_mailPassword.Password;
            mail.FromAddress = TrimOrNull(txt_mailFrom.Text);
            mail.FromName = TrimOrNull(txt_mailFromName.Text);
            mail.CcList = TrimOrNull(txt_mailCc.Text);
            if (int.TryParse(txt_mailTimeout.Text?.Trim(), out int timeout) && timeout >= 5 && timeout <= 300)
            {
                mail.TimeoutSeconds = timeout;
            }
            mail.BodyIsHtml = chk_mailHtml.IsChecked == true;
            if (int.TryParse(txt_mailDaysBefore.Text?.Trim(), out int days) && days >= 0 && days <= 365)
            {
                mail.WarningDaysBefore = days;
            }
            if (int.TryParse(txt_mailDedupe.Text?.Trim(), out int dedupe) && dedupe >= 0 && dedupe <= 365)
            {
                mail.DedupeDays = dedupe;
            }
            mail.WarningIncludeOverdue = chk_mailIncludeOverdue.IsChecked == true;
            foreach (MailTypeDefinition type in MailKind.All)
            {
                if (_mailCcAdminBoxes.TryGetValue(type.Code, out CheckBox ccBox))
                {
                    mail.SetCcAdmins(type, ccBox.IsChecked == true);
                }
            }
            foreach (MailTypeDefinition type in MailKind.All)
            {
                if (_mailSubjectBoxes.TryGetValue(type.Code, out TextBox subjectBox))
                {
                    mail.SetTemplate(type, true, subjectBox.Text);
                }
                if (_mailBodyBoxes.TryGetValue(type.Code, out TextBox bodyBox))
                {
                    mail.SetTemplate(type, false, bodyBox.Text);
                }
            }
        }

        /// <summary>
        /// 按邮件类型动态生成模板编辑区（新增邮件类型时自动多出一组）
        /// </summary>
        private void BuildMailTemplateEditors()
        {
            panel_mailTemplates.Children.Clear();
            _mailSubjectBoxes.Clear();
            _mailBodyBoxes.Clear();
            panel_mailCcAdmin.Children.Clear();
            _mailCcAdminBoxes.Clear();
            foreach (MailTypeDefinition type in MailKind.All)
            {
                // 抄送管理员开关（按邮件类型，可在设置中管理）
                CheckBox ccBox = new()
                {
                    Content = LanguageService.Get(type.NameKey),
                    Margin = new Thickness(0, 0, 16, 0)
                };
                _mailCcAdminBoxes[type.Code] = ccBox;
                panel_mailCcAdmin.Children.Add(ccBox);
                StackPanel block = new() { Margin = new Thickness(0, 10, 0, 6) };
                block.Children.Add(new TextBlock
                {
                    Text = LanguageService.Get(type.NameKey),
                    FontWeight = FontWeights.SemiBold
                });
                block.Children.Add(new TextBlock
                {
                    Text = LanguageService.Get(type.VariablesKey),
                    Foreground = (Brush)FindResource("TextSecondaryBrush"),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 2, 0, 6)
                });

                Grid subjectRow = new();
                // 标签列自适应字号（与设置界面其余标签列一致），不再写死像素宽
                subjectRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto, MinWidth = 40 });
                subjectRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                TextBlock subjectLabel = new() { Text = LanguageService.Get("Settings_MailSubject"), VerticalAlignment = VerticalAlignment.Center };
                TextBox subjectBox = new() { VerticalContentAlignment = VerticalAlignment.Center };
                Grid.SetColumn(subjectBox, 1);
                subjectRow.Children.Add(subjectLabel);
                subjectRow.Children.Add(subjectBox);
                block.Children.Add(subjectRow);

                TextBlock bodyLabel = new() { Text = LanguageService.Get("Settings_MailBody"), Margin = new Thickness(0, 6, 0, 2) };
                TextBox bodyBox = new()
                {
                    AcceptsReturn = true,
                    TextWrapping = TextWrapping.Wrap,
                    MinHeight = 110,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    FontFamily = (FontFamily)FindResource("FontFamilyData")
                };
                block.Children.Add(bodyLabel);
                block.Children.Add(bodyBox);

                _mailSubjectBoxes[type.Code] = subjectBox;
                _mailBodyBoxes[type.Code] = bodyBox;
                panel_mailTemplates.Children.Add(block);
            }
        }

        /// <summary>
        /// 刷新最近发送记录（排查用）
        /// </summary>
        private void RefreshMailLogs()
        {
            StringBuilder sb = new();
            foreach (MailLog log in _mail.RecentLogs(20))
            {
                sb.Append(log.CreatedAt.ToString("MM-dd HH:mm")).Append("  ")
                  .Append(log.Success ? "OK" : "FAIL").Append("  ")
                  .Append(log.Kind).Append("  ")
                  .Append(log.Recipients).Append("  ")
                  .Append(log.Subject);
                if (!log.Success && !string.IsNullOrWhiteSpace(log.Error))
                {
                    sb.Append("  -> ").Append(log.Error);
                }
                sb.AppendLine();
            }
            txt_mailLogs.Text = sb.ToString();
        }

        /// <summary>
        /// 发送测试邮件（先保存当前设置）
        /// </summary>
        private void Btn_MailTest_Click(object sender, RoutedEventArgs e)
        {
            if (!ApplyAll())
            {
                return;
            }
            string to = TrimOrNull(txt_mailTestTo.Text) ?? _settings.Settings.Mail.FromAddress;
            MailSendResult result = _mail.SendTest(to);
            if (result.Success)
            {
                _ = MessageBox.Show(string.Format(LanguageService.Get("Msg_MailSentFormat"), result.Recipients),
                    LanguageService.Get("Cap_Success"));
            }
            else
            {
                _ = MessageBox.Show(result.Message ?? "", LanguageService.Get(result.Skipped ? "Cap_Info" : "Cap_Error"));
            }
            RefreshMailLogs();
        }

        /// <summary>
        /// 立即检查计划结束日期并发送警告邮件
        /// </summary>
        private async void Btn_MailCheck_Click(object sender, RoutedEventArgs e)
        {
            if (!ApplyAll())
            {
                return;
            }
            btn_mailCheck.IsEnabled = false;
            try
            {
                (int sent, int skipped, int failed) = await System.Threading.Tasks.Task.Run(() => _mailNotifier.CheckPlanDeadlines());
                _ = MessageBox.Show(string.Format(LanguageService.Get("Msg_MailCheckResultFormat"), sent, skipped, failed),
                    LanguageService.Get("Cap_Info"));
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_Error"));
            }
            finally
            {
                btn_mailCheck.IsEnabled = true;
            }
            RefreshMailLogs();
        }

        private void LoadValues()
        {
            Models.AppSettings settings = _settings.Settings;
            cb_fontFamily.SelectedItem = settings.UI.FontFamily;
            cb_fontSize.Text = settings.UI.FontSize.ToString();
            LoadFontWeightOptions();
            LoadToastPositions();
            txt_background.Text = settings.UI.BackgroundImage ?? "";
            if (_isAdmin)
            {
                LoadMailValues();
            }

            _initialDataFolder = _settings.GetEffectiveDataFolder();
            txt_dataFolder.Text = _initialDataFolder;
            // 数据文件夹只有管理员能改：非管理员直接禁用，避免"改了却存不进去"的困惑
            txt_dataFolder.IsEnabled = _isAdmin;
            btn_dataFolderBrowse.IsEnabled = _isAdmin;
            _initialAtePath = _settings.GetAteDataPath();
            txt_ate.Text = _initialAtePath;
            _initialEmiPath = _settings.GetEmiDataPath();
            txt_emi.Text = _initialEmiPath;
            txt_schedule.Text = settings.Paths.SchedulePath;
            txt_requisition.Text = settings.Paths.RequisitionPath;
            txt_report.Text = settings.Paths.ReportPath;

            // 计划索引：空闲自动执行（仅管理员界面可见）
            chk_planIndexAuto.IsChecked = _settings.GetBool(PlanIndexScheduler.SettingAutoKey, false);
            txt_planIndexIdle.Text = _settings.GetInt(PlanIndexScheduler.SettingIdleMinutesKey,
                PlanIndexScheduler.DefaultIdleMinutes).ToString();

            // 报告扫描：空闲自动执行（仅管理员界面可见）
            chk_reportScanAuto.IsChecked = _settings.GetBool(ReportScanScheduler.SettingAutoKey, true);
            txt_reportScanIdle.Text = _settings.GetInt(ReportScanScheduler.SettingIdleSecondsKey,
                ReportScanScheduler.DefaultIdleSeconds).ToString();
        }

        /* ###############################  保存/应用/取消  ################################ */

        /// <summary>
        /// 将界面值写入设置并保存（数据库路径仅管理员生效）；成功返回 true
        /// </summary>
        private bool ApplyAll()
        {
            Models.AppSettings settings = _settings.Settings;

            if (cb_fontFamily.SelectedItem is string family && !string.IsNullOrWhiteSpace(family))
            {
                settings.UI.FontFamily = family;
            }
            if (!double.TryParse(cb_fontSize.Text, out double size))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_InvalidFontSize"), LanguageService.Get("Cap_Info"));
                return false;
            }
            if (size > AppSettingsService.MaxUiFontSize)
            {
                // 字号过大：警告并放弃本次修改（保留原字号，输入框回退显示当前值）
                _ = MessageBox.Show(
                    string.Format(LocalizationHelper.Get("Msg_FontSizeTooLargeFormat"), AppSettingsService.MaxUiFontSize),
                    LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                cb_fontSize.Text = settings.UI.FontSize.ToString();
                return false;
            }
            if (size < AppSettingsService.MinUiFontSize)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_InvalidFontSize"), LanguageService.Get("Cap_Info"));
                return false;
            }
            settings.UI.FontSize = size;
            if (cb_toastPos.SelectedValue is string toastPos && !string.IsNullOrEmpty(toastPos))
            {
                settings.UI.ToastPosition = toastPos;
            }
            if (cb_fontWeight.SelectedValue is string fontWeight && !string.IsNullOrEmpty(fontWeight))
            {
                settings.UI.FontWeight = fontWeight;
            }
            // 背景图：立即写入并保存，让主界面当场生效
            settings.UI.BackgroundImage = string.IsNullOrWhiteSpace(txt_background.Text) ? null : txt_background.Text.Trim();

            settings.Paths.SchedulePath = TrimOrNull(txt_schedule.Text);
            settings.Paths.RequisitionPath = TrimOrNull(txt_requisition.Text);
            settings.Paths.ReportPath = TrimOrNull(txt_report.Text);

            // 数据文件夹：仅管理员可改；先确认新目录可用（不可用就不保存）
            if (_isAdmin)
            {
                string typedFolder = FolderUtil.Normalize(txt_dataFolder.Text);
                if (string.IsNullOrEmpty(typedFolder))
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_DataFolderEmpty"), LanguageService.Get("Cap_Info"));
                    txt_dataFolder.Focus();
                    return false;
                }
                if (!string.Equals(typedFolder, _settings.GetEffectiveDataFolder(), StringComparison.OrdinalIgnoreCase)
                    && !FolderUtil.TryPrepare(typedFolder, out string folderError))
                {
                    _ = MessageBox.Show(string.Format(LanguageService.Get("Msg_DataFolderInvalidFormat"), folderError),
                        LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    txt_dataFolder.Focus();
                    return false;
                }
                txt_dataFolder.Text = typedFolder;
            }

            // 计划索引：空闲自动执行（非管理员界面未载入，不得回写）
            if (_isAdmin)
            {
                if (!int.TryParse(txt_planIndexIdle.Text?.Trim(), out int idleMinutes) || idleMinutes < 1)
                {
                    _ = MessageBox.Show(LanguageService.Get("PlanIndex_Settings_IdleHint"), LanguageService.Get("Cap_Info"));
                    txt_planIndexIdle.Focus();
                    return false;
                }
                _settings.SetBool(PlanIndexScheduler.SettingAutoKey, chk_planIndexAuto.IsChecked == true);
                _settings.SetInt(PlanIndexScheduler.SettingIdleMinutesKey, idleMinutes);

                // 报告扫描：空闲自动执行（非管理员界面未载入，不得回写）
                if (!int.TryParse(txt_reportScanIdle.Text?.Trim(), out int scanIdleSeconds) || scanIdleSeconds < 10)
                {
                    _ = MessageBox.Show(LanguageService.Get("ReportScan_IdleSecondsHint"), LanguageService.Get("Cap_Info"));
                    txt_reportScanIdle.Focus();
                    return false;
                }
                _settings.SetBool(ReportScanScheduler.SettingAutoKey, chk_reportScanAuto.IsChecked == true);
                _settings.SetInt(ReportScanScheduler.SettingIdleSecondsKey, scanIdleSeconds);
            }
            if (_isAdmin)
            {
                // 非管理员界面未载入邮件设置，不得回写（避免把管理员的配置覆盖掉）
                ApplyMailValues(settings.Mail);
            }
            _settings.Save();

            // ATE/EMI 路径：保存在程序目录本地文件（与数据库路径同位置），仅修改时写入
            string atePath = TrimOrNull(txt_ate.Text);
            if (!string.Equals(atePath, _initialAtePath, StringComparison.OrdinalIgnoreCase))
            {
                _settings.SetAteDataPath(atePath);
                _initialAtePath = atePath;
            }
            string emiPath = TrimOrNull(txt_emi.Text);
            if (!string.Equals(emiPath, _initialEmiPath, StringComparison.OrdinalIgnoreCase))
            {
                _settings.SetEmiDataPath(emiPath);
                _initialEmiPath = emiPath;
            }

            // 数据文件夹：仅管理员。判据是"输入框里的路径 vs 磁盘上已保存的路径"，
            // 不能用"是否与打开窗口时的值不同"这类内存标记——否则改完再打开会回到旧值。
            if (TrySaveDataFolderFromUi(out string activeFolder, out string newFolder, out bool offerCopy))
            {
                PromptRestartForDataFolder(activeFolder, newFolder, offerCopy);
            }
            return true;
        }

        /// <summary>
        /// 把界面上的数据文件夹写回本机设置（仅管理员）。
        /// 返回是否需要提示重启；activeFolder 为本次运行配置的数据文件夹（重启前不会变），
        /// newFolder 为已保存的新文件夹，offerCopy 表示要不要问"是否把现有数据复制过去"。
        /// </summary>
        private bool TrySaveDataFolderFromUi(out string activeFolder, out string newFolder, out bool offerCopy)
        {
            // 与"本次运行真正在用的数据文件夹"比较时要用配置值（UNC 会被映射成盘符，实际路径不同）
            activeFolder = FolderUtil.Normalize(_db.ConfiguredDataFolder) ?? _db.DataDir;
            newFolder = null;
            offerCopy = false;
            if (!_isAdmin)
            {
                return false;
            }
            string typed = FolderUtil.Normalize(txt_dataFolder.Text);
            string stored = _settings.GetEffectiveDataFolder();
            if (string.Equals(typed, stored, StringComparison.OrdinalIgnoreCase))
            {
                // 输入框没改：只有"上次改过但还没重启"时才需要再提示一次
                if (string.Equals(stored, activeFolder, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
                newFolder = stored;
                _initialDataFolder = stored;
                return true;
            }
            _settings.SetDataFolder(typed);
            newFolder = _settings.GetEffectiveDataFolder();
            _initialDataFolder = newFolder;
            txt_dataFolder.Text = newFolder;
            offerCopy = true;
            return true;
        }

        /// <summary>
        /// 数据文件夹变更后的收尾：可选把当前数据复制到新文件夹，并提示重启（支持一键重启）
        /// </summary>
        private void PromptRestartForDataFolder(string oldFolder, string newFolder, bool offerCopy)
        {
            // 新文件夹里没有数据库时，问一下要不要把当前数据带过去（换到远程共享时最常用）
            string newDb = System.IO.Path.Combine(newFolder, "ort_plans.db");
            if (offerCopy && !System.IO.File.Exists(newDb) && System.IO.Directory.Exists(oldFolder))
            {
                MessageBoxResult copy = MessageBox.Show(
                    LanguageService.Get("Msg_DataFolderCopyAsk"),
                    LanguageService.Get("Cap_Info"), MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (copy == MessageBoxResult.Yes)
                {
                    System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                    try
                    {
                        (int copied, List<string> errors) = DataFolderMigrator.CopyTo(oldFolder, newFolder, _db.FreeSql);
                        string message = errors.Count == 0
                            ? string.Format(LanguageService.Get("Msg_DataFolderCopiedFormat"), copied)
                            : string.Format(LanguageService.Get("Msg_DataFolderCopyFailedFormat"), string.Join(Environment.NewLine, errors));
                        MessageBox.Show(message, LanguageService.Get(errors.Count == 0 ? "Cap_Success" : "Cap_Error"),
                            MessageBoxButton.OK, errors.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
                    }
                    finally
                    {
                        System.Windows.Input.Mouse.OverrideCursor = null;
                    }
                }
            }
            // 重启提示：一键重启
            MessageBoxResult restart = MessageBox.Show(
                LanguageService.Get("Msg_DataFolderRestartAsk"),
                LanguageService.Get("Cap_Info"), MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (restart == MessageBoxResult.Yes)
            {
                App.RestartApplication();
            }
            else
            {
                // 稍后自己重启：明确告知还没生效
                MessageBox.Show(LanguageService.Get("Msg_DataFolderSaved"), LanguageService.Get("Cap_Info"));
            }
        }

        private void Cb_Language_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                    if (_loading || cb_language.SelectedValue is not string code)
                    {
                        return;
                    }
                    LanguageService.SetLanguage(code);
                }
        
                private void Cb_Theme_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                    if (_loading || cb_theme.SelectedValue is not string code)
                    {
                        return;
                    }
                    if (code == ThemeService.CustomCode)
                    {
                        if (!string.IsNullOrEmpty(ThemeService.CustomThemePath))
                        {
                            ThemeService.ApplyCustomTheme(ThemeService.CustomThemePath, out _);
                        }
                    }
                    else
                    {
                        ThemeService.ApplyTheme(code);
                    }
                }

        private void Btn_ImportTheme_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "XAML 主题文件|*.xaml|所有文件|*.*",
                Title = "选择自定义主题文件",
            };
            if (dlg.ShowDialog() != true)
            {
                return;
            }

            if (ThemeService.ApplyCustomTheme(dlg.FileName, out string error))
            {
                LoadThemeOptions();
                cb_theme.SelectedValue = ThemeService.CustomCode;
                MessageBox.Show("自定义主题已导入并应用，下次启动自动生效。", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("导入失败：" + error, "错误",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        
        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            if (ApplyAll())
            {
                Close();
            }
        }

        private void Btn_Apply_Click(object sender, RoutedEventArgs e)
        {
            _ = ApplyAll();
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private static string TrimOrNull(string text)
            => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        /* ###############################  路径浏览  ################################ */

        private void Btn_Browse_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button button || button.Tag is not string key)
            {
                return;
            }
            TextBox target = key switch
            {
                "SchedulePath" => txt_schedule,
                "RequisitionPath" => txt_requisition,
                "ReportPath" => txt_report,
                "AteDataPath" => txt_ate,
                "EmiDataPath" => txt_emi,
                _ => null
            };
            if (target == null)
            {
                return;
            }
            string dir = _pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectDir"), initPath: target.Text, isDir: true);
            if (dir != null)
            {
                target.Text = dir;
            }
        }

        private void Btn_DataFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            if (!_isAdmin)
            {
                return;
            }
            string dir = _pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectDataFolder"), initPath: txt_dataFolder.Text, isDir: true);
            if (dir != null)
            {
                txt_dataFolder.Text = dir;
            }
        }

        /* ###############################  同步滚动  ################################ */

        /// <summary>
        /// 树节点点击 → 右侧滚动到对应设置节
        /// </summary>
        private void Tv_Settings_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_syncing || e.NewValue is not TreeViewItem item || item.Tag is not string tag)
            {
                return;
            }
            Border section = _sections.FirstOrDefault(s => s.Tag == tag).Section;
            if (section == null)
            {
                return;
            }
            _syncing = true;
            GeneralTransform transform = section.TransformToAncestor(sv_right);
            double offset = transform.Transform(new Point(0, 0)).Y + sv_right.VerticalOffset;
            sv_right.ScrollToVerticalOffset(offset);
            Dispatcher.BeginInvoke(new Action(() => _syncing = false));
        }

        /// <summary>
        /// 右侧滚动 → 左侧树高亮当前可见的设置节
        /// </summary>
        private void Sv_Right_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (_syncing)
            {
                return;
            }
            // 找到顶部最接近视口顶端且已滚过的节
            string currentTag = _sections[0].Tag;
            foreach ((string Tag, Border Section) pair in _sections)
            {
                GeneralTransform transform = pair.Section.TransformToAncestor(sv_right);
                double top = transform.Transform(new Point(0, 0)).Y;
                if (top <= 1)
                {
                    currentTag = pair.Tag;
                }
                else
                {
                    break;
                }
            }
            if (_treeNodes.TryGetValue(currentTag, out TreeViewItem node) && tv_settings.SelectedItem != node)
            {
                _syncing = true;
                node.IsSelected = true;
                Dispatcher.BeginInvoke(new Action(() => _syncing = false));
            }
        }
    }
}
