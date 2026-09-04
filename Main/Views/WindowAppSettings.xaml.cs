using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowAppSettings.xaml 的交互逻辑：参考 VSCode 的设置页面。
    /// 左侧树状目录 + 右侧设置详情，支持同步滚动与点击跳转；
    /// 修改后需点击“保存/应用”才生效（取消即时保存）；数据库路径仅管理员可修改。
    /// </summary>
    public partial class WindowAppSettings : Window
    {
        private readonly AppSettingsService _settings;
        private readonly IPathService _pathService;
        private readonly IPermissionService _permission;

        /// <summary>
        /// 防止树选择与滚动互相触发的同步标记
        /// </summary>
        private bool _syncing;

        /// <summary>
        /// 初始加载标记
        /// </summary>
        private bool _loading = true;

        /// <summary>
        /// 打开窗口时的数据库路径（用于判断是否修改）
        /// </summary>
        private string _initialDbPath;

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
            _isAdmin = _permission.Can("admin.manage");

            CollectSections();
            CollectTreeNodes(tv_settings);

            LoadFontOptions();
            LoadLanguageOptions();
            LoadThemeOptions();
            LoadValues();

            // 数据库路径仅管理员可修改
            if (!_isAdmin)
            {
                txt_dbpath.IsReadOnly = true;
                btn_dbpathBrowse.IsEnabled = false;
                txt_dbpath.ToolTip = "仅管理员可修改数据库路径";
            }

            _loading = false;
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
        }

        private void CollectTreeNodes(DependencyObject parent)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is TreeViewItem item)
                {
                    if (item.Tag is string tag)
                    {
                        _treeNodes[tag] = item;
                    }
                }
                CollectTreeNodes(child);
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

            cb_fontSize.ItemsSource = new[] { 10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 28, 32 };
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

        private void LoadValues()
        {
            Models.AppSettings settings = _settings.Settings;
            cb_fontFamily.SelectedItem = settings.UI.FontFamily;
            cb_fontSize.Text = settings.UI.FontSize.ToString();

            _initialDbPath = _settings.GetDatabasePath();
            txt_dbpath.Text = _initialDbPath;
            _initialAtePath = _settings.GetAteDataPath();
            txt_ate.Text = _initialAtePath;
            _initialEmiPath = _settings.GetEmiDataPath();
            txt_emi.Text = _initialEmiPath;
            txt_schedule.Text = settings.Paths.SchedulePath;
            txt_requisition.Text = settings.Paths.RequisitionPath;
            txt_report.Text = settings.Paths.ReportPath;
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
            if (double.TryParse(cb_fontSize.Text, out double size) && size >= 8 && size <= 72)
            {
                settings.UI.FontSize = size;
            }
            else
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_InvalidFontSize"), LanguageService.Get("Cap_Info"));
                return false;
            }

            settings.Paths.SchedulePath = TrimOrNull(txt_schedule.Text);
            settings.Paths.RequisitionPath = TrimOrNull(txt_requisition.Text);
            settings.Paths.ReportPath = TrimOrNull(txt_report.Text);
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

            // 数据库路径：仅管理员，且仅在修改时保存（重启后生效）
            if (_isAdmin)
            {
                string dbPath = TrimOrNull(txt_dbpath.Text);
                if (!string.Equals(dbPath, _initialDbPath, StringComparison.OrdinalIgnoreCase))
                {
                    _settings.SetDatabasePath(dbPath);
                    _initialDbPath = dbPath;
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_DBPathSaved"), LanguageService.Get("Cap_Info"));
                }
            }
            return true;
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

        private void Btn_DbPathBrowse_Click(object sender, RoutedEventArgs e)
        {
            if (!_isAdmin)
            {
                return;
            }
            string dir = _pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectDBDir"), initPath: txt_dbpath.Text, isDir: true);
            if (dir != null)
            {
                txt_dbpath.Text = dir;
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
