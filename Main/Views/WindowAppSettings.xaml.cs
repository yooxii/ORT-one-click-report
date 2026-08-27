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
    /// 左侧树状目录 + 右侧设置详情，支持同步滚动与点击跳转；所有设置以 JSON 保存（AppSettingsService）。
    /// </summary>
    public partial class WindowAppSettings : Window
    {
        private readonly AppSettingsService _settings;
        private readonly IPathService _pathService;

        /// <summary>
        /// 防止树选择与滚动互相触发的同步标记
        /// </summary>
        private bool _syncing;

        /// <summary>
        /// 初始加载标记（加载期间不触发保存）
        /// </summary>
        private bool _loading = true;

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

            CollectSections();
            CollectTreeNodes(tv_settings);

            LoadFontOptions();
            LoadValues();

            _loading = false;
        }

        /* ###############################  收集  ################################ */

        private void CollectSections()
        {
            _sections.Add(("sec_ui", sec_ui));
            _sections.Add(("sec_font", sec_font));
            _sections.Add(("sec_paths", sec_paths));
            _sections.Add(("sec_schedule", sec_schedule));
            _sections.Add(("sec_requisition", sec_requisition));
            _sections.Add(("sec_report", sec_report));
            _sections.Add(("sec_ate", sec_ate));
            _sections.Add(("sec_emi", sec_emi));
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

        private void LoadValues()
        {
            Models.AppSettings settings = _settings.Settings;
            cb_fontFamily.SelectedItem = settings.UI.FontFamily;
            cb_fontSize.Text = settings.UI.FontSize.ToString();

            txt_schedule.Text = settings.Paths.SchedulePath;
            txt_requisition.Text = settings.Paths.RequisitionPath;
            txt_report.Text = settings.Paths.ReportPath;
            txt_ate.Text = settings.Paths.AteDataPath;
            txt_emi.Text = settings.Paths.EmiDataPath;
        }

        /* ###############################  字体  ################################ */

        private void Cb_FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading || cb_fontFamily.SelectedItem is not string family)
            {
                return;
            }
            _settings.Settings.UI.FontFamily = family;
            _settings.Save();
            _settings.ApplyFontToAll();
        }

        private void Cb_FontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading)
            {
                return;
            }
            if (double.TryParse(cb_fontSize.Text, out double size) && size >= 8 && size <= 72)
            {
                _settings.Settings.UI.FontSize = size;
                _settings.Save();
                _settings.ApplyFontToAll();
            }
        }

        /* ###############################  路径  ################################ */

        private void Txt_Path_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_loading || sender is not TextBox box || box.Tag is not string key)
            {
                return;
            }
            string value = string.IsNullOrWhiteSpace(box.Text) ? null : box.Text.Trim();
            switch (key)
            {
                case "SchedulePath": _settings.Settings.Paths.SchedulePath = value; break;
                case "RequisitionPath": _settings.Settings.Paths.RequisitionPath = value; break;
                case "ReportPath": _settings.Settings.Paths.ReportPath = value; break;
                case "AteDataPath": _settings.Settings.Paths.AteDataPath = value; break;
                case "EmiDataPath": _settings.Settings.Paths.EmiDataPath = value; break;
            }
            _settings.Save();
        }

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
            string dir = _pathService.OpenPathDialog("选择目录", initPath: target.Text, isDir: true);
            if (dir != null)
            {
                target.Text = dir;
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
