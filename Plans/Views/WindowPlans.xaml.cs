using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Reports.Models;
using ORT一键报告.Reports.Views;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowPlans.xaml 的交互逻辑：领退表/计划表两个 Tab 展示、单元格编辑、
    /// 右键菜单（编辑/删除/复制/粘贴/显示隐藏列）、机种联动、行号显示与列顺序持久化。
    /// </summary>
    public partial class WindowPlans : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly PlansViewModel _vm;
        private readonly AppSettingsService _appSettings;

        private static readonly string LayoutFile
            = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "plans_layout.json");

        public WindowPlans()
        {
            InitializeComponent();
            _vm = App.ServiceProvider.GetRequiredService<PlansViewModel>();
            _appSettings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            DataContext = _vm;
            SetupHeaderMenus();
            Loaded += (s, e) =>
            {
                RestoreColumnState();
                _vm.Refresh();
                // 菜单子项必须在菜单展开之前就存在：WPF 对没有子项的 MenuItem 不会展开，
                // 也就不会触发 SubmenuOpened（否则「排序/显示隐藏列」点开是空的）
                BuildSortMenu(menu_window_sort, ActiveGrid());
                BuildColumnMenu(menu_window_columns, ActiveGrid());
                RefreshHeaderStyles();
            };
            tabs.SelectionChanged += (s, e) =>
            {
                // 窗口菜单作用于当前 Tab，切换后同步为对应表格的字段/列
                BuildSortMenu(menu_window_sort, ActiveGrid());
                BuildColumnMenu(menu_window_columns, ActiveGrid());
            };
            // 右键菜单同理：在打开前构建子项
            dg_requisitions.ContextMenuOpening += (s, e) =>
            {
                BuildSortMenu(menu_req_sort, dg_requisitions);
                BuildColumnMenu(menu_req_columns, dg_requisitions);
            };
            dg_plans.ContextMenuOpening += (s, e) =>
            {
                BuildSortMenu(menu_plan_sort, dg_plans);
                BuildColumnMenu(menu_plan_columns, dg_plans);
            };
            Closing += (s, e) => SaveColumnState();
            // 报告文件夹扫描完成后，提示用户建立计划索引（每个窗口实例只提示一次）
            _vm.ReportScanCompleted += OnReportScanCompleted;
        }

        /* ###############################  计划索引提示  ################################ */

        /// <summary>是否已经提示过建立计划索引（避免每次刷新都弹）</summary>
        private bool _indexPrompted;

        /// <summary>
        /// 扫描完成后提示建立计划索引：用户同意即后台开始（可在其他客户端断点继续）
        /// </summary>
        private async void OnReportScanCompleted(int reportCount)
        {
            if (_indexPrompted || reportCount <= 0)
            {
                return;
            }
            try
            {
                PlanIndexService indexService = App.ServiceProvider.GetRequiredService<PlanIndexService>();
                if (indexService.IsRunning || indexService.GetLatestJob(_appSettings.ReportDir) != null)
                {
                    return; // 已经有索引任务（做过或正在做）就不再打扰
                }
                _indexPrompted = true;
                MessageBoxResult choice = MessageBox.Show(
                    string.Format(LanguageService.Get("Plans_Msg_IndexPrompt"), reportCount),
                    LanguageService.Get("Plans_Msg_IndexPromptTitle"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (choice != MessageBoxResult.Yes)
                {
                    return;
                }
                PlanIndexRunResult result = await indexService.RunAsync(_appSettings.ReportDir,
                    App.ServiceProvider.GetRequiredService<IPermissionService>().CurrentUser);
                ToastService.Show(result.Completed || result.Started
                    ? result.Message
                    : string.Format(LanguageService.Get("PlanIndex_Msg_IndexFailedFormat"), result.Message),
                    result.Started ? ToastType.Info : ToastType.Warning);
            }
            catch (Exception ex)
            {
                _logger.Warn($"提示建立计划索引失败: {ex.Message}");
            }
        }

        /* ###############################  视图菜单开关  ################################ */

        /// <summary>
        /// 视图菜单里的显示开关（工具栏/搜索）：勾选状态即显示状态
        /// </summary>
        private void Menu_View_Toggle_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem item)
            {
                return;
            }
            if (ReferenceEquals(item, menu_view_toolbar))
            {
                bar_toolbar.Visibility = item.IsChecked ? Visibility.Visible : Visibility.Collapsed;
            }
            else if (ReferenceEquals(item, menu_view_search))
            {
                panel_search.Visibility = item.IsChecked ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        /* ###############################  表头右键：排序 + 本列筛选  ################################ */

        /// <summary>筛选菜单里最多列出的可选值个数（超出只提示，避免菜单过长）</summary>
        private const int MaxHeaderFilterValues = 200;

        /// <summary>筛选项超过这个数量时，在筛选菜单里显示搜索框</summary>
        private const int FilterSearchThreshold = 10;

        /// <summary>带搜索框时最多创建的可选值项（搜索可在其中查找，避免一次性建上千个菜单项）</summary>
        private const int MaxSearchableFilterValues = 1000;

        /// <summary>已筛选列的列头样式（加粗+主题色，作为 Excel 漏斗的替代提示）</summary>
        private readonly Dictionary<DataGrid, Style> _filteredHeaderStyles = [];

        /// <summary>
        /// 准备「已筛选列头」样式（加粗 + 主题色，替代 Excel 的漏斗标记）
        /// </summary>
        private void SetupHeaderMenus()
        {
            Style baseStyle = TryFindResource(typeof(DataGridColumnHeader)) as Style;
            Style filtered = baseStyle != null
                ? new Style(typeof(DataGridColumnHeader), baseStyle)
                : new Style(typeof(DataGridColumnHeader));
            filtered.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Bold));
            if (TryFindResource("PrimaryBrush") is Brush primary)
            {
                filtered.Setters.Add(new Setter(Control.ForegroundProperty, primary));
                filtered.Setters.Add(new Setter(Control.BorderBrushProperty, primary));
            }
            _filteredHeaderStyles[dg_requisitions] = filtered;
            _filteredHeaderStyles[dg_plans] = filtered;
        }

        /// <summary>
        /// 构建表头右键菜单：该列的升/降序、该列的筛选（值多选，类似资源管理器/Excel）、清除筛选
        /// </summary>
        private void BuildHeaderMenu(ContextMenu menu, DataGrid grid, DataGridColumn column)
        {
            bool isPlan = ReferenceEquals(grid, dg_plans);
            string property = column.SortMemberPath;
            string label = column.Header?.ToString() ?? "?";

            menu.Items.Clear();
            menu.Items.Add(new MenuItem { Header = label, IsEnabled = false, FontWeight = FontWeights.Bold });
            menu.Items.Add(new Separator());

            MenuItem ascending = new() { Header = LanguageService.Get("Common_Ascending") };
            ascending.Click += (s, args) => ApplySort(grid, property, ListSortDirection.Ascending);
            menu.Items.Add(ascending);
            MenuItem descending = new() { Header = LanguageService.Get("Common_Descending") };
            descending.Click += (s, args) => ApplySort(grid, property, ListSortDirection.Descending);
            menu.Items.Add(descending);

            if (string.IsNullOrWhiteSpace(property))
            {
                return;
            }

            ColumnFilter filter = _vm.GetColumnFilter(isPlan, property);
            List<string> values = _vm.GetColumnValues(isPlan, property);
            menu.Items.Add(new Separator());

            MenuItem filterRoot = new() { Header = LanguageService.Get("Common_Filter") };
            MenuItem all = new()
            {
                Header = LanguageService.Get("Menu_FilterAll"),
                IsCheckable = true,
                IsChecked = filter == null || !filter.IsActive
            };
            all.Click += (s, args) => ApplyColumnFilter(isPlan, property, label, values, null);
            filterRoot.Items.Add(all);

            if (values.Count > 0)
            {
                // 未筛选时所有值都是勾选状态（"全部包含"，与资源管理器/Excel 一致），
                // 取消勾选某个值即为"排除它"
                bool allIncluded = filter == null || !filter.IsActive;
                filterRoot.Items.Add(new Separator());

                // 值多时提供搜索框：空查询只显示前 MaxHeaderFilterValues 项（保持菜单不过长），
                // 输入关键字后在全部候选中筛选（可搜到前 200 项之外的取值）
                bool withSearch = values.Count > FilterSearchThreshold;
                List<(MenuItem Item, string Value)> valueItems = [];
                int created = 0;
                int createLimit = withSearch ? MaxSearchableFilterValues : MaxHeaderFilterValues;
                foreach (string value in values)
                {
                    if (created++ >= createLimit)
                    {
                        filterRoot.Items.Add(new MenuItem
                        {
                            Header = string.Format(LanguageService.Get("Menu_FilterMoreFormat"), createLimit),
                            IsEnabled = false
                        });
                        break;
                    }
                    MenuItem valueItem = new()
                    {
                        Header = value,
                        IsCheckable = true,
                        IsChecked = allIncluded || filter.Selected.Contains(value)
                    };
                    string captured = value;
                    valueItem.Click += (s, args) =>
                    {
                        List<string> selected = allIncluded ? [.. values] : [.. filter.Selected];
                        if (valueItem.IsChecked)
                        {
                            if (!selected.Contains(captured))
                            {
                                selected.Add(captured);
                            }
                        }
                        else
                        {
                            selected.Remove(captured);
                        }
                        ApplyColumnFilter(isPlan, property, label, values, selected);
                    };
                    filterRoot.Items.Add(valueItem);
                    valueItems.Add((valueItem, value));
                }

                if (withSearch)
                {
                    AddFilterSearchBox(filterRoot, valueItems);
                }
            }
            menu.Items.Add(filterRoot);

            menu.Items.Add(new Separator());
            MenuItem clearColumn = new()
            {
                Header = LanguageService.Get("Menu_ClearColumnFilter"),
                IsEnabled = filter?.IsActive == true
            };
            clearColumn.Click += (s, args) => ApplyColumnFilter(isPlan, property, label, values, null);
            menu.Items.Add(clearColumn);

            MenuItem clearAll = new()
            {
                Header = LanguageService.Get("Menu_ClearAllFilters"),
                IsEnabled = _vm.HasFilters(isPlan)
            };
            clearAll.Click += (s, args) =>
            {
                _vm.ClearAllFilters(isPlan);
                RefreshHeaderStyles();
            };
            menu.Items.Add(clearAll);
        }

        /// <summary>
        /// 在筛选菜单里插入搜索框（筛选项多时使用）：输入关键字即时过滤候选值；
        /// 空查询时仍只显示前 MaxHeaderFilterValues 项，输入后可在全部候选（最多 MaxSearchableFilterValues 项）里查找。
        /// 说明：TextBox 放在 MenuItem 的 Header 里，宿主要设 StaysOpenOnClick 且不可聚焦，
        /// 并在子菜单打开时把焦点交给输入框，否则键盘输入会被菜单当成导航。
        /// </summary>
        private void AddFilterSearchBox(MenuItem filterRoot, List<(MenuItem Item, string Value)> valueItems)
        {
            TextBox search = new()
            {
                Width = 190,
                Padding = new Thickness(4, 1, 4, 1)
            };
            // 供无障碍与自动化识别（同时也便于测试脚本精确定位这个输入框）
            AutomationProperties.SetName(search, LanguageService.Get("Menu_FilterSearchHint"));
            TextBlock hint = new()
            {
                Text = LanguageService.Get("Menu_FilterSearchHint"),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false,
                Foreground = TryFindResource("TextSecondaryBrush") as Brush ?? Brushes.Gray
            };
            Grid searchPanel = new() { Width = 190, Margin = new Thickness(4, 2, 4, 2) };
            searchPanel.Children.Add(search);
            searchPanel.Children.Add(hint);

            MenuItem searchHost = new()
            {
                Header = searchPanel,
                StaysOpenOnClick = true,
                Focusable = false
            };
            searchHost.PreviewMouseLeftButtonDown += (s, args) =>
            {
                search.Focus();
                args.Handled = true;
            };
            filterRoot.SubmenuOpened += (s, args) =>
                search.Dispatcher.BeginInvoke(new Action(() => search.Focus()), DispatcherPriority.Input);

            MenuItem noMatch = new()
            {
                Header = LanguageService.Get("Menu_FilterNoMatch"),
                IsEnabled = false,
                Visibility = Visibility.Collapsed
            };

            void ApplySearch()
            {
                string query = search.Text?.Trim() ?? "";
                hint.Visibility = search.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
                int visible = 0;
                for (int i = 0; i < valueItems.Count; i++)
                {
                    (MenuItem item, string value) = valueItems[i];
                    bool match = query.Length > 0
                        ? value?.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                        : i < MaxHeaderFilterValues;
                    item.Visibility = match ? Visibility.Visible : Visibility.Collapsed;
                    if (match)
                    {
                        visible++;
                    }
                }
                noMatch.Visibility = visible == 0 ? Visibility.Visible : Visibility.Collapsed;
            }

            search.TextChanged += (s, args) => ApplySearch();

            int insertAt = Math.Min(2, filterRoot.Items.Count);
            filterRoot.Items.Insert(insertAt, searchHost);
            filterRoot.Items.Insert(insertAt + 1, new Separator());
            filterRoot.Items.Add(noMatch);
            ApplySearch();
        }

        /// <summary>
        /// 应用某列筛选并刷新表头样式（selected 为空表示该列恢复"全部"）
        /// </summary>
        private void ApplyColumnFilter(bool isPlan, string property, string label, IReadOnlyList<string> values, IReadOnlyList<string> selected)
        {
            _vm.SetColumnFilter(isPlan, property, label, values, selected ?? []);
            RefreshHeaderStyles();
        }

        /// <summary>
        /// 按当前筛选状态刷新两个表格的列头样式（已筛选的列头加粗并用主题色）
        /// </summary>
        private void RefreshHeaderStyles()
        {
            ApplyHeaderStyles(dg_requisitions, false);
            ApplyHeaderStyles(dg_plans, true);
        }

        private void ApplyHeaderStyles(DataGrid grid, bool isPlan)
        {
            if (!_filteredHeaderStyles.TryGetValue(grid, out Style filtered))
            {
                return;
            }
            foreach (DataGridColumn column in grid.Columns)
            {
                column.HeaderStyle = _vm.IsColumnFiltered(isPlan, column.SortMemberPath) ? filtered : null;
            }
        }

        /* ###############################  行号  ################################ */

        /// <summary>
        /// 当前单元格所在行的高亮改用语义资源键 TableRowSelectedBrush（动态引用，随主题切换），
        /// 不再硬编码浅灰，避免深色主题下浅色文字不可读。
        /// </summary>
        private DataGridRow _lastReqHighlightRow;
        private DataGridRow _lastPlanHighlightRow;

        /// <summary>计划表下拉列（浏览时只显示文本，双击编辑才出现下拉框）</summary>
        private static readonly System.Collections.Generic.HashSet<string> PlanComboColumns =
            new() { "Product", "Customer", "Stage", "TestItem", "Status" };

        /// <summary>下拉编辑开始时的快照（行、属性、原值），用于结束时判断是否修改并提示</summary>
        private (Plan item, string prop, string original)? _planComboEditSnapshot;

        private static string GetPropString(object obj, string prop)
            => obj?.GetType().GetProperty(prop)?.GetValue(obj)?.ToString();

        /// <summary>
        /// 下拉列进入编辑时记录原值，供 CellEditEnding 比较是否修改
        /// </summary>
        private void Dg_Plans_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            string prop = e.Column.SortMemberPath;
            _planComboEditSnapshot = e.Row.Item is Plan plan && prop != null && PlanComboColumns.Contains(prop)
                ? (plan, prop, GetPropString(plan, prop))
                : null;
        }

        private void Dg_Requisitions_CurrentCellChanged(object sender, EventArgs e)
            => UpdateCurrentRowHighlight(dg_requisitions, ref _lastReqHighlightRow);

        private void Dg_Plans_CurrentCellChanged(object sender, EventArgs e)
            => UpdateCurrentRowHighlight(dg_plans, ref _lastPlanHighlightRow);

        /// <summary>
        /// 当前单元格变化时高亮其所在行（与单元格选中效果同时显示，两色区分）
        /// </summary>
        private static void UpdateCurrentRowHighlight(DataGrid grid, ref DataGridRow lastRow)
        {
            if (lastRow != null)
            {
                lastRow.ClearValue(DataGridRow.BackgroundProperty);
                lastRow = null;
            }
            if (grid.CurrentCell.IsValid && grid.CurrentCell.Item != null
                && grid.ItemContainerGenerator.ContainerFromItem(grid.CurrentCell.Item) is DataGridRow row)
            {
                // 动态资源引用：跟随当前主题的 TableRowSelectedBrush，切换主题自动刷新
                row.SetResourceReference(DataGridRow.BackgroundProperty, "TableRowSelectedBrush");
                lastRow = row;
            }
        }

        private void Dg_Requisitions_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void Dg_Plans_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        /* ###############################  单元格编辑结束  ################################ */

        private void Dg_Requisitions_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is Requisition req)
            {
                _vm.NotifyPendingChanged();
                _vm.StatusMessage = _vm.PendingText;
            }
        }

        private void Dg_Plans_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is not Plan plan)
            {
                return;
            }
            // 校验字典/格式字段
            string column = e.Column.SortMemberPath;
            if (column == "JobNo")
            {
                string error = _vm.ValidateField("JobNo", plan.JobNo);
                if (error != null)
                {
                    _ = MessageBox.Show(error, LanguageService.Get("Cap_FormatValidationFailed"));
                    e.Cancel = true;
                    return;
                }
            }
            if (column == "Status")
            {
                string error = _vm.ValidateField("Status", plan.Status);
                if (error != null)
                {
                    _ = MessageBox.Show(error, LanguageService.Get("Cap_ValidationFailed"));
                    e.Cancel = true;
                    return;
                }
            }
            if (column == "TestItem")
            {
                string error = _vm.ValidateField("TestItem", plan.TestItem);
                if (error != null)
                {
                    _ = MessageBox.Show(error, LanguageService.Get("Cap_ValidationFailed"));
                    e.Cancel = true;
                    return;
                }
                // 测试项目联动：同步负责人/试验时间/结束日期
                _vm.AutoFillByTestItem(plan);
            }
            // 机种联动：修改机种名称后带出产品别/客户别（仅填充空字段）
            if (column == "ModelName")
            {
                _vm.AutoFillByModel(plan);
            }

            // 下拉列：选择与原值不同时弹出修改提示（Toast）
            if (PlanComboColumns.Contains(column)
                && _planComboEditSnapshot is { } snap
                && snap.item == plan && snap.prop == column)
            {
                string newVal = GetPropString(plan, column);
                if (!string.Equals(snap.original, newVal, StringComparison.Ordinal))
                {
                    ToastService.Show(string.Format(LanguageService.Get("Toast_ValueChanged"),
                        column, snap.original ?? "", newVal ?? ""), ToastType.Info);
                }
            }
            _planComboEditSnapshot = null;

            _vm.NotifyPendingChanged();
            _vm.StatusMessage = _vm.PendingText;
        }

        /* ###############################  右键单元格定位  ################################ */

        private void Dg_Requisitions_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ShowHeaderMenuIfOnHeader(dg_requisitions, e))
            {
                return;
            }
            SelectCellUnderMouse(sender as DataGrid, e);
        }

        private void Dg_Plans_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (ShowHeaderMenuIfOnHeader(dg_plans, e))
            {
                return;
            }
            SelectCellUnderMouse(sender as DataGrid, e);
        }

        /// <summary>
        /// 右键落在表头上时改为弹出「该列」的排序/筛选菜单（表格自身的行菜单只用于单元格）。
        /// 说明：不通过列头样式挂 ContextMenu——那样会被 DataGrid 级右键菜单截走。
        /// </summary>
        private bool ShowHeaderMenuIfOnHeader(DataGrid grid, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is not DependencyObject source)
            {
                return false;
            }
            DataGridColumnHeader header = FindAncestor<DataGridColumnHeader>(source);
            if (header?.Column == null)
            {
                return false;
            }
            ContextMenu menu = new() { PlacementTarget = header, Placement = PlacementMode.Bottom };
            BuildHeaderMenu(menu, grid, header.Column);
            menu.IsOpen = true;
            e.Handled = true;
            return true;
        }

        /// <summary>
        /// 沿可视树向上查找指定类型的祖先
        /// </summary>
        private static T FindAncestor<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T match)
                {
                    return match;
                }
                node = node is Visual ? VisualTreeHelper.GetParent(node) : LogicalTreeHelper.GetParent(node);
            }
            return null;
        }

        /// <summary>
        /// 右键点击时定位到鼠标所在单元格并设为当前单元格（保证右键菜单操作针对正确的单元格）
        /// </summary>
        private static void SelectCellUnderMouse(DataGrid grid, MouseButtonEventArgs e)
        {
            if (grid == null)
            {
                return;
            }
            System.Windows.DependencyObject dep = e.OriginalSource as System.Windows.DependencyObject;
            while (dep != null && dep is not DataGridCell)
            {
                dep = VisualTreeHelper.GetParent(dep);
            }
            if (dep is DataGridCell cell && cell.DataContext == grid.CurrentItem)
            {
                cell.Focus();
                grid.CurrentCell = new DataGridCellInfo(cell);
            }
        }

        /* ###############################  右键菜单：编辑/删除  ################################ */

        private void Menu_EditRequisition_Click(object sender, RoutedEventArgs e)
        {
            // 仅选择单个单元格时也可进入：优先取右键单元格所在行
            _vm.SelectedRequisition = (dg_requisitions.CurrentCell.Item as Requisition)
                ?? dg_requisitions.SelectedItem as Requisition;
            if (_vm.EditRequisitionCommand.CanExecute(null))
            {
                _vm.EditRequisitionCommand.Execute(null);
            }
        }

        private void Menu_EditPlan_Click(object sender, RoutedEventArgs e)
        {
            _vm.SelectedPlan = (dg_plans.CurrentCell.Item as Plan)
                ?? dg_plans.SelectedItem as Plan;
            if (_vm.EditPlanCommand.CanExecute(null))
            {
                _vm.EditPlanCommand.Execute(null);
            }
        }

        private void Menu_DeleteRequisition_Click(object sender, RoutedEventArgs e)
        {
            if (dg_requisitions.SelectedItem is Requisition req)
            {
                _vm.DeleteRequisitionCommand.Execute(req);
            }
        }

        private void Menu_DeletePlan_Click(object sender, RoutedEventArgs e)
        {
            if (dg_plans.SelectedItem is Plan plan)
            {
                _vm.DeletePlanCommand.Execute(plan);
            }
        }

        /* ###############################  右键菜单：报告对应  ################################ */

        private Plan CurrentPlan => (dg_plans.CurrentCell.Item as Plan) ?? dg_plans.SelectedItem as Plan;

        private void Menu_OpenReportFolder_Click(object sender, RoutedEventArgs e)
        {
            ToastService.WarnIfReportPathEmpty();
            Plan plan = CurrentPlan;
            ReportLink link = plan == null ? null : _vm.FindReportLink(plan.JobNo);
            if (link == null || !System.IO.Directory.Exists(link.ReportDir))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FolderNotFound"), LanguageService.Get("Cap_Info"));
                return;
            }
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", $"\"{link.ReportDir}\"");
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"打开失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void Menu_OpenReportOverview_Click(object sender, RoutedEventArgs e)
        {
            ToastService.WarnIfReportPathEmpty();
            Plan plan = CurrentPlan;
            ReportLink link = plan == null ? null : _vm.FindReportLink(plan.JobNo);
            if (link == null || !System.IO.File.Exists(link.OverviewFile))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_OverviewNotFound"), LanguageService.Get("Cap_Info"));
                return;
            }
            try
            {
                System.Diagnostics.Process.Start(link.OverviewFile);
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"打开失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        /// <summary>
        /// 打开一键报告，并将选中的计划记录（含对应领退数据）携带到报告服务；
        /// 同时构建预填的 BurnInReportModel 实例，下游生成报告时可直接使用该模型。
        /// </summary>
        private void Menu_OpenMainReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Plan plan = CurrentPlan;
                if (plan == null)
                {
                    return;
                }
                ReportService reportService = App.ServiceProvider.GetRequiredService<ReportService>();
                reportService.MatchedPlan = plan;
                Requisition req = _vm.FindRequisitionForPlan(plan);
                reportService.MatchedRequisition = req;
                // 该计划绑定的报告文件夹：报告页的测试信息/测试图片改为从这里按报告类型读取
                reportService.MatchedReportDir = _vm.FindReportLink(plan.JobNo)?.ReportDir;
                // 标记"从计划表进入"，并清掉上一次的预填结果，避免残留串味
                reportService.EnteredFromPlan = true;
                reportService.UUTInfos = null;
                reportService.PrefilledReportModel = null;

                // 报告类型文件检查：缺失的类型提示用户；表头信息随后只用报告文件夹里读到的，
                // 读不到就置空（只保留计划表+领用表能提供的 序列号/工令/周期/版本）
                List<string> missingTypes = [];
                if (string.IsNullOrWhiteSpace(reportService.MatchedReportDir))
                {
                    missingTypes.AddRange(["Thermal Shock", "Burn In", "EMI"]);
                    _ = MessageBox.Show(
                        "该计划还没有绑定的报告文件夹，表头信息（测试人/审核人/项目名/阶段/图片等）将留空。\n"
                        + "请在设置里配置报告路径后，在计划表界面刷新以重新扫描报告文件夹。",
                        LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    foreach (string type in new[] { "Thermal Shock", "Burn In", "EMI" })
                    {
                        string typeFile = ORT一键报告.Utils.Report.GetTemplatePath(reportService.MatchedReportDir, type);
                        if (string.IsNullOrWhiteSpace(typeFile) || !File.Exists(typeFile))
                        {
                            missingTypes.Add(type);
                        }
                    }
                    if (missingTypes.Count > 0)
                    {
                        _ = MessageBox.Show(
                            $"绑定的报告文件夹里缺少以下报告：{string.Join("、", missingTypes)}。\n"
                            + "这些报告的表头信息（测试人/审核人/项目名/阶段/图片等）将留空，只填序列号/工令/周期/版本。\n"
                            + $"文件夹：{reportService.MatchedReportDir}",
                            LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                _logger.Info($"一键报告：报告文件夹={reportService.MatchedReportDir ?? "(未绑定)"}，缺少的报告={string.Join(",", missingTypes)}");

                // 预填 UUTInfos（用户未读取报告概览时也能让 Tab 有数据）
                // 只携带计划表+领用表能提供的信息：序列号/工令/周期/版本(DC)/测试项目，
                // 表头文字（测试人/审核人/项目名/阶段/描述/图片）一律以绑定的报告文件夹里的报告为准
                List<string> snList = ParseSnLines(req?.SN);
                reportService.UUTInfos = new UUTInfoFromExcel
                {
                    SNs = snList,
                    WorkOrder = req?.WorkOrder ?? "",
                    Revision = req?.Rev ?? "",
                    DC = req?.DC ?? "",
                    // 测试周期：报告文件夹里没有该报告时，表头的"测试周期"用计划表的起止日期
                    TestStart = plan.StartDate,
                    TestItems = string.IsNullOrWhiteSpace(plan.TestItem)
                        ? []
                        : [new TestItemInfo { TestItemName = plan.TestItem, Date = plan.StartDate?.ToString("yyyy/M/d") ?? "" }]
                };

                // 构建 BurnInReportModel（ORT 最常用的 Burn In 报告类型）：
                // 表头只保留计划表能给的"测试周期"（StartDate/EndDate），其余（测试人/审核人/项目名/阶段/
                // 描述/图片）一律以绑定的报告文件夹里的本地报告文件为准，读不到就置空 —— 不再用计划表文案兜底
                int testTimeDays = 7; // Burn In 默认 7 天
                DateTime periodStart = plan.StartDate ?? DateTime.Now;
                ReportHeaderData planHeader = new()
                {
                    TestStart = periodStart,
                    TestEnd = plan.EndDate ?? periodStart.AddDays(testTimeDays)
                };
                string burnInReportFile = ORT一键报告.Utils.Report.GetTemplatePath(reportService.MatchedReportDir, "Burn In");
                ReportHeaderData reportFolderHeader = ORT一键报告.Utils.Report.ReadHeaderData(burnInReportFile, testTimeDays, planHeader);
                if (reportFolderHeader != null)
                {
                    // 报告文件里没写周期时，仍用计划表的周期
                    reportFolderHeader.TestStart = reportFolderHeader.TestStart == default ? planHeader.TestStart : reportFolderHeader.TestStart;
                    _logger.Info($"一键报告表头取自本地报告文件：{burnInReportFile}");
                }
                else
                {
                    _logger.Info("未读到本地报告文件表头，表头除周期外留空");
                }
                BurnInReportModel prefilled = new()
                {
                    Header = reportFolderHeader ?? planHeader,
                    UUTSource = new UUTSourceData
                    {
                        SNs = snList,
                        WorkOrder = req?.WorkOrder,
                        Revision = req?.Rev,
                        DC = req?.DC,
                        TestItems = string.IsNullOrWhiteSpace(plan.TestItem)
                            ? []
                            : [new TestItemInfo { TestItemName = plan.TestItem, Date = plan.StartDate?.ToString("yyyy/M/d") ?? "" }]
                    },
                    RootReportPath = reportService.RootPath,
                    TempPath = reportService.TempPath
                };
                reportService.PrefilledReportModel = prefilled;

                ToastService.WarnIfReportPathEmpty();
                WindowMainReport window = new();
                window.Show();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "从计划表打开一键报告失败");
                _ = MessageBox.Show($"打开一键报告失败：{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        /// <summary>
        /// 右键"建立报告模板"：只对还没有报告文件夹的记录有意义；带着该行机种打开报告模板工具
        /// </summary>
        private void Menu_BuildReportTemplate_Click(object sender, RoutedEventArgs e)
        {
            Plan plan = CurrentPlan;
            if (plan == null)
            {
                _ = MessageBox.Show(LanguageService.Get("Plans_Msg_TemplateSelectPlan"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (!string.IsNullOrWhiteSpace(plan.JobNo) && _vm.FindReportLink(plan.JobNo) != null)
            {
                _ = MessageBox.Show(LanguageService.Get("Plans_Msg_TemplateHasReport"), LanguageService.Get("Cap_Info"));
                return;
            }
            OpenReportTemplateWindow(plan);
        }

        /// <summary>
        /// 工具菜单/工具栏的通用入口：打开报告模板工具（不带机种，由用户在窗口里选择）
        /// </summary>
        private void Menu_ReportTemplate_Click(object sender, RoutedEventArgs e) => OpenReportTemplateWindow(null);

        private void OpenReportTemplateWindow(Plan plan)
        {
            try
            {
                WindowReportTemplate window = new();
                if (plan != null)
                {
                    window.PrefillFromPlan(plan, _vm.FindRequisitionForPlan(plan));
                }
                window.Show();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "打开报告模板工具失败");
                _ = MessageBox.Show($"打开报告模板工具失败：{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        /// <summary>
        /// 工具菜单/工具栏：建立计划索引（后台执行，可在其他客户端断点继续）
        /// </summary>
        private async void Menu_PlanIndex_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PlanIndexService indexService = App.ServiceProvider.GetRequiredService<PlanIndexService>();
                if (indexService.IsRunning)
                {
                    ToastService.Show(LanguageService.Get("PlanIndex_Msg_IndexRunning"), ToastType.Info);
                    return;
                }
                string root = _appSettings.ReportDir;
                if (string.IsNullOrWhiteSpace(root))
                {
                    ToastService.Show(LanguageService.Get("PlanIndex_NoRoot"), ToastType.Warning);
                    return;
                }
                PlanIndexRunResult result = await indexService.RunAsync(root,
                    App.ServiceProvider.GetRequiredService<IPermissionService>().CurrentUser);
                ToastService.Show(string.IsNullOrWhiteSpace(result.Message)
                    ? LanguageService.Get("PlanIndex_Msg_OtherClient")
                    : result.Message, result.Started ? ToastType.Info : ToastType.Warning);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "建立计划索引失败");
                _ = MessageBox.Show(string.Format(LanguageService.Get("PlanIndex_Msg_IndexFailedFormat"), ex.Message),
                    LanguageService.Get("Cap_Error"));
            }
        }

        /// <summary>
        /// 解析 SN 字符串（支持多行换行）为 SN 列表
        /// </summary>
        private static List<string> ParseSnLines(string snText)
        {
            if (string.IsNullOrWhiteSpace(snText))
            {
                return [];
            }
            return snText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        /* ###############################  右键菜单：复制/粘贴  ################################ */

        private void Menu_CopyReqRow_Click(object sender, RoutedEventArgs e)
        {
            if (dg_requisitions.SelectedItem is Requisition req)
            {
                Clipboard.SetText(RequisitionToTsv(req));
            }
        }

        private void Menu_CopyPlanRow_Click(object sender, RoutedEventArgs e)
        {
            if (dg_plans.SelectedItem is Plan plan)
            {
                Clipboard.SetText(PlanToTsv(plan));
            }
        }

        private void Menu_CopyReqCell_Click(object sender, RoutedEventArgs e)
        {
            if (dg_requisitions.CurrentCell.Column != null && dg_requisitions.SelectedItem is Requisition req)
            {
                Clipboard.SetText(GetRequisitionFieldValue(req, dg_requisitions.CurrentCell.Column.SortMemberPath) ?? "");
            }
        }

        private void Menu_CopyPlanCell_Click(object sender, RoutedEventArgs e)
        {
            if (dg_plans.CurrentCell.Column != null && dg_plans.SelectedItem is Plan plan)
            {
                Clipboard.SetText(GetPlanFieldValue(plan, dg_plans.CurrentCell.Column.SortMemberPath) ?? "");
            }
        }

        private void Menu_PasteReq_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.CanGridEdit)
            {
                return;
            }
            string text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }
            if (dg_requisitions.CurrentItem is not Requisition start)
            {
                return;
            }
            string[][] rows = ParseTsv(text);
            List<DataGridColumn> columns = dg_requisitions.Columns.ToList();
            int startCol = dg_requisitions.CurrentCell.Column?.DisplayIndex ?? 0;
            int startRow = dg_requisitions.Items.IndexOf(start);
            for (int i = 0; i < rows.Length; i++)
            {
                int rowIndex = startRow + i;
                if (rowIndex >= dg_requisitions.Items.Count)
                {
                    break;
                }
                if (dg_requisitions.Items[rowIndex] is not Requisition target)
                {
                    continue;
                }
                for (int j = 0; j < rows[i].Length; j++)
                {
                    int colIndex = startCol + j;
                    if (colIndex >= columns.Count)
                    {
                        break;
                    }
                    SetRequisitionFieldValue(target, columns[colIndex].SortMemberPath, rows[i][j]);
                }
            }
            _vm.NotifyPendingChanged();
        }

        private void Menu_PastePlan_Click(object sender, RoutedEventArgs e)
        {
            if (!_vm.CanGridEdit)
            {
                return;
            }
            string text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }
            if (dg_plans.CurrentItem is not Plan start)
            {
                return;
            }
            string[][] rows = ParseTsv(text);
            List<DataGridColumn> columns = dg_plans.Columns.ToList();
            int startCol = dg_plans.CurrentCell.Column?.DisplayIndex ?? 0;
            int startRow = dg_plans.Items.IndexOf(start);
            for (int i = 0; i < rows.Length; i++)
            {
                int rowIndex = startRow + i;
                if (rowIndex >= dg_plans.Items.Count)
                {
                    break;
                }
                if (dg_plans.Items[rowIndex] is not Plan target)
                {
                    continue;
                }
                for (int j = 0; j < rows[i].Length; j++)
                {
                    int colIndex = startCol + j;
                    if (colIndex >= columns.Count)
                    {
                        break;
                    }
                    string error = SetPlanFieldValue(target, columns[colIndex].SortMemberPath, rows[i][j]);
                    if (error != null)
                    {
                        _ = MessageBox.Show($"{columns[colIndex].Header}: {error}", LanguageService.Get("Cap_PasteValidationFailed"));
                    }
                }
            }
            _vm.NotifyPendingChanged();
        }

        private static string[][] ParseTsv(string text)
        {
            string[] lines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            string[][] rows = new string[lines.Length][];
            for (int i = 0; i < lines.Length; i++)
            {
                rows[i] = lines[i].Split('\t');
            }
            return rows;
        }

        private static string RequisitionToTsv(Requisition r)
            => string.Join("\t",
                r.RequisitionDate?.ToString("yyyy/M/d"), r.RequisitionNo, r.ModelName, r.OutQty,
                r.SN ?? r.SnFilePath, r.DC, r.Rev, r.WorkOrder, r.ReturnRtOrder, r.ReturnQty,
                r.LineNo, r.ReturnDate?.ToString("yyyy/M/d"), r.StockInNo, r.StockInQty,
                r.StockInDate?.ToString("yyyy/M/d"), r.Remark);

        private static string PlanToTsv(Plan p)
            => string.Join("\t",
                p.JobNo, p.Product, p.Customer, p.ModelName, p.Stage, p.TestItem, p.SampleSize,
                p.TestPeriod, p.Owner, p.StartDate?.ToString("yyyy/M/d"), p.EndDate?.ToString("yyyy/M/d"),
                p.Status, p.Remark);

        private static string GetRequisitionFieldValue(Requisition r, string field) => field switch
        {
            nameof(Requisition.RequisitionDate) => r.RequisitionDate?.ToString("yyyy/M/d"),
            nameof(Requisition.RequisitionNo) => r.RequisitionNo,
            nameof(Requisition.ModelName) => r.ModelName,
            nameof(Requisition.OutQty) => r.OutQty,
            nameof(Requisition.SN) => r.SN ?? r.SnFilePath,
            nameof(Requisition.DC) => r.DC,
            nameof(Requisition.Rev) => r.Rev,
            nameof(Requisition.WorkOrder) => r.WorkOrder,
            nameof(Requisition.ReturnRtOrder) => r.ReturnRtOrder,
            nameof(Requisition.ReturnQty) => r.ReturnQty,
            nameof(Requisition.LineNo) => r.LineNo,
            nameof(Requisition.ReturnDate) => r.ReturnDate?.ToString("yyyy/M/d"),
            nameof(Requisition.StockInNo) => r.StockInNo,
            nameof(Requisition.StockInQty) => r.StockInQty,
            nameof(Requisition.StockInDate) => r.StockInDate?.ToString("yyyy/M/d"),
            nameof(Requisition.Remark) => r.Remark,
            _ => null
        };

        private static string GetPlanFieldValue(Plan p, string field) => field switch
        {
            nameof(Plan.JobNo) => p.JobNo,
            nameof(Plan.Product) => p.Product,
            nameof(Plan.Customer) => p.Customer,
            nameof(Plan.ModelName) => p.ModelName,
            nameof(Plan.Stage) => p.Stage,
            nameof(Plan.TestItem) => p.TestItem,
            nameof(Plan.SampleSize) => p.SampleSize,
            nameof(Plan.TestPeriod) => p.TestPeriod,
            nameof(Plan.Owner) => p.Owner,
            nameof(Plan.StartDate) => p.StartDate?.ToString("yyyy/M/d"),
            nameof(Plan.EndDate) => p.EndDate?.ToString("yyyy/M/d"),
            nameof(Plan.Status) => p.Status,
            nameof(Plan.Remark) => p.Remark,
            _ => null
        };

        private static void SetRequisitionFieldValue(Requisition r, string field, string value)
        {
            if (value == "")
            {
                value = null;
            }
            switch (field)
            {
                case nameof(Requisition.RequisitionDate): r.RequisitionDate = ParseDate(value); break;
                case nameof(Requisition.RequisitionNo): r.RequisitionNo = value; break;
                case nameof(Requisition.ModelName): r.ModelName = value; break;
                case nameof(Requisition.OutQty): r.OutQty = value; break;
                case nameof(Requisition.SN): r.SN = value; break;
                case nameof(Requisition.DC): r.DC = value; break;
                case nameof(Requisition.Rev): r.Rev = value; break;
                case nameof(Requisition.WorkOrder): r.WorkOrder = value; break;
                case nameof(Requisition.ReturnRtOrder): r.ReturnRtOrder = value; break;
                case nameof(Requisition.ReturnQty): r.ReturnQty = value; break;
                case nameof(Requisition.LineNo): r.LineNo = value; break;
                case nameof(Requisition.ReturnDate): r.ReturnDate = ParseDate(value); break;
                case nameof(Requisition.StockInNo): r.StockInNo = value; break;
                case nameof(Requisition.StockInQty): r.StockInQty = value; break;
                case nameof(Requisition.StockInDate): r.StockInDate = ParseDate(value); break;
                case nameof(Requisition.Remark): r.Remark = value; break;
            }
        }

        private string SetPlanFieldValue(Plan p, string field, string value)
        {
            if (value == "")
            {
                value = null;
            }
            switch (field)
            {
                case nameof(Plan.JobNo): p.JobNo = value; break;
                case nameof(Plan.Product): p.Product = value; break;
                case nameof(Plan.Customer): p.Customer = value; break;
                case nameof(Plan.ModelName): p.ModelName = value; break;
                case nameof(Plan.Stage): p.Stage = value; break;
                case nameof(Plan.TestItem): p.TestItem = value; break;
                case nameof(Plan.SampleSize): p.SampleSize = value; break;
                case nameof(Plan.TestPeriod): p.TestPeriod = value; break;
                case nameof(Plan.Owner): p.Owner = value; break;
                case nameof(Plan.StartDate): p.StartDate = ParseDate(value); break;
                case nameof(Plan.EndDate): p.EndDate = ParseDate(value); break;
                case nameof(Plan.Status): p.Status = value; break;
                case nameof(Plan.Remark): p.Remark = value; break;
            }
            return _vm.ValidateField(field switch
            {
                nameof(Plan.JobNo) => "JobNo",
                nameof(Plan.Status) => "Status",
                nameof(Plan.TestItem) => "TestItem",
                nameof(Plan.Product) => "Product",
                nameof(Plan.Customer) => "Customer",
                nameof(Plan.Stage) => "Stage",
                _ => ""
            }, value);
        }

        private static DateTime? ParseDate(string text)
        {
            if (text == null)
            {
                return null;
            }
            return DateTime.TryParseExact(text, ["yyyy/M/d", "yyyy/M/d H:mm:ss", "yyyy-M-d"],
                CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt) ? dt : null;
        }

        /* ###############################  排序（列头点击已关闭，统一走菜单）  ################################ */

        /// <summary>领退表当前排序字段与方向（为空表示默认顺序）</summary>
        private string _reqSortField;
        private ListSortDirection _reqSortDirection = ListSortDirection.Ascending;

        /// <summary>计划表当前排序字段与方向（为空表示默认顺序）</summary>
        private string _planSortField;
        private ListSortDirection _planSortDirection = ListSortDirection.Ascending;

        /// <summary>窗口菜单「排序」：作用于当前 Tab 对应的表</summary>
        private void Menu_Sort_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item)
            {
                BuildSortMenu(item, ActiveGrid());
            }
        }

        /// <summary>右键菜单「排序」：按 Tag 指明是领退表(req)还是计划表(plan)</summary>
        private void Menu_ContextSort_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item)
            {
                BuildSortMenu(item, (item.Tag as string) == "plan" ? dg_plans : dg_requisitions);
            }
        }

        /// <summary>窗口菜单「视图」：显示/隐藏列（作用于当前 Tab 对应的表）</summary>
        private void Menu_View_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item)
            {
                BuildColumnMenu(item, ActiveGrid());
            }
        }

        /// <summary>当前 Tab 对应的表格（0=领退表，1=计划表）</summary>
        private DataGrid ActiveGrid() => tabs.SelectedIndex == 1 ? dg_plans : dg_requisitions;

        /// <summary>
        /// 构建排序菜单（类似资源管理器的"排序方式"）：默认顺序 + 各可排序字段 + 升序/降序。
        /// 字段来自列的 SortMemberPath，菜单文字直接用已本地化的列头。
        /// </summary>
        private void BuildSortMenu(MenuItem parent, DataGrid grid)
        {
            bool isPlan = ReferenceEquals(grid, dg_plans);
            string current = isPlan ? _planSortField : _reqSortField;
            ListSortDirection direction = isPlan ? _planSortDirection : _reqSortDirection;

            parent.Items.Clear();
            MenuItem defaultItem = new()
            {
                Header = LanguageService.Get("Sort_Default"),
                IsCheckable = true,
                IsChecked = string.IsNullOrEmpty(current)
            };
            defaultItem.Click += (s, e) => ApplySort(grid, null, ListSortDirection.Descending);
            parent.Items.Add(defaultItem);
            parent.Items.Add(new Separator());

            foreach (DataGridColumn column in grid.Columns)
            {
                string property = column.SortMemberPath;
                if (string.IsNullOrWhiteSpace(property))
                {
                    continue;
                }
                bool isCurrent = property == current;
                MenuItem item = new()
                {
                    Header = column.Header?.ToString(),
                    IsCheckable = true,
                    IsChecked = isCurrent
                };
                string captured = property;
                // 再点当前字段：升序/降序互换；点其他字段：升序
                ListSortDirection next = isCurrent && direction == ListSortDirection.Ascending
                    ? ListSortDirection.Descending
                    : ListSortDirection.Ascending;
                item.Click += (s, e) => ApplySort(grid, captured, next);
                parent.Items.Add(item);
            }

            parent.Items.Add(new Separator());
            MenuItem ascending = new()
            {
                Header = LanguageService.Get("Common_Ascending"),
                IsCheckable = true,
                IsChecked = direction == ListSortDirection.Ascending,
                IsEnabled = !string.IsNullOrEmpty(current)
            };
            ascending.Click += (s, e) => ApplySort(grid, current, ListSortDirection.Ascending);
            parent.Items.Add(ascending);

            MenuItem descending = new()
            {
                Header = LanguageService.Get("Common_Descending"),
                IsCheckable = true,
                IsChecked = direction == ListSortDirection.Descending,
                IsEnabled = !string.IsNullOrEmpty(current)
            };
            descending.Click += (s, e) => ApplySort(grid, current, ListSortDirection.Descending);
            parent.Items.Add(descending);
        }

        /// <summary>
        /// 应用排序并记录当前字段/方向（供菜单勾选状态使用）
        /// </summary>
        private void ApplySort(DataGrid grid, string property, ListSortDirection direction)
        {
            if (ReferenceEquals(grid, dg_plans))
            {
                _planSortField = property;
                _planSortDirection = direction;
                _vm.SortPlans(property, direction);
            }
            else
            {
                _reqSortField = property;
                _reqSortDirection = direction;
                _vm.SortRequisitions(property, direction);
            }
        }

        /* ###############################  显示/隐藏列  ################################ */

        private void Menu_ReqColumns_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            BuildColumnMenu(menu_req_columns, dg_requisitions);
        }

        private void Menu_PlanColumns_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            BuildColumnMenu(menu_plan_columns, dg_plans);
        }

        private static void BuildColumnMenu(MenuItem parent, DataGrid grid)
        {
            parent.Items.Clear();
            foreach (DataGridColumn column in grid.Columns)
            {
                MenuItem item = new()
                {
                    Header = column.Header?.ToString(),
                    IsCheckable = true,
                    IsChecked = column.Visibility == Visibility.Visible
                };
                DataGridColumn captured = column;
                item.Click += (s, e) =>
                    captured.Visibility = item.IsChecked ? Visibility.Visible : Visibility.Collapsed;
                parent.Items.Add(item);
            }
        }

        /* ###############################  回线转移单  ################################ */

        private void Btn_ReturnLine_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                WindowReturnLine window = new() { Topmost = true };
                window.Show();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "打开回线转移单失败");
                _ = MessageBox.Show($"打开回线转移单失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        /* ###############################  列顺序与可见性持久化  ################################ */

        private void SaveColumnState()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LayoutFile));
                Dictionary<string, List<string>> state = new()
                {
                    ["requisitions"] = dg_requisitions.Columns.Select(ColumnKey).ToList(),
                    ["plans"] = dg_plans.Columns.Select(ColumnKey).ToList()
                };
                File.WriteAllText(LayoutFile, Newtonsoft.Json.JsonConvert.SerializeObject(state));
            }
            catch (Exception ex)
            {
                _logger.Warn($"保存列布局失败: {ex.Message}");
            }
        }

        private static string ColumnKey(DataGridColumn column)
        {
            string order = column.DisplayIndex.ToString("D3");
            string visible = column.Visibility == Visibility.Visible ? "V" : "H";
            string name = column.Header?.ToString() ?? "?";
            return $"{order}|{visible}|{name}";
        }

        private void RestoreColumnState()
        {
            try
            {
                if (!File.Exists(LayoutFile))
                {
                    return;
                }
                Dictionary<string, List<string>> state = Newtonsoft.Json.JsonConvert
                    .DeserializeObject<Dictionary<string, List<string>>>(File.ReadAllText(LayoutFile));
                RestoreColumns(state.TryGetValue("requisitions", out List<string> reqKeys) ? reqKeys : null, dg_requisitions);
                RestoreColumns(state.TryGetValue("plans", out List<string> planKeys) ? planKeys : null, dg_plans);
            }
            catch (Exception ex)
            {
                _logger.Warn($"恢复列布局失败: {ex.Message}");
            }
        }

        private static void RestoreColumns(List<string> keys, DataGrid grid)
        {
            if (keys == null)
            {
                return;
            }
            foreach (string key in keys)
            {
                string[] parts = key.Split('|');
                if (parts.Length < 3)
                {
                    continue;
                }
                DataGridColumn column = grid.Columns.FirstOrDefault(c => (c.Header?.ToString() ?? "?") == parts[2]);
                if (column == null)
                {
                    continue;
                }
                column.Visibility = parts[1] == "V" ? Visibility.Visible : Visibility.Collapsed;
                if (int.TryParse(parts[0], out int displayIndex))
                {
                    column.DisplayIndex = Math.Min(displayIndex, grid.Columns.Count - 1);
                }
            }
        }
    }
}
