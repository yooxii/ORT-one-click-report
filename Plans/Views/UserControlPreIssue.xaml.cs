using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// 预领标签页：内层三个页签（月排程 / ORT看板 / 准备领用）。
    /// 月排程的日列随「周视图 / 整月视图」动态生成（前两列線別、機種冻结），
    /// 其余两个页签是导入进来的数据，只读展示 + 右键复制 / 加黑名单 / 删除待开单。
    /// </summary>
    public partial class UserControlPreIssue : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private PreIssueViewModel _vm;

        /// <summary>右键按下时所在单元格的值（上下文菜单里的「复制该值」用）</summary>
        private string _contextCellValue;

        /// <summary>右键按下时所在的行（上下文菜单里的「加入黑名单 / 删除待开单」用）</summary>
        private object _contextRow;

        public UserControlPreIssue()
        {
            InitializeComponent();
            // 准备领用：列头的右键菜单（筛选 / 排序）由本类自己弹，别让表格的右键菜单跟着一起冒出来
            dg_req.ContextMenuOpening += (s, e) =>
            {
                if (e.OriginalSource is DependencyObject source && FindAncestor<DataGridColumnHeader>(source) != null)
                {
                    e.Handled = true;
                }
            };
            Loaded += (s, e) => Attach();
            DataContextChanged += (s, e) => Attach();
            // 月排程：列头右键是筛选菜单（由本类自己弹），别让表格的单元格右键菜单跟着一起冒出来
            dg_schedule.ContextMenuOpening += (s, e) =>
            {
                if (e.OriginalSource is DependencyObject source && FindAncestor<DataGridColumnHeader>(source) != null)
                {
                    e.Handled = true;
                }
            };
            // 浮动/回到窗口会把控件挪到另一个窗口：Unloaded 会先触发，这里只解除事件，不动数据
            Unloaded += (s, e) => Detach();
        }

        /// <summary>绑定视图模型并重建表格列（可重复调用）</summary>
        private void Attach()
        {
            PreIssueViewModel vm = DataContext as PreIssueViewModel;
            if (ReferenceEquals(vm, _vm))
            {
                if (vm != null && dg_schedule.Columns.Count <= 2)
                {
                    RebuildScheduleColumns();
                }
                return;
            }
            Detach();
            _vm = vm;
            if (_vm == null)
            {
                return;
            }
            _vm.ScheduleColumnsChanged += RebuildScheduleColumns;
            RebuildScheduleColumns();
        }

        private void Detach()
        {
            if (_vm != null)
            {
                _vm.ScheduleColumnsChanged -= RebuildScheduleColumns;
            }
        }

        /* ###############################  月排程的列  ################################ */

        /// <summary>
        /// 重建月排程的列：線別、機種（冻结两列） + 当前显示的日列 + 合计。
        /// 周视图只出 7 列，整月视图出当月天数，所以平时不用左右拖滚动条。
        /// </summary>
        private void RebuildScheduleColumns()
        {
            try
            {
                dg_schedule.Columns.Clear();
                dg_schedule.FrozenColumnCount = 0;
                dg_schedule.Columns.Add(new DataGridTextColumn
                {
                    Header = LanguageService.Get("PreIssue_Line"),
                    Binding = new Binding(nameof(PreScheduleRow.Line)),
                    SortMemberPath = nameof(PreScheduleRow.Line),
                    Width = new DataGridLength(64),
                    ElementStyle = (Style)Resources["PreText"]
                });
                dg_schedule.Columns.Add(new DataGridTextColumn
                {
                    Header = LanguageService.Get("PreIssue_Model"),
                    Binding = new Binding(nameof(PreScheduleRow.Model)),
                    SortMemberPath = nameof(PreScheduleRow.Model),
                    Width = new DataGridLength(140),
                    ElementStyle = (Style)Resources["PreText"]
                });
                // 線別 + 機種 两列冻结：左右翻日期时这两列始终可见（列都加完再设冻结数）
                dg_schedule.FrozenColumnCount = dg_schedule.Columns.Count;
                if (_vm != null)
                {
                    foreach (ScheduleDayColumn day in _vm.ScheduleColumns)
                    {
                        int index = day.InMonth ? day.Day : 0;
                        DataGridTextColumn column = new()
                        {
                            Header = day.Header,
                            Binding = new Binding($"{nameof(PreScheduleRow.Days)}[{index}]"),
                            Width = new DataGridLength(56),
                            ElementStyle = (Style)Resources["PreText"]
                        };
                        if (day.IsWeekend)
                        {
                            column.CellStyle = BuildDayCellStyle(true);
                        }
                        dg_schedule.Columns.Add(column);
                    }
                }
                DataGridTextColumn total = new()
                {
                    Header = LanguageService.Get("PreIssue_Total"),
                    Binding = new Binding(nameof(PreScheduleRow.Total)),
                    Width = new DataGridLength(80),
                    ElementStyle = (Style)Resources["PreText"],
                    FontWeight = FontWeights.Bold
                };
                dg_schedule.Columns.Add(total);
            }
            catch (Exception ex)
            {
                _logger.Warn($"重建月排程列失败: {ex.Message}");
            }
        }

        /// <summary>周末列底色（沿用单元格样式，只加底色）</summary>
        private Style BuildDayCellStyle(bool weekend)
        {
            Style basedOn = (Style)Resources["PreCell"];
            Style style = new(typeof(DataGridCell), basedOn);
            if (weekend)
            {
                style.Setters.Add(new Setter(BackgroundProperty, FindResource("StatusWarnBgBrush")));
            }
            return style;
        }

        /* ###############################  按钮  ################################ */

        private void Btn_AddPending_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null)
            {
                return;
            }
            WindowPreIssuePending dialog = new(_vm.PendingModelCandidates())
            {
                Owner = Window.GetWindow(this)
            };
            if (dialog.ShowDialog() == true)
            {
                _vm.AddPendingModel(dialog.Model, dialog.JobNo, dialog.Qty, dialog.Note);
            }
        }

        /// <summary>月排程右键「添加到待开单」：机种与本月合计数量预填进添加窗口</summary>
        private void Menu_ScheduleAddPending_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_AddPendingTitle"),
            () =>
            {
                if (_vm == null || _contextRow is not PreScheduleRow row || string.IsNullOrWhiteSpace(row.Model))
                {
                    return;
                }
                WindowPreIssuePending dialog = new(_vm.PendingModelCandidates(), row.Model, null,
                    row.Total > 0 ? row.Total.ToString(System.Globalization.CultureInfo.InvariantCulture) : null)
                {
                    Owner = Window.GetWindow(this)
                };
                if (dialog.ShowDialog() == true)
                {
                    _vm.AddPendingModel(dialog.Model, dialog.JobNo, dialog.Qty, dialog.Note);
                }
            });

        /// <summary>ORT看板右键「添加到待开单」：机种 / 工令 / 生產數量预填进添加窗口</summary>
        private void Menu_BoardAddPending_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_AddPendingTitle"),
            () =>
            {
                if (_vm == null || _contextRow is not OrtBoardRow row || string.IsNullOrWhiteSpace(row.Model))
                {
                    return;
                }
                WindowPreIssuePending dialog = new(_vm.PendingModelCandidates(), row.Model, row.JobNo, row.Qty)
                {
                    Owner = Window.GetWindow(this)
                };
                if (dialog.ShowDialog() == true)
                {
                    _vm.AddPendingModel(dialog.Model, dialog.JobNo, dialog.Qty, dialog.Note);
                }
            });

        private void Btn_Blacklist_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null)
            {
                return;
            }
            WindowOrtBlacklist dialog = new(_vm)
            {
                Owner = Window.GetWindow(this)
            };
            dialog.ShowDialog();
            _vm.ReloadAfterBlacklistChange();
        }

        private void Btn_Changes_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null)
            {
                return;
            }
            WindowPreIssueChanges dialog = new(_vm)
            {
                Owner = Window.GetWindow(this)
            };
            dialog.ShowDialog();
        }

        private void Btn_ReminderSettings_Click(object sender, RoutedEventArgs e)
        {
            WindowPreIssueReminder dialog = new()
            {
                Owner = Window.GetWindow(this)
            };
            dialog.ShowDialog();
        }

        /// <summary>同系列客户优先级：编辑「前 6 位相同则只开一单」时优先挑哪个客户代码</summary>
        private void Btn_CustomerPriority_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null)
            {
                return;
            }
            WindowPreIssueCustomerPriority dialog = new(_vm)
            {
                Owner = Window.GetWindow(this)
            };
            dialog.ShowDialog();
        }

        /* ###############################  右键：复制 / 黑名单 / 删除  ################################ */

        private void Menu_CopyCell_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_Menu_CopyCell"),
            () =>
            {
                if (!string.IsNullOrWhiteSpace(_contextCellValue))
                {
                    RightClickCopy.Copy(_contextCellValue);
                }
            });

        private void Menu_CopyScheduleRow_Click(object sender, RoutedEventArgs e)
        {
            if (_contextRow is PreScheduleRow row && _vm != null)
            {
                StringBuilder text = new();
                text.Append(row.Line).Append('\t').Append(row.Model);
                foreach (ScheduleDayColumn day in _vm.ScheduleColumns)
                {
                    text.Append('\t').Append(day.InMonth ? row.Days[day.Day] : "");
                }
                text.Append('\t').Append(row.Total);
                RightClickCopy.Copy(text.ToString());
            }
        }

        private void Menu_CopyBoardRow_Click(object sender, RoutedEventArgs e)
        {
            if (_contextRow is OrtBoardRow row)
            {
                string text = string.Join("\t", new[]
                {
                    row.ReportDate, row.Category, row.Model, row.Line, row.ProduceDate, row.JobNo,
                    row.Qty, row.Hours, row.Reported65, row.Stocked87, row.ReceiveDate, row.PartNo,
                    row.MaterialStock, row.TestResult, row.Forbidden94
                }.Select(v => v ?? ""));
                RightClickCopy.Copy(text);
            }
        }

        private void Menu_CopyReqRow_Click(object sender, RoutedEventArgs e)
        {
            if (_contextRow is PreIssueRequisition row)
            {
                string text = string.Join("\t", new[]
                {
                    row.Vchno ?? "", row.WareCode ?? "", row.TrCode ?? "", row.PartNo ?? "", row.ReqQty ?? "",
                    row.IverId ?? "", row.VchStatus ?? "", row.IssueDept ?? "", row.CloseCode ?? "", row.Pc ?? ""
                });
                RightClickCopy.Copy(text);
            }
        }

        private void Menu_BoardAddBlacklist_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_AddBlacklistTitle"),
            () => OpenBlacklistEditor(_contextRow is OrtBoardRow board ? board.Model : null));

        private void Menu_ReqAddBlacklist_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_AddBlacklistTitle"),
            () => OpenBlacklistEditor(_contextRow is PreIssueRequisition req ? req.PartNo : null));

        /// <summary>
        /// 加入黑名单统一走「黑名单条目编辑框」：先把机种填进去，用户可以再改/换成同系列正则，
        /// 点「完成」才真正写入。
        /// </summary>
        private void OpenBlacklistEditor(string model)
        {
            if (_vm == null || string.IsNullOrWhiteSpace(model))
            {
                return;
            }
            WindowOrtBlacklistEdit dialog = new(_vm, model)
            {
                Owner = Window.GetWindow(this)
            };
            if (dialog.ShowDialog() == true && _vm.AddBlacklistModel(dialog.Value, dialog.Reason, dialog.IsPattern))
            {
                _vm.ReloadAfterBlacklistChange();
            }
        }

        private void Menu_DeletePending_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_DeletePendingTitle"),
            () =>
            {
                if (_vm == null || _contextRow is not PreIssueRequisition row)
                {
                    return;
                }
                if (!_vm.CanDeletePending(row))
                {
                    _ = MessageBox.Show(LanguageService.Get("PreIssue_OnlyPendingDeletable"),
                        LanguageService.Get("PreIssue_DeletePendingTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                // 三按钮确认：是（只删）/ 是并加入黑名单（删 + 加黑名单）/ 否
                WindowPreIssueDeletePending dialog = new(row.PartNo)
                {
                    Owner = Window.GetWindow(this)
                };
                if (dialog.ShowDialog() != true)
                {
                    return;
                }
                _vm.DeletePending(row, dialog.AddBlacklist);
            });

        /* ###############################  借用表：右键借用 / 归还  ################################ */

        /// <summary>右键「登记借用」：选中的（或右键那一行）改成借用中并填借用日期</summary>
        private void Menu_Borrow_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_Borrow_Menu_Borrow"),
            () =>
            {
                foreach (PreIssueBorrowItem row in BorrowTargets())
                {
                    if (!_vm.MarkBorrowedRow(row))
                    {
                        break;
                    }
                }
            });

        /// <summary>右键「归还」：选中的（或右键那一行）改成已归还并填归还日期</summary>
        private void Menu_BorrowReturn_Click(object sender, RoutedEventArgs e) => SafeRun(
            LanguageService.Get("PreIssue_Borrow_Menu_Return"),
            () =>
            {
                foreach (PreIssueBorrowItem row in BorrowTargets())
                {
                    if (!_vm.MarkReturnedRow(row))
                    {
                        break;
                    }
                }
            });

        /// <summary>右键作用范围：右键落在多选范围内就整批处理，否则只处理右键那一行</summary>
        private List<PreIssueBorrowItem> BorrowTargets()
        {
            List<PreIssueBorrowItem> targets = [];
            if (_vm == null)
            {
                return targets;
            }
            if (_contextRow is PreIssueBorrowItem hit && dg_borrow.SelectedItems.Contains(hit))
            {
                foreach (object item in dg_borrow.SelectedItems)
                {
                    if (item is PreIssueBorrowItem row)
                    {
                        targets.Add(row);
                    }
                }
            }
            if (targets.Count == 0 && _contextRow is PreIssueBorrowItem single)
            {
                targets.Add(single);
            }
            return targets;
        }

        private void Menu_CopyBorrowRow_Click(object sender, RoutedEventArgs e)
        {
            if (_contextRow is PreIssueBorrowItem row)
            {
                string text = string.Join("\t", new[]
                {
                    row.Status ?? "", row.Model ?? "", row.Qty ?? "",
                    row.BorrowDate?.ToString("yyyy/M/d") ?? "", row.ReturnDate?.ToString("yyyy/M/d") ?? ""
                });
                RightClickCopy.Copy(text);
            }
        }

        /* ###############################  右键定位  ################################ */

        /// <summary>
        /// 三个表格共用的右键定位：记下单元格值与所在行，并把该行定位选中。
        /// 注意：这几张表都是「整行选中」（SelectionUnit=FullRow），这种模式下改单元格的
        /// IsSelected 会抛 InvalidOperationException（“SelectionUnit 为 FullRow 时，无法更改单元格选择”），
        /// 所以只动行选中；已经多选的行保持原样（右键操作仍按原来的选择范围走）。
        /// 右键只是定位/复制，任何意外都不该把程序带崩，整段加兜底。
        /// </summary>
        private void DataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                _contextCellValue = null;
                _contextRow = null;
                if (sender is not DataGrid grid || e.OriginalSource is not DependencyObject source)
                {
                    return;
                }
                // 准备领用：右键列头 = 筛选 / 排序 菜单（不弹单元格菜单）
                if (ReferenceEquals(grid, dg_req) && FindAncestor<DataGridColumnHeader>(source) is DataGridColumnHeader header)
                {
                    ShowReqHeaderMenu(header);
                    e.Handled = true;
                    return;
                }
                // 月排程：右键线别 / 机种列头 = 筛选菜单
                if (ReferenceEquals(grid, dg_schedule) && FindAncestor<DataGridColumnHeader>(source) is DataGridColumnHeader scheduleHeader)
                {
                    ShowScheduleHeaderMenu(scheduleHeader);
                    e.Handled = true;
                    return;
                }
                DataGridCell cell = FindAncestor<DataGridCell>(source);
                if (cell == null)
                {
                    return;
                }
                if (cell.Content is TextBlock text)
                {
                    _contextCellValue = text.Text;
                }
                object row = cell.DataContext;
                _contextRow = row;
                if (row == null || grid.SelectionUnit == DataGridSelectionUnit.FullRow)
                {
                    if (row != null && !grid.SelectedItems.Contains(row))
                    {
                        grid.SelectedItem = row;
                    }
                    return;
                }
                if (!grid.SelectedItems.Contains(row))
                {
                    cell.IsSelected = true;
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"右键定位失败（已忽略，不影响继续操作）: {ex.Message}");
            }
        }

        /* ###############################  准备领用：双击编辑  ################################ */

        /// <summary>没有编辑权限时拦住进入编辑；有权限才允许改</summary>
        private void Dg_Req_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (_vm == null || _vm.CanEditRequisitionData)
            {
                return;
            }
            e.Cancel = true;
            _ = MessageBox.Show(LanguageService.Get("PreIssue_ReqEditDenied"),
                LanguageService.Get("PreIssue_RequisitionTab"), MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        /// <summary>
        /// 单元格编辑结束（提交）后保存这一行：绑定是即时更新（UpdateSourceTrigger=PropertyChanged），
        /// 这里延后到下一轮消息循环再写库，避免打断 DataGrid 自己的提交事务。
        /// </summary>
        private void Dg_Req_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not PreIssueRequisition row || _vm == null)
            {
                return;
            }
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    _vm.SaveReqRow(row);
                }
                catch (Exception ex)
                {
                    _logger.Warn($"保存准备领用手工修改失败: {ex.Message}");
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /* ###############################  借用表：双击编辑  ################################ */

        /// <summary>没有编辑权限时拦住进入编辑</summary>
        private void Dg_Borrow_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (_vm == null || _vm.CanEditBorrowData)
            {
                return;
            }
            e.Cancel = true;
            _ = MessageBox.Show(LanguageService.Get("PreIssue_ReqEditDenied"),
                LanguageService.Get("PreIssue_BorrowTab"), MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        /// <summary>借用表改完一格后保存这一行（同样延后到下一轮消息循环，避免打断提交事务）</summary>
        private void Dg_Borrow_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not PreIssueBorrowItem row || _vm == null)
            {
                return;
            }
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    _vm.SaveBorrowRow(row);
                }
                catch (Exception ex)
                {
                    _logger.Warn($"保存借用表修改失败: {ex.Message}");
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /* ###############################  准备领用：列头筛选 / 排序  ################################ */

        /// <summary>列对应的属性名（优先 SortMemberPath，其次绑定的路径）</summary>
        private static string PropertyOf(DataGridColumn column)
        {
            if (column == null)
            {
                return null;
            }
            if (!string.IsNullOrWhiteSpace(column.SortMemberPath))
            {
                return column.SortMemberPath;
            }
            if (column is DataGridBoundColumn bound && bound.Binding is Binding binding)
            {
                return binding.Path?.Path;
            }
            return null;
        }

        /// <summary>
        /// 准备领用的列头右键菜单：排序（正序 / 倒序 / 不排序）+ 筛选（按该列的值勾选）。
        /// </summary>
        private void ShowReqHeaderMenu(DataGridColumnHeader header)
        {
            try
            {
                string property = PropertyOf(header?.Column);
                if (_vm == null || string.IsNullOrEmpty(property))
                {
                    return;
                }
                ContextMenu menu = new()
                {
                    PlacementTarget = header,
                    Placement = PlacementMode.Bottom
                };

                // ---- 排序 ----
                MenuItem sort = new() { Header = LanguageService.Get("PreIssue_Menu_Sort") };
                (string Property, ListSortDirection Direction)? current = _vm.CurrentReqSort;
                bool sameColumn = current.HasValue && string.Equals(current.Value.Property, property, StringComparison.Ordinal);
                sort.Items.Add(BuildSortItem(LanguageService.Get("PreIssue_Menu_SortAsc"), property,
                    ListSortDirection.Ascending, sameColumn && current.Value.Direction == ListSortDirection.Ascending));
                sort.Items.Add(BuildSortItem(LanguageService.Get("PreIssue_Menu_SortDesc"), property,
                    ListSortDirection.Descending, sameColumn && current.Value.Direction == ListSortDirection.Descending));
                sort.Items.Add(BuildSortItem(LanguageService.Get("PreIssue_Menu_SortNone"), property,
                    null, !sameColumn));
                menu.Items.Add(sort);

                // ---- 筛选（惰性：勾选只改本地状态，点「确定」才应用） ----
                MenuItem filter = new() { Header = LanguageService.Get("PreIssue_Menu_Filter") };
                List<string> values = _vm.GetReqColumnValues(property);
                HashSet<string> selection = _vm.GetReqColumnSelection(property);
                if (values.Count == 0)
                {
                    menu.Items.Add(new MenuItem { Header = LanguageService.Get("PreIssue_Menu_FilterEmpty"), IsEnabled = false });
                }
                else if (values.Count > 80)
                {
                    menu.Items.Add(new MenuItem
                    {
                        Header = string.Format(LanguageService.Get("PreIssue_Menu_FilterTooMany"), values.Count),
                        IsEnabled = false
                    });
                }
                else
                {
                    // 编辑中的勾选集合：不动表格，等「确定」再一次性写回
                    HashSet<string> pending = selection == null
                        ? new HashSet<string>(values, StringComparer.Ordinal)
                        : new HashSet<string>(selection, StringComparer.Ordinal);
                    List<(string Value, MenuItem Item)> items = [];
                    // 先建出「确定」项，勾选变化时好同步它标题上的已选数量
                    MenuItem applyFilter = new()
                    {
                        Header = LanguageService.Get("Common_OK"),
                        FontWeight = FontWeights.Bold,
                        ToolTip = LanguageService.Get("Filter_LazyHint")
                    };

                    void RefreshChecks()
                    {
                        foreach ((string value, MenuItem item) in items)
                        {
                            item.IsChecked = pending.Contains(value);
                        }
                        SyncApplyText();
                    }

                    filter.Items.Add(BuildBulkFilterItem(LanguageService.Get("Filter_SelectAll"), () =>
                    {
                        pending.Clear();
                        pending.UnionWith(values);
                        RefreshChecks();
                    }));
                    filter.Items.Add(BuildBulkFilterItem(LanguageService.Get("Filter_SelectNone"), () =>
                    {
                        pending.Clear();
                        RefreshChecks();
                    }));
                    filter.Items.Add(BuildBulkFilterItem(LanguageService.Get("Filter_Invert"), () =>
                    {
                        List<string> inverted = values.Where(v => !pending.Contains(v)).ToList();
                        pending.Clear();
                        pending.UnionWith(inverted);
                        RefreshChecks();
                    }));
                    filter.Items.Add(new Separator());

                    foreach (string value in values)
                    {
                        MenuItem item = new()
                        {
                            Header = string.IsNullOrEmpty(value) ? LanguageService.Get("PreIssue_Menu_FilterBlank") : value,
                            IsCheckable = true,
                            IsChecked = pending.Contains(value),
                            StaysOpenOnClick = true
                        };
                        string captured = value;
                        item.Click += (s, e) =>
                        {
                            if (item.IsChecked)
                            {
                                pending.Add(captured);
                            }
                            else
                            {
                                pending.Remove(captured);
                            }
                            RefreshChecks();   // 同步「确定」上的已选数量
                        };
                        items.Add((value, item));
                        filter.Items.Add(item);
                    }

                    filter.Items.Add(new Separator());
                    applyFilter.Click += (s, e) =>
                    {
                        _vm.SetReqColumnFilter(property, pending);
                        menu.IsOpen = false;
                    };
                    filter.Items.Add(applyFilter);
                    MenuItem cancelFilter = new() { Header = LanguageService.Get("Common_Cancel") };
                    cancelFilter.Click += (s, e) => menu.IsOpen = false;
                    filter.Items.Add(cancelFilter);

                    // 「确定」上写清已选数量：一个都没勾时说明等于「显示全部」
                    void SyncApplyText()
                    {
                        applyFilter.Header = pending.Count == 0
                            ? string.Format(LanguageService.Get("Filter_ApplyNoneFormat"), values.Count)
                            : string.Format(LanguageService.Get("Filter_ApplyCountFormat"), pending.Count, values.Count);
                    }
                    SyncApplyText();

                    menu.Items.Add(filter);
                }
                menu.Items.Add(new Separator());
                menu.Items.Add(BuildMenuItem(LanguageService.Get("PreIssue_Menu_FilterClear"), () =>
                {
                    _vm.ClearReqColumnFilter(property);
                }));
                menu.Items.Add(BuildMenuItem(LanguageService.Get("PreIssue_Menu_FilterClearAll"), _vm.ClearAllReqColumnFilters));

                menu.IsOpen = true;
            }
            catch (Exception ex)
            {
                _logger.Warn($"弹出准备领用列头菜单失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 月排程的列头右键菜单：線別 / 機種 两列按值筛选（与准备领用一样是惰性的，点「确定」才应用）。
        /// </summary>
        private void ShowScheduleHeaderMenu(DataGridColumnHeader header)
        {
            try
            {
                string property = PropertyOf(header?.Column);
                if (_vm == null
                    || (property != nameof(PreScheduleRow.Line) && property != nameof(PreScheduleRow.Model)))
                {
                    return;
                }
                ContextMenu menu = new()
                {
                    PlacementTarget = header,
                    Placement = PlacementMode.Bottom
                };
                MenuItem filter = new() { Header = LanguageService.Get("PreIssue_Menu_Filter") };
                List<string> values = _vm.GetScheduleColumnValues(property);
                HashSet<string> selection = _vm.GetScheduleColumnSelection(property);
                if (values.Count == 0)
                {
                    menu.Items.Add(new MenuItem { Header = LanguageService.Get("PreIssue_Menu_FilterEmpty"), IsEnabled = false });
                }
                else if (values.Count > 80)
                {
                    menu.Items.Add(new MenuItem
                    {
                        Header = string.Format(LanguageService.Get("PreIssue_Menu_FilterTooMany"), values.Count),
                        IsEnabled = false
                    });
                }
                else
                {
                    HashSet<string> pending = selection == null
                        ? new HashSet<string>(values, StringComparer.Ordinal)
                        : new HashSet<string>(selection, StringComparer.Ordinal);
                    List<(string Value, MenuItem Item)> items = [];
                    MenuItem applyFilter = new()
                    {
                        Header = LanguageService.Get("Common_OK"),
                        FontWeight = FontWeights.Bold,
                        ToolTip = LanguageService.Get("Filter_LazyHint")
                    };

                    void SyncApplyText()
                    {
                        applyFilter.Header = pending.Count == 0
                            ? string.Format(LanguageService.Get("Filter_ApplyNoneFormat"), values.Count)
                            : string.Format(LanguageService.Get("Filter_ApplyCountFormat"), pending.Count, values.Count);
                    }

                    void RefreshChecks()
                    {
                        foreach ((string value, MenuItem item) in items)
                        {
                            item.IsChecked = pending.Contains(value);
                        }
                        SyncApplyText();
                    }

                    filter.Items.Add(BuildBulkFilterItem(LanguageService.Get("Filter_SelectAll"), () =>
                    {
                        pending.Clear();
                        pending.UnionWith(values);
                        RefreshChecks();
                    }));
                    filter.Items.Add(BuildBulkFilterItem(LanguageService.Get("Filter_SelectNone"), () =>
                    {
                        pending.Clear();
                        RefreshChecks();
                    }));
                    filter.Items.Add(BuildBulkFilterItem(LanguageService.Get("Filter_Invert"), () =>
                    {
                        List<string> inverted = values.Where(v => !pending.Contains(v)).ToList();
                        pending.Clear();
                        pending.UnionWith(inverted);
                        RefreshChecks();
                    }));
                    filter.Items.Add(new Separator());

                    foreach (string value in values)
                    {
                        MenuItem item = new()
                        {
                            Header = string.IsNullOrEmpty(value) ? LanguageService.Get("PreIssue_Menu_FilterBlank") : value,
                            IsCheckable = true,
                            IsChecked = pending.Contains(value),
                            StaysOpenOnClick = true
                        };
                        string captured = value;
                        item.Click += (s, e) =>
                        {
                            if (item.IsChecked)
                            {
                                pending.Add(captured);
                            }
                            else
                            {
                                pending.Remove(captured);
                            }
                            RefreshChecks();
                        };
                        items.Add((value, item));
                        filter.Items.Add(item);
                    }

                    filter.Items.Add(new Separator());
                    applyFilter.Click += (s, e) =>
                    {
                        _vm.SetScheduleColumnFilter(property, pending);
                        menu.IsOpen = false;
                    };
                    filter.Items.Add(applyFilter);
                    MenuItem cancelFilter = new() { Header = LanguageService.Get("Common_Cancel") };
                    cancelFilter.Click += (s, e) => menu.IsOpen = false;
                    filter.Items.Add(cancelFilter);
                    SyncApplyText();
                    menu.Items.Add(filter);
                }
                menu.Items.Add(new Separator());
                menu.Items.Add(BuildMenuItem(LanguageService.Get("PreIssue_Menu_FilterClear"), () =>
                    _vm.ClearScheduleColumnFilter(property)));
                menu.Items.Add(BuildMenuItem(LanguageService.Get("PreIssue_Menu_FilterClearAll"), () =>
                    _vm.ClearScheduleColumnFilter(null)));
                menu.IsOpen = true;
            }
            catch (Exception ex)
            {
                _logger.Warn($"弹出月排程列头菜单失败: {ex.Message}");
            }
        }

        /// <summary>筛选里的批量勾选按钮（全选 / 全不选 / 反选）：只改本地勾选，不关菜单、不立即应用</summary>
        private static MenuItem BuildBulkFilterItem(string header, Action click)
        {
            MenuItem item = new() { Header = header, StaysOpenOnClick = true };
            item.Click += (s, e) => click();
            return item;
        }

        private MenuItem BuildSortItem(string header, string property, ListSortDirection? direction, bool isChecked)
        {
            MenuItem item = new()
            {
                Header = header,
                IsCheckable = true,
                IsChecked = isChecked
            };
            item.Click += (s, e) => _vm?.SortRequisitions(property, direction);
            return item;
        }

        private static MenuItem BuildMenuItem(string header, Action click)
        {
            MenuItem item = new() { Header = header };
            item.Click += (s, e) => click();
            return item;
        }
        /// <summary>右键菜单里的动作统一兜底：出错只提示，不把程序带崩</summary>
        private void SafeRun(string action, Action run)
        {
            try
            {
                run();
            }
            catch (Exception ex)
            {
                _logger.Warn($"{action}失败: {ex.Message}");
                _ = MessageBox.Show(ex.Message, action, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static T FindAncestor<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T match)
                {
                    return match;
                }
                node = node is Visual || node is System.Windows.Media.Media3D.Visual3D
                    ? VisualTreeHelper.GetParent(node)
                    : LogicalTreeHelper.GetParent(node);
            }
            return null;
        }
    }
}
