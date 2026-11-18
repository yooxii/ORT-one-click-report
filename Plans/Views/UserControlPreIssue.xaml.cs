using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
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
            Loaded += (s, e) => Attach();
            DataContextChanged += (s, e) => Attach();
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
                    Width = new DataGridLength(64),
                    ElementStyle = (Style)Resources["PreText"]
                });
                dg_schedule.Columns.Add(new DataGridTextColumn
                {
                    Header = LanguageService.Get("PreIssue_Model"),
                    Binding = new Binding(nameof(PreScheduleRow.Model)),
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
                _vm.AddPendingModel(dialog.Model, dialog.JobNo);
            }
        }

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

        /* ###############################  右键：复制 / 黑名单 / 删除  ################################ */

        private void Menu_CopyCell_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_contextCellValue))
            {
                RightClickCopy.Copy(_contextCellValue);
            }
        }

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

        private void Menu_BoardAddBlacklist_Click(object sender, RoutedEventArgs e)
        {
            if (_vm != null && _contextRow is OrtBoardRow row && !string.IsNullOrWhiteSpace(row.Model))
            {
                _vm.AddBlacklistModel(row.Model, LanguageService.Get("PreIssue_BlacklistFromBoard"));
            }
        }

        private void Menu_ReqAddBlacklist_Click(object sender, RoutedEventArgs e)
        {
            if (_vm != null && _contextRow is PreIssueRequisition row && !string.IsNullOrWhiteSpace(row.PartNo))
            {
                _vm.AddBlacklistModel(row.PartNo, LanguageService.Get("PreIssue_BlacklistFromBoard"));
            }
        }

        private void Menu_DeletePending_Click(object sender, RoutedEventArgs e)
        {
            if (_vm != null && _contextRow is PreIssueRequisition row)
            {
                _vm.DeletePendingCommand.Execute(row);
            }
        }

        /* ###############################  右键定位  ################################ */

        /// <summary>三个表格共用的右键定位：记下单元格值、所在行，并选中该行</summary>
        private void DataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            _contextCellValue = null;
            _contextRow = null;
            if (sender is not DataGrid grid || e.OriginalSource is not DependencyObject source)
            {
                return;
            }
            DataGridCell cell = FindAncestor<DataGridCell>(source);
            if (cell == null)
            {
                return;
            }
            cell.IsSelected = true;
            if (cell.Content is TextBlock text)
            {
                _contextCellValue = text.Text;
            }
            _contextRow = cell.DataContext;
            if (_contextRow != null)
            {
                grid.SelectedItem = _contextRow;
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
