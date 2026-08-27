using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

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

        private static readonly string LayoutFile
            = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "plans_layout.json");

        public WindowPlans()
        {
            InitializeComponent();
            _vm = App.ServiceProvider.GetRequiredService<PlansViewModel>();
            DataContext = _vm;
            Loaded += (s, e) =>
            {
                RestoreColumnState();
                _vm.Refresh();
            };
            Closing += (s, e) => SaveColumnState();
        }

        /* ###############################  行号  ################################ */

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
                    _ = MessageBox.Show(error, "格式校验失败");
                    e.Cancel = true;
                    return;
                }
            }
            if (column == "Status")
            {
                string error = _vm.ValidateField("Status", plan.Status);
                if (error != null)
                {
                    _ = MessageBox.Show(error, "校验失败");
                    e.Cancel = true;
                    return;
                }
            }
            if (column == "TestItem")
            {
                string error = _vm.ValidateField("TestItem", plan.TestItem);
                if (error != null)
                {
                    _ = MessageBox.Show(error, "校验失败");
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
            _vm.NotifyPendingChanged();
            _vm.StatusMessage = _vm.PendingText;
        }

        /* ###############################  右键单元格定位  ################################ */

        private void Dg_Requisitions_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            SelectCellUnderMouse(sender as DataGrid, e);
        }

        private void Dg_Plans_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            SelectCellUnderMouse(sender as DataGrid, e);
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
                        _ = MessageBox.Show($"{columns[colIndex].Header}: {error}", "粘贴校验失败");
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
                _ = MessageBox.Show($"打开回线转移单失败:\n{ex.Message}", "错误");
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
