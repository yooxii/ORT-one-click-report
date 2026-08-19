using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowPlans.xaml 的交互逻辑
    /// </summary>
    public partial class WindowPlans : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public PlansViewModel PlansVM { get; }

        /// <summary>
        /// 列顺序布局文件（Data目录下），窗口关闭时保存，下次开启恢复
        /// </summary>
        private string LayoutFilePath => Path.Combine(
            App.ServiceProvider.GetService(typeof(DatabaseService)) is DatabaseService db ? db.DataDir : "", "plans_layout.json");

        /// <summary>
        /// 允许复制/粘贴的列（属性名，按表格列顺序）
        /// </summary>
        private static readonly string[] CopyableFields =
        [
            "JobNo", "RequisitionNo", "ModelName", "TestItem", "OutQty", "SN", "DC", "Rev",
            "WorkOrder", "ReturnRtOrder", "LineNo", "RequisitionDate", "ReturnDate",
            "StockInNo", "StockInDate", "Product", "Customer", "Stage",
            "SampleSize", "TestPeriod", "Owner", "StartDate", "EndDate", "Status", "Remark"
        ];

        public WindowPlans()
        {
            InitializeComponent();
            PlansVM = App.ServiceProvider.GetRequiredService<PlansViewModel>();
            DataContext = PlansVM;

            Loaded += (s, e) => RestoreColumnLayout();
            Closing += (s, e) => SaveColumnLayout();
        }

        /// <summary>
        /// 打开回线转移单工具（从一键报告迁移至此）
        /// </summary>
        private void Btn_ReturnLine_Click(object sender, RoutedEventArgs e)
        {
            WindowReturnLine windowReturnLine = new()
            {
                Owner = this
            };
            windowReturnLine.Show();
        }

        /* ###############################  Excel 风格：行号 / 编辑校验 / 联动  ################################ */

        /// <summary>
        /// 行头显示行号（像 Excel）
        /// </summary>
        private void Dg_Plans_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.GetIndex() + 1;
        }

        /// <summary>
        /// 单元格编辑结束：校验格式 → 机种联动 → 标记暂存修改
        /// </summary>
        private void Dg_Plans_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not Plan plan)
            {
                return;
            }
            string field = (e.Column as DataGridTextColumn)?.SortMemberPath
                ?? e.Column.SortMemberPath;
            string newValue = (e.EditingElement as TextBox)?.Text;

            // 格式校验
            string error = PlansVM.ValidateField(field, newValue);
            if (error != null)
            {
                e.Cancel = true;
                _ = MessageBox.Show(error, "输入校验失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 机种名称修改后自动带出产品别/客户别（还原计划表公式关系）；
            // Plan 已实现属性通知，单元格自动刷新，无需 Items.Refresh（编辑事务中 Refresh 会抛异常）
            if (field == "ModelName")
            {
                PlansVM.AutoFillByModel(plan);
            }
            PlansVM.MarkModified(plan);
        }

        /* ###############################  右键菜单：编辑 / 删除 / 复制 / 粘贴  ################################ */

        private void Menu_Edit_Click(object sender, RoutedEventArgs e)
        {
            if (PlansVM.EditCommand.CanExecute(null))
            {
                PlansVM.EditCommand.Execute(null);
            }
        }

        private void Menu_Delete_Click(object sender, RoutedEventArgs e)
        {
            if (dg_plans.CurrentItem is Plan plan)
            {
                PlansVM.DeleteRowCommand.Execute(plan);
            }
        }

        /// <summary>
        /// 复制选中行（TSV，多行以换行分隔）
        /// </summary>
        private void Menu_CopyRow_Click(object sender, RoutedEventArgs e)
        {
            List<Plan> rows = dg_plans.SelectedItems.OfType<Plan>().ToList();
            if (rows.Count == 0 && dg_plans.CurrentItem is Plan current)
            {
                rows.Add(current);
            }
            if (rows.Count == 0)
            {
                return;
            }
            List<string> lines = [];
            foreach (Plan row in rows)
            {
                lines.Add(string.Join("\t", CopyableFields.Select(f => GetFieldValue(row, f) ?? "")));
            }
            Clipboard.SetText(string.Join(Environment.NewLine, lines));
            PlansVM.StatusMessage = $"已复制 {rows.Count} 行";
        }

        /// <summary>
        /// 复制当前单元格
        /// </summary>
        private void Menu_CopyCell_Click(object sender, RoutedEventArgs e)
        {
            if (dg_plans.CurrentCell.Item is not Plan plan || dg_plans.CurrentCell.Column == null)
            {
                return;
            }
            string field = dg_plans.CurrentCell.Column.SortMemberPath;
            if (field == "Id" || string.IsNullOrEmpty(field))
            {
                return;
            }
            Clipboard.SetText(GetFieldValue(plan, field) ?? "");
            PlansVM.StatusMessage = "已复制单元格";
        }

        /// <summary>
        /// 粘贴：从剪贴板（TSV）按当前单元格起向右向下写入，逐格校验并标记暂存修改
        /// </summary>
        private void Menu_Paste_Click(object sender, RoutedEventArgs e)
        {
            if (PlansVM.IsGridReadOnly)
            {
                _ = MessageBox.Show("当前身份不支持表格内编辑", "提示");
                return;
            }
            if (dg_plans.CurrentCell.Item is not Plan startPlan || dg_plans.CurrentCell.Column == null)
            {
                _ = MessageBox.Show("请先选择粘贴起始单元格", "提示");
                return;
            }
            string text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            List<Plan> viewRows = PlansVM.PlansView.Cast<Plan>().ToList();
            int startRowIndex = viewRows.IndexOf(startPlan);
            int startColIndex = dg_plans.Columns.IndexOf(dg_plans.CurrentCell.Column);
            if (startRowIndex < 0 || startColIndex < 0)
            {
                return;
            }

            string[] lines = text.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries);
            int pastedCells = 0;
            List<string> errors = [];
            for (int i = 0; i < lines.Length; i++)
            {
                int rowIndex = startRowIndex + i;
                if (rowIndex >= viewRows.Count)
                {
                    break;
                }
                Plan row = viewRows[rowIndex];
                string[] cells = lines[i].Split('\t');
                for (int j = 0; j < cells.Length; j++)
                {
                    int colIndex = startColIndex + j;
                    if (colIndex >= dg_plans.Columns.Count)
                    {
                        break;
                    }
                    DataGridColumn column = dg_plans.Columns[colIndex];
                    if (column.IsReadOnly)
                    {
                        continue;
                    }
                    string field = column.SortMemberPath;
                    if (!CopyableFields.Contains(field))
                    {
                        continue;
                    }
                    string value = string.IsNullOrWhiteSpace(cells[j]) ? null : cells[j].Trim();
                    string error = PlansVM.ValidateField(field, value);
                    if (error != null)
                    {
                        errors.Add($"行{rowIndex + 1} {field}: {error}");
                        continue;
                    }
                    SetFieldValue(row, field, value);
                    PlansVM.MarkModified(row);
                    pastedCells++;
                }
            }
            PlansVM.StatusMessage = $"已粘贴 {pastedCells} 个单元格" + (errors.Count > 0 ? $"，{errors.Count} 个校验失败" : "");
            if (errors.Count > 0)
            {
                _ = MessageBox.Show("以下单元格校验失败，未粘贴：\n\n" + string.Join("\n", errors.Take(20)),
                    "粘贴校验", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static string GetFieldValue(Plan plan, string field)
        {
            PropertyInfo prop = typeof(Plan).GetProperty(field);
            return prop?.GetValue(plan)?.ToString();
        }

        private static void SetFieldValue(Plan plan, string field, string value)
        {
            PropertyInfo prop = typeof(Plan).GetProperty(field);
            if (prop != null && prop.PropertyType == typeof(string))
            {
                prop.SetValue(plan, value);
            }
        }

        /* ###############################  列顺序持久化  ################################ */

        /// <summary>
        /// 保存当前列显示顺序（列头名 -> DisplayIndex）
        /// </summary>
        private void SaveColumnLayout()
        {
            try
            {
                List<KeyValuePair<string, int>> layout = dg_plans.Columns
                    .OrderBy(c => c.DisplayIndex)
                    .Select(c => new KeyValuePair<string, int>(c.Header?.ToString() ?? "", c.DisplayIndex))
                    .ToList();
                File.WriteAllText(LayoutFilePath, JsonConvert.SerializeObject(layout, Formatting.Indented));
            }
            catch (Exception ex)
            {
                _logger.Warn($"保存列顺序布局失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 恢复上次保存的列显示顺序；文件不存在或不匹配时保持默认
        /// </summary>
        private void RestoreColumnLayout()
        {
            try
            {
                if (!File.Exists(LayoutFilePath))
                {
                    return;
                }
                var layout = JsonConvert.DeserializeObject<List<KeyValuePair<string, int>>>(File.ReadAllText(LayoutFilePath));
                if (layout == null)
                {
                    return;
                }
                foreach (KeyValuePair<string, int> kv in layout.OrderBy(kv => kv.Value))
                {
                    DataGridColumn col = dg_plans.Columns.FirstOrDefault(c => c.Header?.ToString() == kv.Key);
                    if (col != null && kv.Value >= 0 && kv.Value < dg_plans.Columns.Count)
                    {
                        col.DisplayIndex = kv.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"恢复列顺序布局失败: {ex.Message}");
            }
        }
    }
}
