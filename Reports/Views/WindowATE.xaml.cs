using Microsoft.Win32;
using NLog;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using static ORT一键报告.Utils.Report;

namespace ORT一键报告.Reports.Views
{
    /// <summary>
    /// ATE 数据工具：读取 ATE 原始数据（xls/xlsx），人工选择试验前/试验后数据（可拖动排序、可直接改数值），
    /// 超出上下限的数据标红并可一键跳转，最后保存成报告文件或直接提交给某类报告。
    /// </summary>
    public partial class ATEWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 「提交到报告」的回调（报告类型, 生成好的文件路径）→ 是否成功；由打开本窗口的一键报告窗口提供
        /// </summary>
        public Func<string, string, bool> SubmitHandler { get; set; }

        /// <summary>当前打开的 ATE 原始数据文件</summary>
        public string ATEFilePath { get; private set; }

        private AteSheet _sheet;
        private readonly ObservableCollection<RowVm> _pool = [];
        private readonly ObservableCollection<RowVm> _pre = [];
        private readonly ObservableCollection<RowVm> _post = [];
        private ItemsControl _activeControl;
        private int _badCursor;
        private int _suppressCount;

        private static readonly Brush BadForeground = new SolidColorBrush(Color.FromRgb(0xE5, 0x39, 0x35));
        private static readonly Brush BadBackground = new SolidColorBrush(Color.FromArgb(0x33, 0xE5, 0x39, 0x35));
        private static readonly Brush DropHighlightBrush = new SolidColorBrush(Color.FromRgb(0x4A, 0x90, 0xE2));

        /// <summary>三个分组区域（用于拖动时高亮"这一行会落到哪一组"）</summary>
        private sealed class PaneInfo
        {
            public Control Body { get; set; }
            public TextBlock Title { get; set; }
            public Brush BorderBrushDefault { get; set; }
            public Thickness BorderThicknessDefault { get; set; }
            public Brush TitleBrushDefault { get; set; }
        }

        private readonly List<PaneInfo> _panes = [];

        public ATEWindow()
        {
            InitializeComponent();
            lb_pool.ItemsSource = _pool;
            dg_pre.ItemsSource = _pre;
            dg_post.ItemsSource = _post;
            _activeControl = lb_pool;
            _panes.Add(new PaneInfo
            {
                Body = dg_pre,
                Title = txt_preTitle,
                BorderBrushDefault = dg_pre.BorderBrush,
                BorderThicknessDefault = dg_pre.BorderThickness,
                TitleBrushDefault = txt_preTitle.Foreground
            });
            _panes.Add(new PaneInfo
            {
                Body = dg_post,
                Title = txt_postTitle,
                BorderBrushDefault = dg_post.BorderBrush,
                BorderThicknessDefault = dg_post.BorderThickness,
                TitleBrushDefault = txt_postTitle.Foreground
            });
            _panes.Add(new PaneInfo
            {
                Body = lb_pool,
                Title = null,
                BorderBrushDefault = lb_pool.BorderBrush,
                BorderThicknessDefault = lb_pool.BorderThickness,
                TitleBrushDefault = null
            });
            Loaded += ATEWindow_Loaded;
            Closed += (s, e) => Utils.Report.ClearTempDir();
        }

        private void ATEWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 可提交的报告类型：直接取静态列表，不依赖 Owner（ATE 窗口已不再给报告窗口当属主）
            cmb_reportType.ItemsSource = WindowMainReport.AteReportTypes;
            if (cmb_reportType.Items.Count > 0)
            {
                cmb_reportType.SelectedIndex = 0;
            }
            if (string.IsNullOrWhiteSpace(text_ATETemplate.Text))
            {
                try
                {
                    text_ATETemplate.Text = GetATETemplate().FullName;
                }
                catch (Exception ex)
                {
                    _logger.Warn($"未能定位默认 ATE 模板：{ex.Message}");
                }
            }
            UpdateCounts();
        }

        /* ###############################  数据模型  ################################ */

        /// <summary>一个测试值单元格：可直接编辑，超出限值时标红</summary>
        private sealed class CellVm : INotifyPropertyChanged
        {
            private readonly RowVm _owner;
            private readonly int _index;
            private string _value;

            public CellVm(RowVm owner, int index, string value)
            {
                _owner = owner;
                _index = index;
                _value = value;
            }

            public string Value
            {
                get => _value;
                set
                {
                    if (_value == value)
                    {
                        return;
                    }
                    _value = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsOutOfSpec));
                    _owner.NotifyChanged();
                }
            }

            /// <summary>是否超出上下限（上下限为空或 "*" 表示不考虑该侧）</summary>
            public bool IsOutOfSpec => _owner.Sheet != null
                && _index < _owner.Sheet.ItemCount
                && AteReportData.IsOutOfSpec(_value, _owner.Sheet.MaxSpecs[_index], _owner.Sheet.MinSpecs[_index]);

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged([CallerMemberName] string name = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        /// <summary>一行数据（一个样品的全部测试值）</summary>
        private sealed class RowVm : INotifyPropertyChanged
        {
            private string _sn;

            public RowVm(AteSheet sheet, AteRow row, int sourceIndex)
            {
                Sheet = sheet;
                SourceIndex = sourceIndex;
                _sn = row.SN;
                Cells = new ObservableCollection<CellVm>(row.Values.Select((v, i) => new CellVm(this, i, v)));
            }

            public AteSheet Sheet { get; }
            public int SourceIndex { get; }
            public ObservableCollection<CellVm> Cells { get; }

            public string SN
            {
                get => _sn;
                set
                {
                    if (_sn == value)
                    {
                        return;
                    }
                    _sn = value;
                    OnPropertyChanged();
                }
            }

            public int BadCount => Cells.Count(c => c.IsOutOfSpec);

            public AteRow ToAteRow() => new() { SN = SN, Values = Cells.Select(c => c.Value).ToList() };

            public void NotifyChanged() => OnPropertyChanged(nameof(BadCount));

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged([CallerMemberName] string name = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        /* ###############################  读取数据  ################################ */

        private async void OpenATEDatas_Click(object sender, RoutedEventArgs e)
        {
            AppSettingsService settings = App.ServiceProvider.GetService(typeof(AppSettingsService)) as AppSettingsService;
            OpenFileDialog dialog = new()
            {
                Filter = "ATE|*.xls;*.xlsx",
                InitialDirectory = settings?.AteDataDir
            };
            if (dialog.ShowDialog() != true)
            {
                _logger.Warn("未选择ATE数据文件");
                _ = MessageBox.Show(LanguageService.Get("Msg_NoATEFile"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            await LoadFromFileAsync(dialog.FileName);
        }

        private async void btn_reread_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ATEFilePath))
            {
                _ = MessageBox.Show(LanguageService.Get("Msg_NoATEFile"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            await LoadFromFileAsync(ATEFilePath);
        }

        /// <summary>读取 ATE 原始数据并重建界面（自动分组，不可靠时全部放到未使用让用户自己选）</summary>
        private async Task LoadFromFileAsync(string fileName)
        {
            PopupWindow popup = PopupWindow.ShowBusy(LanguageService.Get("ATE_Opening"), this);
            try
            {
                // 先让"正在打开"提示画出来，再放到后台线程读取，界面不会假死
                await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                AteSheet sheet = await Task.Run(() => AteReportData.Read(fileName));
                if (sheet == null)
                {
                    _logger.Warn($"ATE 数据文件里找不到 S/N / MAX_SPEC 行：{fileName}");
                    _ = MessageBox.Show(LanguageService.Get("Msg_NoATEData"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                _sheet = sheet;
                ATEFilePath = fileName;
                BuildRows();
                BuildColumns();
                UpdateCounts();
                _logger.Info($"ATE 数据读取完成：{fileName}（{sheet.Rows.Count} 条数据 / {sheet.ItemCount} 个测试项）");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取ATE数据发生错误");
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                popup.Close();
            }
        }

        private void BuildRows()
        {
            _pool.Clear();
            _pre.Clear();
            _post.Clear();
            _badCursor = 0;
            List<RowVm> rows = _sheet.Rows.Select((r, i) => new RowVm(_sheet, r, i)).ToList();
            bool[] guess = AteReportData.GuessIsBefore(_sheet.Rows.Select(r => r.SN).ToList());
            if (guess != null && guess.Count(x => x) <= _sheet.MaxPerGroup && guess.Count(x => !x) <= _sheet.MaxPerGroup)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    (guess[i] ? _pre : _post).Add(rows[i]);
                }
                _logger.Info($"ATE 自动分组：试验前 {_pre.Count} 条 / 试验后 {_post.Count} 条");
            }
            else
            {
                // 人工填写的数据分组规律不可靠：全部放进"未使用"，由用户自己选
                foreach (RowVm row in rows)
                {
                    _pool.Add(row);
                }
                _logger.Warn("ATE 自动分组不可靠，已把全部数据放到「未使用」，请手动选择试验前/试验后");
            }
        }

        /* ###############################  界面列  ################################ */

        private void BuildColumns()
        {
            BuildGridColumns(dg_pre);
            BuildGridColumns(dg_post);
        }

        private void BuildGridColumns(DataGrid grid)
        {
            grid.Columns.Clear();
            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "S/N",
                Binding = new Binding(nameof(RowVm.SN)) { Mode = BindingMode.TwoWay },
                Width = 150
            });
            for (int i = 0; i < _sheet.ItemCount; i++)
            {
                int index = i;
                grid.Columns.Add(new DataGridTextColumn
                {
                    Header = BuildHeader(index),
                    Binding = new Binding($"Cells[{index}].Value") { Mode = BindingMode.TwoWay },
                    Width = 96,
                    ElementStyle = BadStyle(index, false),
                    EditingElementStyle = BadStyle(index, true)
                });
            }
        }

        /// <summary>列标题：测试项目名 + 上下限（鼠标悬停看测试条件）</summary>
        private object BuildHeader(int index)
        {
            string max = AteReportData.SpecText(_sheet.MaxSpecs[index]);
            string min = AteReportData.SpecText(_sheet.MinSpecs[index]);
            return new TextBlock
            {
                Text = $"{_sheet.OutputTypes[index]}\n{min} ~ {max}",
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                ToolTip = $"{_sheet.Conditions[index]}\n{LanguageService.Get("ATE_Limits")}: {min} ~ {max}"
            };
        }

        /// <summary>超出限值的单元格标红（显示态与编辑态都要标）</summary>
        private static Style BadStyle(int index, bool editing)
        {
            Style style = new(editing ? typeof(TextBox) : typeof(TextBlock));
            DataTrigger trigger = new()
            {
                Binding = new Binding($"Cells[{index}].IsOutOfSpec"),
                Value = true
            };
            trigger.Setters.Add(new Setter(editing ? Control.ForegroundProperty : TextBlock.ForegroundProperty, BadForeground));
            trigger.Setters.Add(new Setter(editing ? Control.BackgroundProperty : TextBlock.BackgroundProperty, BadBackground));
            style.Triggers.Add(trigger);
            return style;
        }

        private void UpdateCounts()
        {
            if (_suppressCount > 0)
            {
                return;
            }
            int max = _sheet?.MaxPerGroup ?? 0;
            txt_preTitle.Text = $"{LanguageService.Get("ATE_PreTest")}  {_pre.Count}/{max}";
            txt_postTitle.Text = $"{LanguageService.Get("ATE_PostTest")}  {_post.Count}/{max}";
            txt_counts.Text = string.Format(LanguageService.Get("ATE_PoolCount"), _pool.Count);
            int bad = _pre.Concat(_post).Sum(r => r.BadCount);
            txt_bad.Text = bad > 0 ? string.Format(LanguageService.Get("ATE_BadCount"), bad) : "";
        }

        /* ###############################  拖动排序 / 换组  ################################ */

        private Point _dragStart;
        private RowVm _dragRow;
        private bool _dragArmed;

        private static T FindAncestor<T>(DependencyObject source) where T : DependencyObject
        {
            while (source != null && source is not T)
            {
                source = VisualTreeHelper.GetParent(source) ?? LogicalTreeHelper.GetParent(source);
            }
            return source as T;
        }

        private void Pool_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _activeControl = lb_pool;
            RowVm row = (FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject) as FrameworkElement)?.DataContext as RowVm;
            _dragRow = row;
            _dragArmed = row != null;
            _dragStart = e.GetPosition(null);
        }

        private void Pool_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            StartDragIfNeeded(lb_pool, e);
        }

        private void Grid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DataGrid grid = (DataGrid)sender;
            _activeControl = grid;
            // 只有从行标题（左侧窄条）按住才能拖动，避免影响单元格编辑与选择
            bool onRowHeader = FindAncestor<DataGridRowHeader>(e.OriginalSource as DependencyObject) != null;
            DataGridRow container = FindAncestor<DataGridRow>(e.OriginalSource as DependencyObject);
            _dragRow = onRowHeader ? container?.Item as RowVm : null;
            _dragArmed = _dragRow != null;
            _dragStart = e.GetPosition(null);
        }

        private void Grid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            StartDragIfNeeded((DataGrid)sender, e);
        }

        private void StartDragIfNeeded(DependencyObject source, MouseEventArgs e)
        {
            if (!_dragArmed || _dragRow == null || e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }
            Vector delta = e.GetPosition(null) - _dragStart;
            if (Math.Abs(delta.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(delta.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }
            _dragArmed = false;
            _ = DragDrop.DoDragDrop(source, _dragRow, DragDropEffects.Move);
            HighlightDropTarget(null);
        }

        private void DropTarget_DragOver(object sender, DragEventArgs e)
        {
            bool acceptable = e.Data.GetDataPresent(typeof(RowVm));
            e.Effects = acceptable ? DragDropEffects.Move : DragDropEffects.None;
            // 拖到哪个分组就高亮哪个分组，避免"拖到别的组去了"却看不出来
            HighlightDropTarget(acceptable ? sender as Control : null);
            e.Handled = true;
        }

        private void DropTarget_DragLeave(object sender, DragEventArgs e)
        {
            HighlightDropTarget(null);
        }

        /// <summary>高亮将要接收这一行的分组（传入 null 表示全部取消高亮）</summary>
        private void HighlightDropTarget(Control target)
        {
            foreach (PaneInfo pane in _panes)
            {
                bool on = ReferenceEquals(pane.Body, target);
                pane.Body.BorderBrush = on ? DropHighlightBrush : pane.BorderBrushDefault;
                pane.Body.BorderThickness = on ? new Thickness(2) : pane.BorderThicknessDefault;
                if (pane.Title != null)
                {
                    pane.Title.Foreground = on ? DropHighlightBrush : pane.TitleBrushDefault;
                }
            }
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(RowVm)) is not RowVm row)
            {
                return;
            }
            DataGrid grid = (DataGrid)sender;
            ObservableCollection<RowVm> target = ReferenceEquals(grid, dg_pre) ? _pre : _post;
            _ = MoveRow(row, target, DropIndex(grid, e.GetPosition(grid)));
            HighlightDropTarget(null);
            e.Handled = true;
        }

        private void Pool_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(typeof(RowVm)) is not RowVm row)
            {
                return;
            }
            _ = MoveRow(row, _pool, _pool.Count);
            HighlightDropTarget(null);
            e.Handled = true;
        }

        private static int DropIndex(DataGrid grid, Point point)
        {
            DataGridRow container = FindAncestor<DataGridRow>(grid.InputHitTest(point) as DependencyObject);
            if (container == null)
            {
                return grid.Items.Count;
            }
            int index = grid.ItemContainerGenerator.IndexFromContainer(container);
            return index < 0 ? grid.Items.Count : index;
        }

        /// <summary>
        /// 把一行数据移到目标分组（未使用/试验前/试验后）；每组最多 AteSheet.MaxPerGroup 条
        /// </summary>
        private bool MoveRow(RowVm row, ObservableCollection<RowVm> target, int index)
        {
            if (target != _pool && target.Count >= (_sheet?.MaxPerGroup ?? int.MaxValue) && !target.Contains(row))
            {
                string group = ReferenceEquals(target, _pre) ? LanguageService.Get("ATE_PreTest") : LanguageService.Get("ATE_PostTest");
                _ = MessageBox.Show(
                    string.Format(LanguageService.Get("ATE_CapacityFull"), group, _sheet.MaxPerGroup),
                    LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            ObservableCollection<RowVm> source = _pool.Contains(row) ? _pool : _pre.Contains(row) ? _pre : _post;
            int oldIndex = source.IndexOf(row);
            if (ReferenceEquals(source, target) && oldIndex >= 0 && oldIndex < index)
            {
                index--;
            }
            source.Remove(row);
            index = Math.Max(0, Math.Min(index, target.Count));
            target.Insert(index, row);
            UpdateCounts();
            return true;
        }

        private void MoveToPre_Click(object sender, RoutedEventArgs e) => MoveSelected(_pre);

        private void MoveToPost_Click(object sender, RoutedEventArgs e) => MoveSelected(_post);

        private void MoveToPool_Click(object sender, RoutedEventArgs e) => MoveSelected(_pool);

        private void MoveSelected(ObservableCollection<RowVm> target)
        {
            List<RowVm> selected = SelectedRows();
            if (selected.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("ATE_NoSelection"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            foreach (RowVm row in selected)
            {
                _ = MoveRow(row, target, target.Count);
            }
        }

        /// <summary>当前焦点所在控件里选中的数据行</summary>
        private List<RowVm> SelectedRows()
        {
            if (_activeControl is ListBox list)
            {
                return list.SelectedItems.Cast<RowVm>().ToList();
            }
            if (_activeControl is DataGrid grid)
            {
                // selected cells or whole selected rows (clicking the row header) both work
                List<RowVm> rows = grid.SelectedCells.Select(c => c.Item as RowVm).Where(r => r != null).Distinct().ToList();
                if (rows.Count == 0)
                {
                    rows = grid.SelectedItems.Cast<RowVm>().ToList();
                }
                return rows;
            }
            return [];
        }

        /* ###############################  超限跳转  ################################ */

        private void NextOutOfSpec_Click(object sender, RoutedEventArgs e)
        {
            List<(DataGrid Grid, RowVm Row, int Index)> bad = [];
            foreach (DataGrid grid in new[] { dg_pre, dg_post })
            {
                foreach (RowVm row in ReferenceEquals(grid, dg_pre) ? _pre : _post)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        if (row.Cells[i].IsOutOfSpec)
                        {
                            bad.Add((grid, row, i));
                        }
                    }
                }
            }
            if (bad.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("ATE_NoOutOfSpec"), LanguageService.Get("ATE_Title"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            _badCursor %= bad.Count;
            (DataGrid badGrid, RowVm badRow, int badIndex) = bad[_badCursor];
            _badCursor = (_badCursor + 1) % bad.Count;
            // 第 0 列是 S/N，项目列从 1 开始
            if (badIndex + 1 < badGrid.Columns.Count)
            {
                DataGridColumn column = badGrid.Columns[badIndex + 1];
                badGrid.ScrollIntoView(badRow, column);
                badGrid.CurrentCell = new DataGridCellInfo(badRow, column);
                badGrid.SelectedCells.Clear();
                badGrid.SelectedCells.Add(badGrid.CurrentCell);
                _ = badGrid.Focus();
                badGrid.BeginEdit();
            }
            _logger.Info($"跳到第 {_badCursor}/{bad.Count} 处超限：{badRow.SN} / {_sheet.OutputTypes[badIndex]} = {badRow.Cells[badIndex].Value}");
            txt_bad.Text = string.Format(LanguageService.Get("ATE_BadCount"), bad.Count);
        }

        /* ###############################  生成 / 提交  ################################ */

        private async void SaveATEDatas_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded())
            {
                return;
            }
            FileInfo template;
            try
            {
                template = GetATETemplate();
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_ATEReportFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            string extension = template.Extension;
            SaveFileDialog dialog = new()
            {
                FileName = Path.GetFileNameWithoutExtension(ATEFilePath) + " ATE report" + extension,
                Filter = $"Excel|*{extension}",
                InitialDirectory = Path.GetDirectoryName(ATEFilePath)
            };
            if (dialog.ShowDialog() != true)
            {
                return;
            }
            (bool ok, string error) = await GenerateReportAsync(template, dialog.FileName);
            if (ok)
            {
                _ = MessageBox.Show(string.Format(LanguageService.Get("ATE_Saved"), dialog.FileName), LanguageService.Get("ATE_Title"), MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_ATEReportFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void SubmitToReport_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureDataLoaded())
            {
                return;
            }
            if (SubmitHandler == null)
            {
                _ = MessageBox.Show(LanguageService.Get("ATE_SubmitNoTarget"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (cmb_reportType.SelectedItem is not string reportType)
            {
                _ = MessageBox.Show(LanguageService.Get("ATE_NoReportType"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            FileInfo template;
            try
            {
                template = GetATETemplate();
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_ATEReportFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            AppSettingsService settings = App.ServiceProvider.GetService(typeof(AppSettingsService)) as AppSettingsService;
            string dir = settings?.AteDataDir;
            if (string.IsNullOrWhiteSpace(dir))
            {
                dir = Path.GetDirectoryName(ATEFilePath);
            }
            string outputPath = Path.Combine(dir, Path.GetFileNameWithoutExtension(ATEFilePath) + " ATE report" + template.Extension);
            (bool generated, string error) = await GenerateReportAsync(template, outputPath);
            if (!generated)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_ATEReportFailed"), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (SubmitHandler(reportType, outputPath))
            {
                _ = MessageBox.Show(string.Format(LanguageService.Get("ATE_Submitted"), reportType, outputPath), LanguageService.Get("ATE_Title"), MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                _ = MessageBox.Show(LanguageService.Get("ATE_SubmitNoTarget"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// 生成 ATE 报告：数据列多时比较慢，先弹"正在处理"等待窗口（带进度条），
        /// 再放到后台线程写文件，界面不会假死。
        /// </summary>
        private async Task<(bool Ok, string Error)> GenerateReportAsync(FileInfo template, string outputPath)
        {
            List<AteRow> before = _pre.Select(r => r.ToAteRow()).ToList();
            List<AteRow> after = _post.Select(r => r.ToAteRow()).ToList();
            PopupWindow popup = PopupWindow.ShowBusy(LanguageService.Get("ATE_Saving"), this);
            try
            {
                await System.Windows.Threading.Dispatcher.Yield(System.Windows.Threading.DispatcherPriority.Background);
                await Task.Run(() => AteReportData.Write(_sheet, before, after, template.FullName, outputPath));
                _logger.Info($"ATE 报告已生成：{outputPath}（试验前 {before.Count} 条 / 试验后 {after.Count} 条）");
                return (true, null);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, LanguageService.Get("Cap_ATEReportFailed"));
                return (false, ex.Message);
            }
            finally
            {
                popup.Close();
            }
        }

        private bool EnsureDataLoaded()
        {
            if (_sheet != null && !string.IsNullOrWhiteSpace(ATEFilePath))
            {
                return true;
            }
            _ = MessageBox.Show(LanguageService.Get("Msg_NoATEFile"), LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private FileInfo GetATETemplate()
        {
            string text = text_ATETemplate.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(text) && text != LanguageService.Get("Common_PleaseSelect"))
            {
                FileInfo info = new(Path.GetFullPath(text));
                if (info.Exists)
                {
                    return info;
                }
            }
            string defaultPath = GetTemplatePath(Path.Combine(Directory.GetCurrentDirectory(), "Templates"), "ATE");
            if (File.Exists(defaultPath))
            {
                text_ATETemplate.Text = defaultPath;
                return new FileInfo(defaultPath);
            }
            throw new FileNotFoundException(LanguageService.Get("Msg_NoATETemplate"));
        }

        private void btn_ATETemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new()
            {
                Filter = "ATE|*.xls;*.xlsx"
            };
            if (dialog.ShowDialog() == true)
            {
                text_ATETemplate.Text = dialog.FileName;
                return;
            }
            string defaultPath = GetTemplatePath(Path.Combine(Directory.GetCurrentDirectory(), "Templates"), "ATE");
            if (File.Exists(defaultPath))
            {
                text_ATETemplate.Text = defaultPath;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
