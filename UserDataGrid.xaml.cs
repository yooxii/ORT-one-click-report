using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static ORT一键报告.ReportHeader;

namespace ORT一键报告
{
    // 列配置模型
    public class ColumnConfig : INotifyPropertyChanged
    {
        private string _header;
        private string _bindingPath;
        private double _width = 100;
        private bool _isVisible = true;
        private ColumnType _type = ColumnType.Text;

        public enum ColumnType
        {
            Text, ComboBox
        }

        // 列头显示文本
        public string Header
        {
            get => _header;
            set { _header = value; OnPropertyChanged(); }
        }

        // 绑定的数据属性名 (例如 "BIroom")
        public string BindingPath
        {
            get => _bindingPath;
            set { _bindingPath = value; OnPropertyChanged(); }
        }

        // 列宽
        public double Width
        {
            get => _width;
            set { _width = value; OnPropertyChanged(); }
        }

        // 是否可见 (控制列数量)
        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; OnPropertyChanged(); }
        }

        // 列类型
        public ColumnType Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
    /// <summary>
    /// UserDataGrid.xaml 的交互逻辑
    /// </summary>
    public partial class UserResultDataGrid : UserControl
    {
        public UserResultDataGrid()
        {
            InitializeComponent();
            // 监听集合变化，当外部添加/删除列配置时，自动刷新表格
            ColumnDefinitions.CollectionChanged += ColumnDefinitions_CollectionChanged;
        }

        // --- 核心：列定义集合 ---
        // 外部可以通过绑定或者直接 Add/Remove 来控制列
        public ObservableCollection<ColumnConfig> ColumnDefinitions { get; set; } = new ObservableCollection<ColumnConfig>();

        private void ColumnDefinitions_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            RebuildColumns();
        }

        // 重建列的逻辑
        private void RebuildColumns()
        {
            details_data.Columns.Clear(); // 1. 清空旧列

            foreach (var config in ColumnDefinitions)
            {
                if (!config.IsVisible)
                {
                    continue; // 2. 如果不可见则跳过
                }

                // 3. 创建新列
                switch (config.Type)
                {
                    case ColumnConfig.ColumnType.Text:
                        details_data.Columns.Add(new DataGridTextColumn
                        {
                            Header = config.Header,
                            Width = config.Width,
                            Binding = new Binding(config.BindingPath) // 绑定路径
                        });
                        break;
                    case ColumnConfig.ColumnType.ComboBox:
                        var comboCol = new DataGridComboBoxColumn
                        {
                            Header = config.Header,
                            Width = config.Width,
                            SelectedItemBinding = new Binding(config.BindingPath),
                            EditingElementStyle = (Style)FindResource("DynamicComboBoxStyle"),
                            ElementStyle = (Style)FindResource("DynamicComboBoxStyle")
                        };
                        details_data.Columns.Add(comboCol);
                        break;
                    default:
                        break;
                }
            }
        }

        // --- 依赖属性定义 ---

        // 1. 最小列宽 (对应 DataGrid.MinColumnWidth)
        public double MinColumnWidth
        {
            get => (double)GetValue(MinColumnWidthProperty);
            set => SetValue(MinColumnWidthProperty, value);
        }
        public static readonly DependencyProperty MinColumnWidthProperty =
            DependencyProperty.Register("MinColumnWidth", typeof(double), typeof(UserResultDataGrid), new PropertyMetadata(20.0));

        // 2. 列标题样式 (对应 DataGrid.ColumnHeaderStyle)
        public Style ColumnHeaderStyle
        {
            get => (Style)GetValue(ColumnHeaderStyleProperty);
            set => SetValue(ColumnHeaderStyleProperty, value);
        }
        public static readonly DependencyProperty ColumnHeaderStyleProperty =
            DependencyProperty.Register("ColumnHeaderStyle", typeof(Style), typeof(UserResultDataGrid), new PropertyMetadata(null));

        // 3. 表格源
        public ObservableCollection<ResultDetails> DataGridSource
        {
            get => (ObservableCollection<ResultDetails>)GetValue(DataGridSourceProperty);
            set => SetValue(DataGridSourceProperty, value);
        }

        public static readonly DependencyProperty DataGridSourceProperty =
            DependencyProperty.Register("DataGridSource", typeof(ObservableCollection<ResultDetails>), typeof(UserResultDataGrid), new PropertyMetadata(null));

        public int UUT_Count { set; get; } = 0;

        // --- 初始化表格 ---
        public void InitBurnColumns()
        {
            ColumnDefinitions.Add(new ColumnConfig { Header = "BIroom", BindingPath = "BIroom", Width = 100 });
            ColumnDefinitions.Add(new ColumnConfig { Header = "BIarea", BindingPath = "BIarea", Width = 100 });
            ColumnDefinitions.Add(new ColumnConfig { Header = "BIplace", BindingPath = "BIplace", Width = 100 });
            InitThermalColumns();
        }

        public void InitThermalColumns()
        {
            ColumnDefinitions.Add(new ColumnConfig { Header = "S/N", BindingPath = "SN", Width = 150 });
            ColumnDefinitions.Add(new ColumnConfig { Header = "工令", BindingPath = "WorkOrder", Width = 150 });
            ColumnDefinitions.Add(new ColumnConfig { Header = "版本", BindingPath = "Version", Width = 50 });
            ColumnDefinitions.Add(new ColumnConfig { Header = "周期", BindingPath = "DC", Width = 60 });
            ColumnDefinitions.Add(new ColumnConfig { Header = "外观-前", BindingPath = "InspectionPrev", Width = 50, Type = ColumnConfig.ColumnType.ComboBox });
            ColumnDefinitions.Add(new ColumnConfig { Header = "Fun-前", BindingPath = "FunPrev", Width = 50, Type = ColumnConfig.ColumnType.ComboBox });
            ColumnDefinitions.Add(new ColumnConfig { Header = "外观-后", BindingPath = "InspectionAfter", Width = 50, Type = ColumnConfig.ColumnType.ComboBox });
            ColumnDefinitions.Add(new ColumnConfig { Header = "Fun-后", BindingPath = "FunAfter", Width = 50, Type = ColumnConfig.ColumnType.ComboBox });
            ColumnDefinitions.Add(new ColumnConfig { Header = "Hi-Pot", BindingPath = "HiPot", Width = 50, Type = ColumnConfig.ColumnType.ComboBox });
            ColumnDefinitions.Add(new ColumnConfig { Header = "备注", BindingPath = "Comments", Width = 50 });

        }

        public void AddRow(int index = 0)
        {
            ResultDetails resultDetails = new ResultDetails
            {
                FunAfter = ReportStatus.Pass,
                FunPrev = ReportStatus.Pass,
                InspectionAfter = ReportStatus.Pass,
                InspectionPrev = ReportStatus.Pass,
                HiPot = ReportStatus.Pass,
            };
            if (index < 0 || index > DataGridSource.Count)
            {
                index = DataGridSource.Count; // 如果索引无效，则添加到末尾
            }
            DataGridSource.Insert(index, resultDetails);
        }

        public void DeleteRow(int index)
        {
            if (index < 0 || index >= DataGridSource.Count)
            {
                return; // 索引无效，直接返回
            }
            DataGridSource.RemoveAt(index);
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            if (details_data.SelectedCells.Count == 0)
            {
                AddRow(-1);
                return;
            }
            AddRow(details_data.Items.IndexOf(details_data.SelectedCells[0].Item));
        }

        private void DelRow_Click(object sender, RoutedEventArgs e)
        {
            if (details_data.SelectedCells.Count == 0) return;
            DeleteRow(details_data.Items.IndexOf(details_data.SelectedCells[0].Item));
        }

        private void EqualFirstRow_Click(object sender, RoutedEventArgs e)
        {
            // 1. 基础检查：确保有选中单元格，且不是表头
            if (details_data.SelectedCells.Count == 0) return;

            // 获取当前选中的单元格信息
            DataGridCellInfo currentCell = details_data.SelectedCells[0];

            // 2. 获取当前列的信息
            DataGridColumn currentColumn = currentCell.Column;

            // 3. 获取该列绑定的属性名
            string propertyName = "";
            if (currentColumn is DataGridBoundColumn boundCol && boundCol.Binding is Binding binding)
            {
                propertyName = binding.Path.Path;
            }

            if (string.IsNullOrEmpty(propertyName)) return;

            // 4. 获取第一行 (索引为 0) 的数据对象
            // 注意：这里直接取 Items[0]，如果表格为空需加判断
            if (details_data.Items.Count == 0) return;
            object firstRowItem = details_data.Items[0];

            // 5. 读取第一行该属性的值
            PropertyInfo prop = firstRowItem.GetType().GetProperty(propertyName);
            object firstRowValue = prop?.GetValue(firstRowItem, null);

            for (int i = 1; i < DataGridSource.Count; i++)
            {
                DataGridSource[i].GetType().GetProperty(propertyName).SetValue(DataGridSource[i], firstRowValue, null);
            }
            details_data.Items.Refresh();
        }

        private void ClearCol_Click(object sender, RoutedEventArgs e)
        {
            if (details_data.SelectedCells.Count == 0) return;

            // 获取当前选中的单元格信息
            DataGridCellInfo currentCell = details_data.SelectedCells[0];

            // 2. 获取当前列的信息
            DataGridColumn currentColumn = currentCell.Column;

            // 3. 获取该列绑定的属性名
            string propertyName = "";
            if (currentColumn is DataGridBoundColumn boundCol && boundCol.Binding is Binding binding)
            {
                propertyName = binding.Path.Path;
            }
            if (string.IsNullOrEmpty(propertyName)) return;

            for (int i = 0; i < DataGridSource.Count; i++)
            {
                DataGridSource[i].GetType().GetProperty(propertyName).SetValue(DataGridSource[i], null, null);
            }
            details_data.Items.Refresh();
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            details_data.Items.Refresh();
        }
    }
}
