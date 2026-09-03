using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using ORT一键报告.Services;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 单条日志记录
    /// </summary>
    public class LogEntry
    {
        public DateTime Time { get; set; }
        public string TimeText { get; set; }
        public string Level { get; set; }
        public string Logger { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// WindowLog.xaml 的交互逻辑
    /// </summary>
    public partial class WindowLog : Window
    {
        private readonly string _logDir;
        private readonly List<LogEntry> _allEntries = [];
        private readonly CollectionViewSource _viewSource = new();
        private readonly DispatcherTimer _refreshTimer = new() { Interval = TimeSpan.FromSeconds(2) };
        private long _lastFileLength = -1;

        public WindowLog()
        {
            InitializeComponent();

            _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            _viewSource.Source = _allEntries;
            _viewSource.View.Filter = LogFilter;
            dg_logs.ItemsSource = _viewSource.View;

            _refreshTimer.Tick += (s, e) => TryAutoReload();
            Loaded += WindowLog_Loaded;
            Closed += (s, e) => _refreshTimer.Stop();
        }

        private void WindowLog_Loaded(object sender, RoutedEventArgs e)
        {
            LoadLogFiles();
            _refreshTimer.Start();
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 加载logs目录下所有日志文件，默认选中最新的一个
        /// </summary>
        private void LoadLogFiles()
        {
            cb_logFiles.Items.Clear();
            if (!Directory.Exists(_logDir))
            {
                status_file.Content = string.Format(LanguageService.Get("Log_DirNotExist"), _logDir);
                return;
            }
            string[] files = Directory.GetFiles(_logDir, "*.log")
                .OrderByDescending(f => f)
                .ToArray();
            foreach (string file in files)
            {
                cb_logFiles.Items.Add(new ComboBoxItem { Content = Path.GetFileName(file), Tag = file });
            }
            if (cb_logFiles.Items.Count > 0)
            {
                cb_logFiles.SelectedIndex = 0;
            }
            else
            {
                status_file.Content = LanguageService.Get("Log_NoLogFiles");
            }
        }

        /// <summary>
        /// 解析并加载选中的日志文件
        /// </summary>
        private void LoadLogFile(string filePath)
        {
            _allEntries.Clear();
            if (!File.Exists(filePath))
            {
                _viewSource.View.Refresh();
                UpdateStatus();
                return;
            }

            LogEntry last = null;
            try
            {
                using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using StreamReader reader = new(fs);
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    LogEntry entry = ParseLogLine(line);
                    if (entry != null)
                    {
                        _allEntries.Add(entry);
                        last = entry;
                    }
                    else if (last != null)
                    {
                        // 异常堆栈等续行，并入上一条记录
                        last.Message += Environment.NewLine + line;
                    }
                }
                _lastFileLength = fs.Length;
            }
            catch (Exception ex)
            {
                status_file.Content = string.Format(LanguageService.Get("Log_ReadFailed"), ex.Message);
                return;
            }

            // 默认按时间降序显示
            _viewSource.SortDescriptions.Clear();
            _viewSource.SortDescriptions.Add(new SortDescription(nameof(LogEntry.Time), ListSortDirection.Descending));
            _viewSource.View.Refresh();
            UpdateStatus();
            status_file.Content = filePath;
        }

        /// <summary>
        /// 解析一行日志，格式: "2026-08-06 12:00:00.1234 | INFO | Logger | Message"
        /// 解析失败（如异常堆栈续行）返回null
        /// </summary>
        private static LogEntry ParseLogLine(string line)
        {
            string[] parts = line.Split(new[] { " | " }, 4, StringSplitOptions.None);
            if (parts.Length < 4)
            {
                return null;
            }
            if (!DateTime.TryParseExact(parts[0].Trim(),
                    ["yyyy-MM-dd HH:mm:ss.ffff", "yyyy-MM-dd HH:mm:ss"],
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime time))
            {
                return null;
            }
            return new LogEntry
            {
                Time = time,
                TimeText = parts[0].Trim(),
                Level = parts[1].Trim(),
                Logger = parts[2].Trim(),
                Message = parts[3]
            };
        }

        /// <summary>
        /// 级别筛选 + 关键字搜索
        /// </summary>
        private bool LogFilter(object obj)
        {
            if (obj is not LogEntry entry)
            {
                return false;
            }
            if (cb_level.SelectedItem is ComboBoxItem levelItem
                && levelItem.Tag is string level
                && !string.IsNullOrEmpty(level)
                && !string.Equals(entry.Level, level, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            string keyword = txt_search.Text?.Trim();
            if (!string.IsNullOrEmpty(keyword))
            {
                return (entry.Message?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    || (entry.Logger?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    || (entry.Level?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return true;
        }

        /// <summary>
        /// 自动刷新: 仅当日志文件大小变化时重新加载
        /// </summary>
        private void TryAutoReload()
        {
            if (chk_autoRefresh.IsChecked != true)
            {
                return;
            }
            if (GetSelectedLogFile() is not string filePath || !File.Exists(filePath))
            {
                return;
            }
            long length;
            try
            {
                length = new FileInfo(filePath).Length;
            }
            catch
            {
                return;
            }
            if (length != _lastFileLength)
            {
                LoadLogFile(filePath);
            }
        }

        private string GetSelectedLogFile()
        {
            return cb_logFiles.SelectedItem is ComboBoxItem item ? item.Tag as string : null;
        }

        private void UpdateStatus()
        {
            status_count.Content = string.Format(LanguageService.Get("Log_CountFiltered"), _allEntries.Count, _viewSource.View.Cast<object>().Count());
        }

        /* ###############################  事件函数  ################################ */

        private void Cb_LogFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GetSelectedLogFile() is string filePath)
            {
                LoadLogFile(filePath);
            }
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            // InitializeComponent解析XAML时(如IsSelected="True")可能提前触发本事件，此时字段尚未初始化
            if (_viewSource?.View == null)
            {
                return;
            }
            _viewSource.View.Refresh();
            UpdateStatus();
        }

        private void Btn_Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadLogFiles();
            if (GetSelectedLogFile() is string filePath)
            {
                LoadLogFile(filePath);
            }
        }
    }
}
