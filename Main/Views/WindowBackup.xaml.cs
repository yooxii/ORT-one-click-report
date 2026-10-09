using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowBackup.xaml 的交互逻辑：数据库快照与还原。
    /// 列出数据库同级的 Backups 文件夹里的全量/增量备份，可手动补做备份、删除旧备份，
    /// 也可以把数据库还原到选定的那份快照（仅管理员）。还原会替换数据库文件并要求重启程序。
    /// </summary>
    public partial class WindowBackup : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseBackupService _backup;
        private readonly bool _isAdmin;

        /// <summary>正在备份/还原（期间按钮不可用，避免并发操作）</summary>
        private bool _busy;

        public WindowBackup()
        {
            InitializeComponent();
            _backup = App.ServiceProvider.GetRequiredService<DatabaseBackupService>();
            _isAdmin = App.ServiceProvider.GetRequiredService<IPermissionService>().Can("admin.manage");
            Loaded += (s, e) => RefreshList();
        }

        /* ###############################  列表  ################################ */

        /// <summary>刷新备份清单与顶部文件夹/状态</summary>
        private void RefreshList()
        {
            try
            {
                List<DatabaseBackupEntry> entries = _backup.ListBackups();
                txt_folder.Text = _backup.ResolveBackupRoot();
                dg_backups.ItemsSource = entries.Select(e => new BackupRow(e)).ToList();
                DatabaseBackupEntry latest = entries.FirstOrDefault();
                txt_status.Text = latest == null
                    ? LanguageService.Get("Backup_None")
                    : string.Format(LanguageService.Get("Backup_LastFormat"), latest.CreatedAt.ToString("yyyy/M/d HH:mm"));
                btn_restore.IsEnabled = _isAdmin && !_busy;
                btn_delete.IsEnabled = _isAdmin && !_busy;
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取备份清单失败: {ex.Message}");
                txt_status.Text = ex.Message;
            }
        }

        /// <summary>清单行：把备份数据翻译成界面文字</summary>
        private sealed class BackupRow
        {
            public BackupRow(DatabaseBackupEntry entry)
            {
                Entry = entry;
            }

            public DatabaseBackupEntry Entry { get; }

            public string TimeText => Entry.CreatedAt.ToString("yyyy/M/d HH:mm:ss");

            public string KindText => LanguageService.Get(Entry.Kind == DatabaseBackupKind.Full
                ? "Backup_Kind_Full"
                : "Backup_Kind_Incremental");

            public string SizeText => Entry.SizeText;

            public string ChangedText => Entry.Kind == DatabaseBackupKind.Incremental
                ? Entry.ChangedRows.ToString()
                : "-";

            public string BaseFileName => string.IsNullOrWhiteSpace(Entry.BaseFileName) ? "-" : Entry.BaseFileName;

            public string FileName => Entry.FileName;

            public string StateText => Entry.CanRestore ? LanguageService.Get("Backup_State_Ok") : Entry.Problem;
        }

        /* ###############################  手动备份  ################################ */

        private async void Btn_Full_Click(object sender, RoutedEventArgs e)
        {
            await RunBackupAsync(() => _backup.CreateFullBackup("手动全量备份"));
        }

        private async void Btn_Incremental_Click(object sender, RoutedEventArgs e)
        {
            await RunBackupAsync(() => _backup.CreateIncrementalBackup("手动增量备份"));
        }

        /// <summary>在后台线程执行一次备份，期间禁用按钮并把结果写到状态栏</summary>
        private async Task RunBackupAsync(Func<DatabaseBackupResult> action)
        {
            if (_busy)
            {
                return;
            }
            SetBusy(true, LanguageService.Get("Backup_Busy"));
            try
            {
                DatabaseBackupResult result = await Task.Run(action);
                txt_status.Text = result.Message;
                if (!result.Success)
                {
                    _ = MessageBox.Show(result.Message, LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "手动备份失败");
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                SetBusy(false, null);
                RefreshList();
            }
        }

        /* ###############################  还原  ################################ */

        private async void Btn_Restore_Click(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            List<BackupRow> selected = [.. dg_backups.SelectedItems.OfType<BackupRow>()];
            if (selected.Count != 1)
            {
                _ = MessageBox.Show(LanguageService.Get("Backup_NoSelection"), LanguageService.Get("Cap_Info"));
                return;
            }
            DatabaseBackupEntry target = selected[0].Entry;
            if (!target.CanRestore)
            {
                _ = MessageBox.Show(target.Problem, LanguageService.Get("Cap_Info"));
                return;
            }
            string kind = LanguageService.Get(target.Kind == DatabaseBackupKind.Full ? "Backup_Kind_Full" : "Backup_Kind_Incremental");
            if (MessageBox.Show(
                    string.Format(LanguageService.Get("Backup_RestoreConfirm"), target.CreatedAt.ToString("yyyy/M/d HH:mm:ss"), kind),
                    LanguageService.Get("Backup_RestoreTitle"), MessageBoxButton.YesNo, MessageBoxImage.Warning)
                != MessageBoxResult.Yes)
            {
                return;
            }

            SetBusy(true, LanguageService.Get("Backup_RestoreBusy"));
            DatabaseBackupResult result;
            try
            {
                Progress<string> progress = new(text =>
                {
                    if (txt_status != null)
                    {
                        txt_status.Text = text;
                    }
                });
                result = await Task.Run(() => _backup.Restore(target, progress));
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "还原数据库失败");
                result = DatabaseBackupResult.Fail(ex.Message);
            }
            if (!result.Success)
            {
                SetBusy(false, null);
                _ = MessageBox.Show(result.Message, LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                RefreshList();
                return;
            }
            // 还原成功：数据库连接已释放、库文件已替换，必须重启程序才能继续用
            _ = MessageBox.Show(string.Format(LanguageService.Get("Backup_RestoreDone"), result.Message),
                LanguageService.Get("Backup_RestoreDoneTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
            Application.Current.Shutdown();
        }

        /* ###############################  删除 / 其他  ################################ */

        private void Btn_Delete_Click(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            List<BackupRow> selected = [.. dg_backups.SelectedItems.OfType<BackupRow>()];
            if (selected.Count == 0)
            {
                _ = MessageBox.Show(LanguageService.Get("Backup_NoSelection"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (MessageBox.Show(string.Format(LanguageService.Get("Backup_DeleteConfirm"), selected.Count),
                    LanguageService.Get("Cap_DeleteConfirm"), MessageBoxButton.YesNo, MessageBoxImage.Warning)
                != MessageBoxResult.Yes)
            {
                return;
            }
            int failed = 0;
            foreach (BackupRow row in selected)
            {
                DatabaseBackupResult result = _backup.DeleteBackup(row.Entry);
                if (!result.Success)
                {
                    failed++;
                    _ = MessageBox.Show(result.Message, LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            RefreshList();
            txt_status.Text = failed == 0
                ? string.Format(LanguageService.Get("Backup_DeletedFormat"), selected.Count)
                : string.Format(LanguageService.Get("Backup_DeletedPartialFormat"), selected.Count - failed, failed);
        }

        private void Btn_OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string folder = _backup.ResolveBackupRoot();
                System.IO.Directory.CreateDirectory(folder);
                Process.Start(new ProcessStartInfo { FileName = folder, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                _logger.Warn($"打开备份文件夹失败: {ex.Message}");
                _ = MessageBox.Show(ex.Message, LanguageService.Get("Cap_Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Btn_Refresh_Click(object sender, RoutedEventArgs e) => RefreshList();

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();

        /// <summary>备份/还原期间禁用操作按钮（还原会替换数据库文件，绝不能并发）</summary>
        private void SetBusy(bool busy, string status)
        {
            _busy = busy;
            btn_full.IsEnabled = !busy;
            btn_incremental.IsEnabled = !busy;
            btn_refresh.IsEnabled = !busy;
            btn_delete.IsEnabled = !busy && _isAdmin;
            btn_restore.IsEnabled = !busy && _isAdmin;
            if (status != null)
            {
                txt_status.Text = status;
            }
        }
    }
}
