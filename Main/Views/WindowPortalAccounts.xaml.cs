using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 「部门账号」窗口：登记哪些公司平台账号属于本部门、各自的身份组是什么。
    ///
    /// 平台登录成功只说明「你是本公司的人」，还要靠这张表判断「是不是我们部门、进来算哪个身份」：
    /// - 表里有的账号：按这里的身份组登录（停用即禁止登录）；
    /// - 表里没有的账号：登录时如果区域属于本部门，会自动登记成设置里指定的默认身份（可在设置里关掉）；
    ///   区域不属于本部门的直接拒绝，并提示来找管理员。
    /// </summary>
    public partial class WindowPortalAccounts : Window
    {
        private readonly PortalAccountService _portal;

        /// <summary>界面上的账号行（改完点保存才写库）</summary>
        private readonly ObservableCollection<PortalAccount> _rows = [];

        /// <summary>本次被删掉的登记（保存时真正删除）</summary>
        private readonly List<long> _deletedIds = [];

        public WindowPortalAccounts()
        {
            InitializeComponent();
            _portal = App.ServiceProvider.GetRequiredService<PortalAccountService>();
            // 身份组下拉列不在可视树里，只能按列取到再设 ItemsSource
            DataGridComboBoxColumn roleColumn = dg_accounts.Columns
                .OfType<DataGridComboBoxColumn>().FirstOrDefault();
            if (roleColumn != null)
            {
                roleColumn.ItemsSource = RoleOptions();
            }
            dg_accounts.ItemsSource = _rows;
            Load();
        }

        /// <summary>身份组下拉项（就是本程序原来的身份组，不新增类型）</summary>
        private static List<RoleOption> RoleOptions() =>
        [
            new RoleOption(nameof(UserRole.GeneralUser), LanguageService.Get("Role_GeneralUser")),
            new RoleOption(nameof(UserRole.Technician), LanguageService.Get("Role_Technician")),
            new RoleOption(nameof(UserRole.Reviewer), LanguageService.Get("Role_Reviewer")),
            new RoleOption(nameof(UserRole.Administrator), LanguageService.Get("Role_Administrator"))
        ];

        private void Load()
        {
            _rows.Clear();
            _deletedIds.Clear();
            try
            {
                foreach (PortalAccount account in _portal.ListAccounts())
                {
                    if (string.IsNullOrWhiteSpace(account.Role))
                    {
                        account.Role = nameof(UserRole.Technician);
                    }
                    _rows.Add(account);
                }
                txt_status.Text = _rows.Count == 0
                    ? LanguageService.Get("PortalAccounts_Empty")
                    : string.Format(LanguageService.Get("PortalAccounts_CountFormat"), _rows.Count);
            }
            catch (Exception ex)
            {
                txt_status.Text = "读取部门账号失败：" + ex.Message;
            }
        }

        private void Btn_Add_Click(object sender, RoutedEventArgs e)
        {
            PortalAccount row = new()
            {
                Account = "",
                Role = nameof(UserRole.Technician),
                IsActive = true
            };
            _rows.Add(row);
            dg_accounts.SelectedItem = row;
            dg_accounts.ScrollIntoView(row);
            // 新行直接进入编辑状态，省一次双击
            dg_accounts.BeginEdit();
        }

        private void Btn_Delete_Click(object sender, RoutedEventArgs e)
        {
            List<PortalAccount> selected = dg_accounts.SelectedItems.Cast<PortalAccount>().ToList();
            if (selected.Count == 0)
            {
                txt_status.Text = LanguageService.Get("PortalAccounts_SelectFirst");
                return;
            }
            string names = string.Join("、", selected.Select(a => string.IsNullOrWhiteSpace(a.Account) ? "(新行)" : a.Account));
            if (MessageBox.Show(string.Format(LanguageService.Get("PortalAccounts_DeleteConfirm"), names),
                LanguageService.Get("PortalAccounts_Title"), MessageBoxButton.YesNo, MessageBoxImage.Question)
                != MessageBoxResult.Yes)
            {
                return;
            }
            foreach (PortalAccount row in selected)
            {
                if (row.Id > 0)
                {
                    _deletedIds.Add(row.Id);
                }
                _rows.Remove(row);
            }
            txt_status.Text = string.Format(LanguageService.Get("PortalAccounts_DeletedFormat"), selected.Count);
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            // 先在界面上把明显的错误挡掉（空账号、重号），避免写一半失败
            List<string> seen = [];
            foreach (PortalAccount row in _rows)
            {
                row.Account = (row.Account ?? "").Trim();
                if (string.IsNullOrEmpty(row.Account))
                {
                    txt_status.Text = LanguageService.Get("PortalAccounts_AccountRequired");
                    dg_accounts.SelectedItem = row;
                    return;
                }
                if (seen.Any(a => string.Equals(a, row.Account, StringComparison.OrdinalIgnoreCase)))
                {
                    txt_status.Text = string.Format(LanguageService.Get("PortalAccounts_DuplicateFormat"), row.Account);
                    dg_accounts.SelectedItem = row;
                    return;
                }
                seen.Add(row.Account);
            }

            try
            {
                foreach (long id in _deletedIds.Distinct())
                {
                    _portal.DeleteAccount(id);
                }
                _deletedIds.Clear();

                List<string> errors = [];
                foreach (PortalAccount row in _rows)
                {
                    string error = _portal.SaveAccount(row);
                    if (!string.IsNullOrWhiteSpace(error))
                    {
                        errors.Add($"{row.Account}：{error}");
                    }
                }
                if (errors.Count > 0)
                {
                    txt_status.Text = string.Join("；", errors);
                    MessageBox.Show(string.Join("\n", errors), LanguageService.Get("PortalAccounts_Title"),
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                Load();
                txt_status.Text = LanguageService.Get("PortalAccounts_Saved");
            }
            catch (Exception ex)
            {
                txt_status.Text = "保存失败：" + ex.Message;
                MessageBox.Show(ex.Message, LanguageService.Get("PortalAccounts_Title"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }

    /// <summary>身份组下拉项</summary>
    public class RoleOption
    {
        public RoleOption(string code, string name)
        {
            Code = code;
            Name = name;
        }

        /// <summary>UserRole 名称（落库用）</summary>
        public string Code { get; }

        /// <summary>界面显示名（本地化）</summary>
        public string Name { get; }
    }
}
