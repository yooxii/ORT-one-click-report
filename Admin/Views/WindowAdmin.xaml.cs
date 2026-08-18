using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Admin.Views
{
    /// <summary>
    /// WindowAdmin.xaml 的交互逻辑：人员管理 / 客户管理 / 测试项目管理（仅管理员进入）
    /// </summary>
    public partial class WindowAdmin : Window
    {
        private readonly AdminService _admin;
        private readonly AuthService _auth;
        private readonly IPathService _pathService;

        public WindowAdmin()
        {
            InitializeComponent();
            _admin = App.ServiceProvider.GetRequiredService<AdminService>();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _pathService = App.ServiceProvider.GetRequiredService<IPathService>();

            Loaded += (s, e) =>
            {
                LoadUsers();
                LoadCustomers();
                LoadTestItems();
            };
        }

        /* ###############################  人员管理  ################################ */

        private void LoadUsers()
        {
            dg_users.ItemsSource = _admin.GetUsers();
        }

        private UserView SelectedUser => dg_users.SelectedItem as UserView;

        private void Dg_Users_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                return;
            }
            txt_displayName.Text = user.DisplayName;
            chk_general.IsChecked = user.Roles.Contains(UserRole.GeneralUser);
            chk_tech.IsChecked = user.Roles.Contains(UserRole.Technician);
            chk_reviewer.IsChecked = user.Roles.Contains(UserRole.Reviewer);
            chk_admin.IsChecked = user.Roles.Contains(UserRole.Administrator);
            btn_toggleActive.Content = user.IsActive ? "禁用该用户" : "启用该用户";
        }

        private void Btn_NewUser_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput input = new("新建用户",
                ("用户名", "", false),
                ("显示名称", "", false),
                ("密码（至少6位）", "", true))
            {
                Owner = this
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _auth.CreateUser(input.Values[0], input.Values[1], input.Values[2], [UserRole.GeneralUser]);
            if (error != null)
            {
                _ = MessageBox.Show(error, "新建用户失败");
                return;
            }
            LoadUsers();
            _ = MessageBox.Show("用户已创建（默认角色：普通用户），可在右侧调整其身份。", "成功");
        }

        private void Btn_ApplyRoles_Click(object sender, RoutedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                _ = MessageBox.Show("请先在左侧选择一个用户", "提示");
                return;
            }
            List<UserRole> roles = [];
            if (chk_general.IsChecked == true) roles.Add(UserRole.GeneralUser);
            if (chk_tech.IsChecked == true) roles.Add(UserRole.Technician);
            if (chk_reviewer.IsChecked == true) roles.Add(UserRole.Reviewer);
            if (chk_admin.IsChecked == true) roles.Add(UserRole.Administrator);
            if (roles.Count == 0)
            {
                _ = MessageBox.Show("请至少勾选一个角色", "提示");
                return;
            }
            _admin.UpdateUserRoles(user.Id, roles);
            _admin.UpdateDisplayName(user.Id, txt_displayName.Text?.Trim());
            LoadUsers();
            // 若调整的是当前登录用户自身，触发权限刷新
            if (_auth.CurrentUser?.Id == user.Id)
            {
                _auth.Logout();
                _ = MessageBox.Show("调整了当前登录用户的身份，已注销，请重新登录。", "提示");
            }
        }

        private void Btn_ResetPassword_Click(object sender, RoutedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                _ = MessageBox.Show("请先在左侧选择一个用户", "提示");
                return;
            }
            WindowAdminInput input = new($"重置密码 - {user.Username}", ("新密码（至少6位）", "", true))
            {
                Owner = this
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _auth.ResetPassword(user.Id, input.Values[0]);
            _ = MessageBox.Show(error ?? "密码已重置", error == null ? "成功" : "失败");
        }

        private void Btn_ToggleActive_Click(object sender, RoutedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                _ = MessageBox.Show("请先在左侧选择一个用户", "提示");
                return;
            }
            if (_auth.CurrentUser?.Id == user.Id)
            {
                _ = MessageBox.Show("不能禁用当前登录用户", "提示");
                return;
            }
            _admin.SetUserActive(user.Id, !user.IsActive);
            LoadUsers();
        }

        private void Btn_RefreshUsers_Click(object sender, RoutedEventArgs e) => LoadUsers();

        /* ###############################  客户管理  ################################ */

        private void LoadCustomers()
        {
            dg_customers.ItemsSource = _admin.GetCustomers();
        }

        private Customer SelectedCustomer => dg_customers.SelectedItem as Customer;

        private void Btn_NewCustomer_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput input = new("新增客户", ("客户名称", "", false), ("备注", "", false))
            {
                Owner = this
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveCustomer(new Customer { Name = input.Values[0], Remark = input.Values[1] });
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadCustomers();
        }

        private void Btn_EditCustomer_Click(object sender, RoutedEventArgs e)
        {
            Customer customer = SelectedCustomer;
            if (customer == null)
            {
                _ = MessageBox.Show("请先选择一个客户", "提示");
                return;
            }
            WindowAdminInput input = new("编辑客户", ("客户名称", customer.Name, false), ("备注", customer.Remark, false))
            {
                Owner = this
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            customer.Name = input.Values[0];
            customer.Remark = input.Values[1];
            string error = _admin.SaveCustomer(customer);
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadCustomers();
        }

        private void Btn_DeleteCustomer_Click(object sender, RoutedEventArgs e)
        {
            Customer customer = SelectedCustomer;
            if (customer == null)
            {
                _ = MessageBox.Show("请先选择一个客户", "提示");
                return;
            }
            if (MessageBox.Show($"确认删除客户 [{customer.Name}]？", "删除确认",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            _admin.DeleteCustomer(customer.Id);
            LoadCustomers();
        }

        private void Btn_SyncCustomers_Click(object sender, RoutedEventArgs e)
        {
            int added = _admin.SyncCustomersFromPlans();
            LoadCustomers();
            _ = MessageBox.Show($"同步完成，新增 {added} 个客户", "同步结果");
        }

        private void Btn_RefreshCustomers_Click(object sender, RoutedEventArgs e) => LoadCustomers();

        /* ###############################  测试项目管理  ################################ */

        private void LoadTestItems()
        {
            dg_testItems.ItemsSource = _admin.GetTestItems();
        }

        private TestItemCatalog SelectedTestItem => dg_testItems.SelectedItem as TestItemCatalog;

        private void Btn_NewTestItem_Click(object sender, RoutedEventArgs e)
        {
            TestItemCatalog item = InputTestItem("新增测试项目", new TestItemCatalog());
            if (item == null)
            {
                return;
            }
            string error = _admin.SaveTestItem(item);
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadTestItems();
        }

        private void Btn_EditTestItem_Click(object sender, RoutedEventArgs e)
        {
            TestItemCatalog selected = SelectedTestItem;
            if (selected == null)
            {
                _ = MessageBox.Show("请先选择一个测试项目", "提示");
                return;
            }
            TestItemCatalog item = InputTestItem("编辑测试项目", selected);
            if (item == null)
            {
                return;
            }
            string error = _admin.SaveTestItem(item);
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadTestItems();
        }

        private TestItemCatalog InputTestItem(string title, TestItemCatalog source)
        {
            WindowAdminInput input = new(title,
                ("试验项目", source.Name, false),
                ("试验时间", source.Period, false),
                ("负责人", source.Owner, false),
                ("备注", source.Remark, false))
            {
                Owner = this
            };
            if (input.ShowDialog() != true)
            {
                return null;
            }
            return new TestItemCatalog
            {
                Id = source.Id,
                Name = input.Values[0],
                Period = input.Values[1],
                Owner = input.Values[2],
                Remark = input.Values[3]
            };
        }

        private void Btn_DeleteTestItem_Click(object sender, RoutedEventArgs e)
        {
            TestItemCatalog item = SelectedTestItem;
            if (item == null)
            {
                _ = MessageBox.Show("请先选择一个测试项目", "提示");
                return;
            }
            if (MessageBox.Show($"确认删除测试项目 [{item.Name}]？", "删除确认",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            _admin.DeleteTestItem(item.Id);
            LoadTestItems();
        }

        private void Btn_SyncTestItems_Click(object sender, RoutedEventArgs e)
        {
            string file = _pathService.OpenPathDialog("选择计划表文件(ORT Test Schedule)");
            if (file == null)
            {
                return;
            }
            try
            {
                int added = _admin.SyncTestItemsFromScheduleFile(file);
                LoadTestItems();
                _ = MessageBox.Show($"同步完成，新增 {added} 个测试项目", "同步结果");
            }
            catch (System.Exception ex)
            {
                _ = MessageBox.Show($"同步失败:\n{ex.Message}", "错误");
            }
        }

        private void Btn_RefreshTestItems_Click(object sender, RoutedEventArgs e) => LoadTestItems();
    }
}
