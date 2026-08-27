using Microsoft.Extensions.DependencyInjection;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
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
        private readonly PlanExcelService _planExcelService;
        private readonly AppSettingsService _appSettings;

        public WindowAdmin()
        {
            InitializeComponent();
            _admin = App.ServiceProvider.GetRequiredService<AdminService>();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _pathService = App.ServiceProvider.GetRequiredService<IPathService>();
            _planExcelService = App.ServiceProvider.GetRequiredService<PlanExcelService>();
            _appSettings = App.ServiceProvider.GetRequiredService<AppSettingsService>();

            Loaded += (s, e) =>
            {
                LoadUsers();
                LoadCustomers();
                LoadTestItems();
                LoadProducts();
                LoadStages();
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

        /* ###############################  客户管理（整合产品别）  ################################ */

        private void LoadCustomers()
        {
            dg_customers.ItemsSource = _admin.GetCustomers()
                .OrderBy(c => string.IsNullOrWhiteSpace(c.Code) ? "ZZZ" : c.Code)
                .ToList();
        }

        private Customer SelectedCustomer => dg_customers.SelectedItem as Customer;

        private void Btn_NewCustomer_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput input = new("新增客户",
                ("客户名称", "", false),
                ("客户代码", "", false),
                ("备注", "", false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveCustomer(new Customer
            {
                Name = input.Values[0],
                Code = input.Values[1],
                Remark = input.Values[2]
            });
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
            WindowAdminInput input = new("编辑客户",
                ("客户名称", customer.Name, false),
                ("客户代码", customer.Code, false),
                ("备注", customer.Remark, false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            customer.Name = input.Values[0];
            customer.Code = input.Values[1];
            customer.Remark = input.Values[2];
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
            (int cAdded, int pAdded, int mAdded) = _admin.SyncCatalogsFromPlans();
            LoadCustomers();
            LoadProducts();
            _ = MessageBox.Show($"同步完成，客户+{cAdded}，产品别+{pAdded}，机种映射+{mAdded}", "同步结果");
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
            WindowAdminInput input = new("新增测试项目",
                ("试验项目", "", false),
                ("试验时间", "", false),
                ("负责人", "", false),
                ("备注", "", false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveTestItem(new TestItemCatalog
            {
                Name = input.Values[0],
                Period = input.Values[1],
                Owner = input.Values[2],
                Remark = input.Values[3]
            });
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
            WindowAdminInput input = new("编辑测试项目",
                ("试验项目", selected.Name, false),
                ("试验时间", selected.Period, false),
                ("负责人", selected.Owner, false),
                ("备注", selected.Remark, false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            selected.Name = input.Values[0];
            selected.Period = input.Values[1];
            selected.Owner = input.Values[2];
            selected.Remark = input.Values[3];
            string error = _admin.SaveTestItem(selected);
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadTestItems();
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
            string file = _pathService.OpenPathDialog("选择计划表文件(ORT Test Schedule)", initPath: _appSettings.ScheduleDir);
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
            catch (Exception ex)
            {
                _ = MessageBox.Show($"同步失败:\n{ex.Message}", "错误");
            }
        }

        private void Btn_RefreshTestItems_Click(object sender, RoutedEventArgs e) => LoadTestItems();

        /* ###############################  导入入口（从领退和计划迁入）  ################################ */

        private void Btn_ImportRequisition_Click(object sender, RoutedEventArgs e)
        {
            string file = _pathService.OpenPathDialog("选择领用表(成品領用記錄)", initPath: _appSettings.RequisitionDir);
            if (file == null)
            {
                return;
            }
            try
            {
                (int added, int updated) = _planExcelService.ImportRequisition(file);
                _ = MessageBox.Show($"领用表导入完成: 新增{added}条, 更新{updated}条", "导入结果");
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"导入领用表失败:\n{ex.Message}", "错误");
            }
        }

        private void Btn_ImportPlan_Click(object sender, RoutedEventArgs e)
        {
            string file = _pathService.OpenPathDialog("选择计划表(ORT Test Schedule)", initPath: _appSettings.ScheduleDir);
            if (file == null)
            {
                return;
            }
            try
            {
                (int added, int updated, List<string> unmatched) = _planExcelService.ImportSchedule(file);
                (int c1, int p1, int m1) = _admin.SyncCatalogsFromScheduleFile(file);
                int t1 = _admin.SyncTestItemsFromScheduleFile(file);
                LoadCustomers();
                LoadTestItems();
                LoadProducts();

                string message = $"计划表导入完成: 新增{added}条, 更新{updated}条\n" +
                    $"字典同步: 客户+{c1}, 产品别+{p1}, 机种映射+{m1}, 测试项目+{t1}";
                if (unmatched.Count > 0)
                {
                    string list = unmatched.Count > 30
                        ? string.Join("\n", unmatched.GetRange(0, 30)) + $"\n...等共{unmatched.Count}条"
                        : string.Join("\n", unmatched);
                    message += $"\n\n以下 {unmatched.Count} 条备注中未找到工令且工作編號非 Q 开头，未关联到领用数据:\n{list}";
                }
                _ = MessageBox.Show(message, "导入结果");
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"导入计划表失败:\n{ex.Message}", "错误");
            }
        }

        /// <summary>
        /// 清空全部计划数据（领退表+计划表），仅管理员可操作，二次确认
        /// </summary>
        private void Btn_ClearAll_Click(object sender, RoutedEventArgs e)
        {
            IPermissionService permission = App.ServiceProvider.GetRequiredService<IPermissionService>();
            if (!permission.Can("admin.manage"))
            {
                _ = MessageBox.Show("只有管理员可以清空计划数据", "权限不足");
                return;
            }
            if (MessageBox.Show("确认清空领退表和计划表的全部数据？此操作不可恢复！\n建议操作前先导出备份。",
                "清空确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            if (MessageBox.Show("再次确认：真的要清空全部计划数据吗？",
                "二次确认", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            try
            {
                int n = _planExcelService.ClearAll();
                _ = MessageBox.Show($"已清空 {n} 条记录", "完成");
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"清空失败:\n{ex.Message}", "错误");
            }
        }

        /* ###############################  阶段管理  ################################ */

        private void LoadStages()
        {
            dg_stages.ItemsSource = _admin.GetStages();
        }

        private Stage SelectedStage => dg_stages.SelectedItem as Stage;

        private void Btn_NewStage_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput input = new("新增阶段", ("阶段名", "", false), ("描述", "", false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveStage(new Stage { Name = input.Values[0], Description = input.Values[1] });
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadStages();
        }

        private void Btn_EditStage_Click(object sender, RoutedEventArgs e)
        {
            Stage stage = SelectedStage;
            if (stage == null)
            {
                _ = MessageBox.Show("请先选择一个阶段", "提示");
                return;
            }
            WindowAdminInput input = new("编辑阶段", ("阶段名", stage.Name, false), ("描述", stage.Description, false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            stage.Name = input.Values[0];
            stage.Description = input.Values[1];
            string error = _admin.SaveStage(stage);
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadStages();
        }

        private void Btn_DeleteStage_Click(object sender, RoutedEventArgs e)
        {
            Stage stage = SelectedStage;
            if (stage == null)
            {
                _ = MessageBox.Show("请先选择一个阶段", "提示");
                return;
            }
            if (MessageBox.Show($"确认删除阶段 [{stage.Name}]？", "删除确认",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            _admin.DeleteStage(stage.Id);
            LoadStages();
        }

        private void Btn_RefreshStages_Click(object sender, RoutedEventArgs e) => LoadStages();

        /* ###############################  产品别管理  ################################ */

        private void LoadProducts()
        {
            dg_products.ItemsSource = _admin.GetProductEntities()
                .OrderBy(p => string.IsNullOrWhiteSpace(p.Code) ? "ZZZ" : p.Code)
                .ToList();
        }

        private Product SelectedProduct => dg_products.SelectedItem as Product;

        private void Btn_NewProduct_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput input = new("新增产品类型",
                ("产品类型名称", "", false),
                ("产品代码", "", false),
                ("备注", "", false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveProduct(new Product
            {
                Name = input.Values[0],
                Code = input.Values[1],
                Remark = input.Values[2]
            });
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadProducts();
        }

        private void Btn_EditProduct_Click(object sender, RoutedEventArgs e)
        {
            Product product = SelectedProduct;
            if (product == null)
            {
                _ = MessageBox.Show("请先选择一个产品类型", "提示");
                return;
            }
            WindowAdminInput input = new("编辑产品类型",
                ("产品类型名称", product.Name, false),
                ("产品代码", product.Code, false),
                ("备注", product.Remark, false))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            product.Name = input.Values[0];
            product.Code = input.Values[1];
            product.Remark = input.Values[2];
            string error = _admin.SaveProduct(product);
            if (error != null)
            {
                _ = MessageBox.Show(error, "保存失败");
                return;
            }
            LoadProducts();
        }

        private void Btn_DeleteProduct_Click(object sender, RoutedEventArgs e)
        {
            Product product = SelectedProduct;
            if (product == null)
            {
                _ = MessageBox.Show("请先选择一个产品类型", "提示");
                return;
            }
            if (MessageBox.Show($"确认删除产品类型 [{product.Name}]？", "删除确认",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            _admin.DeleteProduct(product.Id);
            LoadProducts();
        }

        private void Btn_RefreshProducts_Click(object sender, RoutedEventArgs e) => LoadProducts();
    }
}
