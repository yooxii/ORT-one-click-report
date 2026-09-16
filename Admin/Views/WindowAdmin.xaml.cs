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
        private readonly MailNotifier _mailNotifier;

        /// <summary>
        /// 语言变更处理（窗口关闭时解除订阅，避免静态事件持有已关闭窗口）
        /// </summary>
        private readonly Action _onLanguageChanged;

        public WindowAdmin()
        {
            InitializeComponent();
            _admin = App.ServiceProvider.GetRequiredService<AdminService>();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _pathService = App.ServiceProvider.GetRequiredService<IPathService>();
            _planExcelService = App.ServiceProvider.GetRequiredService<PlanExcelService>();
            _appSettings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            _mailNotifier = App.ServiceProvider.GetRequiredService<MailNotifier>();

            Loaded += (s, e) =>
            {
                // 历史数据：按测试项目负责人补齐技术员身份（静默执行，仅记日志；用户列表随后刷新）
                _ = _admin.SyncTechniciansFromTestItemOwners();
                LoadUsers();
                LoadCustomers();
                LoadTestItems();
                LoadProducts();
                LoadStages();
            };

            // 语言切换后刷新用户列表中的角色名与按钮文案（XAML 上的 lex:Loc 由本地化引擎自动刷新）
            _onLanguageChanged = () =>
            {
                long? selectedId = SelectedUser?.Id;
                LoadUsers();
                if (selectedId != null)
                {
                    dg_users.SelectedItem = (dg_users.ItemsSource as IEnumerable<UserView>)
                        ?.FirstOrDefault(u => u.Id == selectedId);
                }
                UpdateToggleActiveText();
            };
            LanguageService.LanguageChanged += _onLanguageChanged;
        }

        /// <summary>
        /// 窗口关闭时解除语言变更订阅
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            LanguageService.LanguageChanged -= _onLanguageChanged;
            base.OnClosed(e);
        }

        /* ###############################  人员管理  ################################ */

        private void LoadUsers(long? selectId = null)
        {
            dg_users.ItemsSource = _admin.GetUsers();
            if (selectId != null)
            {
                dg_users.SelectedItem = (dg_users.ItemsSource as IEnumerable<UserView>)
                    ?.FirstOrDefault(u => u.Id == selectId);
            }
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
            txt_username.Text = user.Username;
            txt_email.Text = user.Email;
            chk_general.IsChecked = user.Roles.Contains(UserRole.GeneralUser);
            chk_tech.IsChecked = user.Roles.Contains(UserRole.Technician);
            chk_reviewer.IsChecked = user.Roles.Contains(UserRole.Reviewer);
            chk_admin.IsChecked = user.Roles.Contains(UserRole.Administrator);
            UpdateToggleActiveText();
        }

        /// <summary>
        /// 刷新“禁用/启用该用户”按钮文案（随当前选中用户的启用状态与界面语言变化）
        /// </summary>
        private void UpdateToggleActiveText()
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                return;
            }
            btn_toggleActive.Content = LanguageService.Get(user.IsActive ? "Admin_DisableUser" : "Admin_EnableUser");
        }

        private void Btn_NewUser_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput input = new(LanguageService.Get("Admin_NewUser"),
                (LanguageService.Get("Admin_Username"), "", false),
                (LanguageService.Get("Admin_DisplayNameFull"), "", false),
                (LanguageService.Get("Admin_PasswordMin"), "", true))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _auth.CreateUser(input.Values[0], input.Values[1], input.Values[2], [UserRole.GeneralUser]);
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_NewUserFailed"));
                return;
            }
            LoadUsers();
            _ = MessageBox.Show(LocalizationHelper.Get("Msg_UserCreated"), LanguageService.Get("Cap_Success"));
        }

        private void Btn_ApplyRoles_Click(object sender, RoutedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectUserFirst"), LanguageService.Get("Cap_Info"));
                return;
            }
            List<UserRole> roles = [];
            if (chk_general.IsChecked == true) roles.Add(UserRole.GeneralUser);
            if (chk_tech.IsChecked == true) roles.Add(UserRole.Technician);
            if (chk_reviewer.IsChecked == true) roles.Add(UserRole.Reviewer);
            if (chk_admin.IsChecked == true) roles.Add(UserRole.Administrator);
            if (roles.Count == 0)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectRole"), LanguageService.Get("Cap_Info"));
                return;
            }
            // 邮箱先校验（技术员/审核员登录时据此提示完善），不合法则不保存其它修改
            string emailError = _admin.UpdateEmail(user.Id, txt_email.Text);
            if (emailError != null)
            {
                _ = MessageBox.Show(emailError, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            // 登录名（唯一）
            string usernameError = _admin.UpdateUsername(user.Id, txt_username.Text);
            if (usernameError != null)
            {
                _ = MessageBox.Show(usernameError, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            _admin.UpdateUserRoles(user.Id, roles);
            _admin.UpdateDisplayName(user.Id, txt_displayName.Text?.Trim());
            LoadUsers(user.Id);
            // 显示名可能已变更：同步刷新测试项目负责人栏的显示（负责人按 uid 锚定，显示名实时解析）
            LoadTestItems();
            // 若调整的是当前登录用户自身：角色变了必须重新登录（权限需重新加载），
            // 仅改登录名/显示名/邮箱则保持登录并刷新左下角身份信息
            if (_auth.CurrentUser?.Id == user.Id)
            {
                bool rolesChanged = !roles.OrderBy(r => r).SequenceEqual(user.Roles.OrderBy(r => r));
                if (rolesChanged)
                {
                    _auth.Logout();
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_IdentityChanged"), LanguageService.Get("Cap_Info"));
                }
                else
                {
                    _auth.ReloadCurrentUser();
                }
            }
        }

        private void Btn_ResetPassword_Click(object sender, RoutedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectUserFirst"), LanguageService.Get("Cap_Info"));
                return;
            }
            WindowAdminInput input = new(
                string.Format(LanguageService.Get("Admin_ResetPasswordTitleFormat"), user.Username),
                (LanguageService.Get("Admin_NewPasswordMin"), "", true))
            {
            };
            if (input.ShowDialog() != true)
            {
                return;
            }
            string error = _auth.ResetPassword(user.Id, input.Values[0]);
            _ = MessageBox.Show(error ?? LanguageService.Get("Msg_PasswordReset"),
                LanguageService.Get(error == null ? "Cap_Success" : "Cap_SaveFailed"));
            if (error == null)
            {
                // 密码变更通知（通知类邮件；邮件失败不影响业务）
                _mailNotifier.NotifyPasswordChanged(user.Username, user.DisplayName, _auth.CurrentOperatorName);
            }
        }

        private void Btn_ToggleActive_Click(object sender, RoutedEventArgs e)
        {
            UserView user = SelectedUser;
            if (user == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectUserFirst"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (_auth.CurrentUser?.Id == user.Id)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_CannotDisableSelf"), LanguageService.Get("Cap_Info"));
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
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadCustomers();
        }

        private void Btn_EditCustomer_Click(object sender, RoutedEventArgs e)
        {
            Customer customer = SelectedCustomer;
            if (customer == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectCustomer"), LanguageService.Get("Cap_Info"));
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
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadCustomers();
        }

        private void Btn_DeleteCustomer_Click(object sender, RoutedEventArgs e)
        {
            Customer customer = SelectedCustomer;
            if (customer == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectCustomer"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (MessageBox.Show($"确认删除客户 [{customer.Name}]？", LanguageService.Get("Cap_DeleteConfirm"),
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
            _ = MessageBox.Show($"同步完成，客户+{cAdded}，产品别+{pAdded}，机种映射+{mAdded}", LanguageService.Get("Cap_SyncResult"));
        }

        private void Btn_RefreshCustomers_Click(object sender, RoutedEventArgs e) => LoadCustomers();

        /* ###############################  测试项目管理  ################################ */

        private void LoadTestItems()
        {
            // 没有测试种类时先按历史报告与关键词自动归类（用户手工改过的不动）
            _admin.EnsureTestItemCategories();
            dg_testItems.ItemsSource = _admin.GetTestItems();
        }

        /// <summary>手工触发自动归类：把"不确定"的项目重新判一遍（手工归类过的不动）</summary>
        private void Btn_AutoCategory_Click(object sender, RoutedEventArgs e)
        {
            int changed = _admin.AutoAssignTestItemCategories();
            LoadTestItems();
            _ = MessageBox.Show(string.Format(LanguageService.Get("Admin_AutoCategoryDone"), changed),
                LanguageService.Get("Cap_Info"));
        }

        private TestItemCatalog SelectedTestItem => dg_testItems.SelectedItem as TestItemCatalog;

        /// <summary>
        /// 可选负责人列表：已有技术员
        /// </summary>
        private List<UserView> Technicians => _admin.GetTechnicians();

        private void Btn_NewTestItem_Click(object sender, RoutedEventArgs e)
        {
            WindowTestItemEdit dialog = new(LanguageService.Get("Admin_NewTestItem"), new TestItemCatalog(), Technicians);
            if (dialog.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveTestItem(new TestItemCatalog
            {
                Name = dialog.ItemName,
                Period = dialog.Period,
                Category = dialog.Category,
                Owner = dialog.Owner,
                Remark = dialog.Remark
            });
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadTestItems();
            SyncTechniciansFromOwnersWithMessage();
        }

        private void Btn_EditTestItem_Click(object sender, RoutedEventArgs e)
        {
            TestItemCatalog selected = SelectedTestItem;
            if (selected == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectTestItem"), LanguageService.Get("Cap_Info"));
                return;
            }
            WindowTestItemEdit dialog = new(LanguageService.Get("Admin_EditTestItem"), selected, Technicians);
            if (dialog.ShowDialog() != true)
            {
                return;
            }
            string error = _admin.SaveTestItem(new TestItemCatalog
            {
                Id = selected.Id,
                Name = dialog.ItemName,
                Period = dialog.Period,
                Category = dialog.Category,
                Owner = dialog.Owner,
                Remark = dialog.Remark
            });
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadTestItems();
            SyncTechniciansFromOwnersWithMessage();
        }

        /// <summary>
        /// 按测试项目负责人同步技术员身份（负责人自动获得技术员身份，新账号初始密码 123456）；
        /// 返回提示行（无变化时为空列表）。有变化时刷新用户列表。
        /// </summary>
        private List<string> SyncTechniciansFromOwners()
        {
            (int created, int roleAdded) = _admin.SyncTechniciansFromTestItemOwners();
            if (created == 0 && roleAdded == 0)
            {
                return [];
            }
            LoadUsers();
            List<string> lines = [];
            if (created > 0)
            {
                lines.Add(string.Format(LanguageService.Get("Msg_TechnicianCreatedFormat"), created));
            }
            if (roleAdded > 0)
            {
                lines.Add(string.Format(LanguageService.Get("Msg_TechnicianRoleAddedFormat"), roleAdded));
            }
            return lines;
        }

        /// <summary>
        /// 同步技术员身份并单独提示（测试项目保存后调用）
        /// </summary>
        private void SyncTechniciansFromOwnersWithMessage()
        {
            List<string> lines = SyncTechniciansFromOwners();
            if (lines.Count > 0)
            {
                _ = MessageBox.Show(string.Join("\n", lines), LanguageService.Get("Cap_Info"));
            }
        }

        private void Btn_DeleteTestItem_Click(object sender, RoutedEventArgs e)
        {
            TestItemCatalog item = SelectedTestItem;
            if (item == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectTestItem"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (MessageBox.Show($"确认删除测试项目 [{item.Name}]？", LanguageService.Get("Cap_DeleteConfirm"),
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            _admin.DeleteTestItem(item.Id);
            LoadTestItems();
        }

        private void Btn_SyncTestItems_Click(object sender, RoutedEventArgs e)
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectPlanFile"), initPath: _appSettings.ScheduleDir);
            if (file == null)
            {
                return;
            }
            try
            {
                int added = _admin.SyncTestItemsFromScheduleFile(file);
                LoadTestItems();
                List<string> technicianLines = SyncTechniciansFromOwners();
                technicianLines.Insert(0, $"同步完成，新增 {added} 个测试项目");
                _ = MessageBox.Show(string.Join("\n", technicianLines), LanguageService.Get("Cap_SyncResult"));
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"同步失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void Btn_RefreshTestItems_Click(object sender, RoutedEventArgs e) => LoadTestItems();

        /* ###############################  导入入口（从领退和计划迁入）  ################################ */

        private void Btn_ImportRequisition_Click(object sender, RoutedEventArgs e)
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectReqFile"), initPath: _appSettings.RequisitionDir);
            if (file == null)
            {
                return;
            }
            try
            {
                (int added, int updated) = _planExcelService.ImportRequisition(file);
                _ = MessageBox.Show($"领用表导入完成: 新增{added}条, 更新{updated}条", LanguageService.Get("Cap_ImportResult"));
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"导入领用表失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void Btn_ImportPlan_Click(object sender, RoutedEventArgs e)
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectPlanFile2"), initPath: _appSettings.ScheduleDir);
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
                List<string> technicianLines = SyncTechniciansFromOwners();

                string message = $"计划表导入完成: 新增{added}条, 更新{updated}条\n" +
                    $"字典同步: 客户+{c1}, 产品别+{p1}, 机种映射+{m1}, 测试项目+{t1}";
                if (technicianLines.Count > 0)
                {
                    message += "\n" + string.Join("\n", technicianLines);
                }
                if (unmatched.Count > 0)
                {
                    string list = unmatched.Count > 30
                        ? string.Join("\n", unmatched.GetRange(0, 30)) + $"\n...等共{unmatched.Count}条"
                        : string.Join("\n", unmatched);
                    message += $"\n\n以下 {unmatched.Count} 条备注中未找到工令且工作編號非 Q 开头，未关联到领用数据:\n{list}";
                }
                _ = MessageBox.Show(message, LanguageService.Get("Cap_ImportResult"));
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"导入计划表失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
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
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_OnlyAdminClear"), LanguageService.Get("Cap_InsufficientPermission"));
                return;
            }
            if (MessageBox.Show(LocalizationHelper.Get("Msg_ConfirmClearAll"), LanguageService.Get("Cap_ClearConfirm"), MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            if (MessageBox.Show(LocalizationHelper.Get("Msg_ConfirmClearAgain"), LanguageService.Get("Cap_SecondConfirm"), MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            try
            {
                int n = _planExcelService.ClearAll();
                _ = MessageBox.Show($"已清空 {n} 条记录", LanguageService.Get("Cap_Complete"));
            }
            catch (Exception ex)
            {
                _ = MessageBox.Show($"清空失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
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
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadStages();
        }

        private void Btn_EditStage_Click(object sender, RoutedEventArgs e)
        {
            Stage stage = SelectedStage;
            if (stage == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectStage"), LanguageService.Get("Cap_Info"));
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
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadStages();
        }

        private void Btn_DeleteStage_Click(object sender, RoutedEventArgs e)
        {
            Stage stage = SelectedStage;
            if (stage == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectStage"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (MessageBox.Show($"确认删除阶段 [{stage.Name}]？", LanguageService.Get("Cap_DeleteConfirm"),
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
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadProducts();
        }

        private void Btn_EditProduct_Click(object sender, RoutedEventArgs e)
        {
            Product product = SelectedProduct;
            if (product == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectProductType"), LanguageService.Get("Cap_Info"));
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
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }
            LoadProducts();
        }

        private void Btn_DeleteProduct_Click(object sender, RoutedEventArgs e)
        {
            Product product = SelectedProduct;
            if (product == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectProductType"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (MessageBox.Show($"确认删除产品类型 [{product.Name}]？", LanguageService.Get("Cap_DeleteConfirm"),
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
