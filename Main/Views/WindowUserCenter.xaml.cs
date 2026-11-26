using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 用户中心：管理当前登录用户自己的资料
    /// （显示名、邮箱、密码设置/修改，本机保存的登录信息，以及向审核员/管理员申请更高的身份组）。
    /// 入口：主界面菜单「用户 → 用户中心」。
    /// </summary>
    public partial class WindowUserCenter : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly AuthService _auth;
        private readonly ReviewService _review;
        private readonly AdminService _admin;

        /// <summary>可选的目标身份（普通用户可申请技术员/审核员；技术员只能再申请审核员）</summary>
        private sealed class RoleOption
        {
            public UserRole Role { get; set; }
            public string Display { get; set; }
        }

        /// <summary>可选审批人（审核员与管理员；审核员身份组的申请只列管理员）</summary>
        private sealed class ApproverOption
        {
            public string Username { get; set; }
            public string Display { get; set; }
        }

        public WindowUserCenter()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _review = App.ServiceProvider.GetRequiredService<ReviewService>();
            _admin = App.ServiceProvider.GetRequiredService<AdminService>();
            Loaded += (s, e) => LoadUser();
        }

        /* ###############################  载入  ################################ */

        /// <summary>
        /// 载入当前登录用户的信息；未登录时提示并关闭窗口
        /// </summary>
        private void LoadUser()
        {
            User user = _auth.CurrentUser;
            if (user == null)
            {
                _ = MessageBox.Show(LanguageService.Get("UserCenter_Msg_NeedLogin"), LanguageService.Get("Cap_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                Close();
                return;
            }
            txt_username.Text = user.Username;
            txt_roles.Text = _auth.CurrentRoles.Count == 0
                ? LanguageService.Get("Role_Guest")
                : string.Join("/", _auth.CurrentRoles.Select(r => LanguageService.Get("Role_" + r)));
            txt_createdAt.Text = user.CreatedAt.ToString("yyyy/M/d HH:mm");
            txt_displayName.Text = user.DisplayName ?? "";
            txt_email.Text = user.Email ?? "";
            txt_currentPassword.Clear();
            txt_newPassword.Clear();
            txt_confirmPassword.Clear();

            // 没设置过密码时不需要"当前密码"，并且把提示改成"请设置密码"
            bool hasPassword = AuthService.HasPassword(user);
            Visibility currentVisibility = hasPassword ? Visibility.Visible : Visibility.Collapsed;
            lbl_currentPassword.Visibility = currentVisibility;
            txt_currentPassword.Visibility = currentVisibility;
            txt_passwordHint.Text = hasPassword
                ? LanguageService.Get("UserCenter_PasswordKeepHint")
                : LanguageService.Get("UserCenter_NoPasswordHint");

            UpdateCookieHint();
            LoadRoleApply();
            if (hasPassword)
            {
                txt_displayName.Focus();
            }
            else
            {
                txt_newPassword.Focus();
            }
        }

        /// <summary>刷新"本机登录信息"（cookie 到期时间）</summary>
        private void UpdateCookieHint()
        {
            DateTime? expiry = _auth.GetCookieExpiry();
            txt_cookie.Text = expiry == null
                ? LanguageService.Get("UserCenter_NoCookie")
                : string.Format(LanguageService.Get("UserCenter_CookieExpiryFormat"), expiry.Value.ToString("yyyy/M/d HH:mm"));
        }

        /* ###############################  事件函数  ################################ */

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            if (_auth.CurrentUser == null)
            {
                return;
            }
            bool hasPassword = AuthService.HasPassword(_auth.CurrentUser);
            string current = txt_currentPassword.Password;
            string newPassword = txt_newPassword.Password;
            string confirm = txt_confirmPassword.Password;

            // 1. 想改/设密码时先校验
            bool wantPasswordChange = !string.IsNullOrEmpty(newPassword) || !string.IsNullOrEmpty(confirm);
            if (hasPassword && wantPasswordChange && !_auth.VerifyCurrentUserPassword(current))
            {
                _ = MessageBox.Show(LanguageService.Get("Msg_CurrentPasswordWrong"), LanguageService.Get("Cap_Info"));
                txt_currentPassword.Focus();
                return;
            }
            if (wantPasswordChange)
            {
                if (newPassword != confirm)
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_PasswordMismatch"), LanguageService.Get("Cap_Info"));
                    txt_confirmPassword.Focus();
                    return;
                }
                if (newPassword.Length < 6)
                {
                    _ = MessageBox.Show(LanguageService.Get("Msg_PasswordTooShort"), LanguageService.Get("Cap_Info"));
                    txt_newPassword.Focus();
                    return;
                }
            }

            // 2. 基本资料（显示名 + 邮箱）
            string error = _auth.SaveCurrentUserProfile(txt_displayName.Text, txt_email.Text);
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                return;
            }

            // 3. 密码
            if (wantPasswordChange)
            {
                error = _auth.SetCurrentUserPassword(newPassword);
                if (error != null)
                {
                    _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"));
                    return;
                }
            }

            _logger.Info($"用户中心：已保存 {_auth.CurrentUser.Username} 的资料");
            ToastService.Show(LanguageService.Get("UserCenter_Saved"), ToastType.Info);
            LoadUser();
        }

        private void Btn_ClearCookie_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(LanguageService.Get("UserCenter_ClearCookieConfirm"), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            _auth.ClearLoginCookie();
            UpdateCookieHint();
            ToastService.Show(LanguageService.Get("UserCenter_CookieCleared"), ToastType.Info);
        }

        /* ###############################  身份组申请  ################################ */

        /// <summary>
        /// 载入「身份组申请」：已经是管理员（最高身份）就隐藏这一块；
        /// 有待审批的申请时只显示进度，不允许重复提交。
        /// </summary>
        private void LoadRoleApply()
        {
            User user = _auth.CurrentUser;
            if (user == null || _auth.HasRole(UserRole.Administrator))
            {
                grp_roleApply.Visibility = Visibility.Collapsed;
                return;
            }
            List<RoleOption> targets = [];
            if (!_auth.HasRole(UserRole.Technician))
            {
                targets.Add(new RoleOption { Role = UserRole.Technician, Display = LanguageService.Get("Role_Technician") });
            }
            if (!_auth.HasRole(UserRole.Reviewer))
            {
                targets.Add(new RoleOption { Role = UserRole.Reviewer, Display = LanguageService.Get("Role_Reviewer") });
            }
            grp_roleApply.Visibility = targets.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            if (targets.Count == 0)
            {
                return;
            }
            cb_roleTarget.ItemsSource = targets;
            cb_roleTarget.SelectedIndex = 0;

            ReviewRequest pending = _review.GetPendingRoleRequest(user.Id);
            bool hasPending = pending != null;
            cb_roleTarget.IsEnabled = !hasPending;
            cb_roleApprover.IsEnabled = !hasPending;
            txt_roleNote.IsEnabled = !hasPending;
            if (hasPending)
            {
                btn_applyRole.IsEnabled = false;
                txt_roleState.Text = string.Format(LanguageService.Get("UserCenter_RoleApplyPendingFormat"),
                    pending.Action, pending.AssigneeName ?? "-");
                return;
            }
            txt_roleState.Text = LanguageService.Get("UserCenter_RoleApplyHint");
            BuildApprovers();
        }

        /// <summary>按目标身份重建审批人列表：技术员申请可交审核员或管理员，审核员申请只能交管理员</summary>
        private void BuildApprovers()
        {
            if (cb_roleTarget.SelectedItem is not RoleOption target)
            {
                return;
            }
            List<ApproverOption> options = [];
            foreach (UserView user in _admin.GetUsers())
            {
                if (!user.IsActive || (_auth.CurrentUser != null && user.Id == _auth.CurrentUser.Id))
                {
                    continue;
                }
                bool isAdmin = user.Roles.Contains(UserRole.Administrator);
                bool isReviewer = user.Roles.Contains(UserRole.Reviewer);
                if (!isAdmin && !isReviewer)
                {
                    continue;
                }
                if (target.Role == UserRole.Reviewer && !isAdmin)
                {
                    continue;   // 审核员身份组的申请只有管理员能批
                }
                options.Add(new ApproverOption
                {
                    Username = user.Username,
                    Display = $"{user.DisplayName ?? user.Username}（{user.RolesText}）"
                });
            }
            cb_roleApprover.ItemsSource = options;
            cb_roleApprover.SelectedIndex = options.Count > 0 ? 0 : -1;
            btn_applyRole.IsEnabled = options.Count > 0;
            if (options.Count == 0)
            {
                txt_roleState.Text = LanguageService.Get("UserCenter_RoleApplyNoApprover");
            }
        }

        /// <summary>换目标身份 → 审批人范围跟着变（审核员申请只能选管理员）</summary>
        private void Cb_RoleTarget_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            User user = _auth.CurrentUser;
            if (user == null || _auth.HasRole(UserRole.Administrator))
            {
                return;
            }
            if (_review.GetPendingRoleRequest(user.Id) != null)
            {
                return;
            }
            BuildApprovers();
        }

        /// <summary>提交身份组申请：生成一条待审核请求，交给用户指定的审核员或管理员</summary>
        private void Btn_ApplyRole_Click(object sender, RoutedEventArgs e)
        {
            User user = _auth.CurrentUser;
            if (user == null)
            {
                _ = MessageBox.Show(LanguageService.Get("UserCenter_Msg_NeedLogin"), LanguageService.Get("Cap_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (cb_roleTarget.SelectedItem is not RoleOption target)
            {
                return;
            }
            if (cb_roleApprover.SelectedItem is not ApproverOption approver)
            {
                _ = MessageBox.Show(LanguageService.Get("UserCenter_RoleApplyNoApprover"),
                    LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (MessageBox.Show(string.Format(LanguageService.Get("UserCenter_RoleApplyConfirm"),
                    target.Display, approver.Display), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            string error = _review.SubmitRoleRequest(user.Id, user.Username, target.Role, approver.Username,
                txt_roleNote.Text);
            if (error != null)
            {
                _ = MessageBox.Show(error, LanguageService.Get("Cap_SaveFailed"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            _logger.Info($"用户中心：{user.Username} 提交了 {target.Display} 身份组申请");
            ToastService.Show(LanguageService.Get("UserCenter_RoleApplySubmitted"), ToastType.Info);
            LoadRoleApply();
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
