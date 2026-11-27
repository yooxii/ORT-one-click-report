using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// 用户中心 →「申请权限」二级窗口：向审核员或管理员申请更高的身份组。
    /// 管理员身份只能由管理员直接指定，不能在这里申请；
    /// 已有待审批申请时只显示进度，不允许重复提交。
    /// </summary>
    public partial class WindowRoleRequest : Window
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

        public WindowRoleRequest()
        {
            InitializeComponent();
            _auth = App.ServiceProvider.GetRequiredService<AuthService>();
            _review = App.ServiceProvider.GetRequiredService<ReviewService>();
            _admin = App.ServiceProvider.GetRequiredService<AdminService>();
            Loaded += (_, _) => LoadRoleApply();
        }

        private void LoadRoleApply()
        {
            User user = _auth.CurrentUser;
            if (user == null)
            {
                Close();
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
            if (targets.Count == 0)
            {
                txt_roleState.Text = LanguageService.Get("UserCenter_RoleApplyAllGranted");
                cb_roleTarget.IsEnabled = false;
                cb_roleApprover.IsEnabled = false;
                txt_roleNote.IsEnabled = false;
                btn_applyRole.IsEnabled = false;
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

        private void Btn_ApplyRole_Click(object sender, RoutedEventArgs e)
        {
            User user = _auth.CurrentUser;
            if (user == null)
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_Msg_NeedLogin"), LanguageService.Get("Cap_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (cb_roleTarget.SelectedItem is not RoleOption target)
            {
                return;
            }
            if (cb_roleApprover.SelectedItem is not ApproverOption approver)
            {
                MessageBox.Show(this, LanguageService.Get("UserCenter_RoleApplyNoApprover"),
                    LanguageService.Get("Cap_Info"), MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (MessageBox.Show(this, string.Format(LanguageService.Get("UserCenter_RoleApplyConfirm"),
                    target.Display, approver.Display), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            string error = _review.SubmitRoleRequest(user.Id, user.Username, target.Role, approver.Username,
                txt_roleNote.Text);
            if (error != null)
            {
                MessageBox.Show(this, error, LanguageService.Get("Cap_SaveFailed"),
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
