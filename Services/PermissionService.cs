using ORT一键报告.Models;
using System.Collections.Generic;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 权限服务接口。操作标识：
    /// plan.view / plan.import / plan.export / plan.edit / plan.delete / plan.review /
    /// report.use / admin.manage / review.view
    /// </summary>
    public interface IPermissionService
    {
        /// <summary>
        /// 当前操作者名称（用于审计字段）
        /// </summary>
        string CurrentUser { get; }

        /// <summary>
        /// 当前角色列表（未登录为空，视为游客）
        /// </summary>
        IReadOnlyList<UserRole> CurrentRoles { get; }

        /// <summary>
        /// 判断当前用户是否可以执行指定操作
        /// </summary>
        bool Can(string action);

        /// <summary>
        /// 计划编辑是否需要提交审核（普通用户：是；技术员及以上：否）
        /// </summary>
        bool PlanEditNeedsReview { get; }
    }

    /// <summary>
    /// 基于角色的权限实现：
    /// 游客：仅浏览计划管理；
    /// 普通用户：浏览+编辑计划（编辑需提交审核）；
    /// 技术员：直接编辑计划 + 一键报告；
    /// 审核员：编辑计划 + 审核计划更改请求；
    /// 管理员：全部权限。
    /// </summary>
    public class PermissionService : IPermissionService
    {
        private readonly AuthService _auth;

        public PermissionService(AuthService auth)
        {
            _auth = auth;
        }

        public string CurrentUser => _auth.CurrentOperatorName;

        public IReadOnlyList<UserRole> CurrentRoles => _auth.CurrentRoles;

        /// <summary>
        /// 计划编辑是否需要提审：仅为普通用户（无技术员/审核员/管理员身份）时需要
        /// </summary>
        public bool PlanEditNeedsReview
            => _auth.HasRole(UserRole.GeneralUser)
            && !_auth.HasRole(UserRole.Technician)
            && !_auth.HasRole(UserRole.Reviewer)
            && !_auth.HasRole(UserRole.Administrator);

        public bool Can(string action)
        {
            bool isTech = _auth.HasRole(UserRole.Technician);
            bool isReviewer = _auth.HasRole(UserRole.Reviewer);
            bool isAdmin = _auth.HasRole(UserRole.Administrator);
            bool isGeneral = _auth.HasRole(UserRole.GeneralUser);

            switch (action)
            {
                // 游客及以上：浏览计划管理
                case "plan.view":
                    return true;

                // 技术员及以上：导入/导出
                case "plan.import":
                case "plan.export":
                    return isTech || isReviewer || isAdmin;

                // 普通用户及以上：编辑/删除（普通用户走审核流）
                case "plan.edit":
                case "plan.delete":
                case "plan.add":
                    return isGeneral || isTech || isReviewer || isAdmin;

                // 审核员及以上：审核请求
                case "plan.review":
                case "review.view":
                    return isReviewer || isAdmin;

                // 技术员与管理员：一键报告（预留：报告提交/审核）
                case "report.use":
                case "report.submit":
                    return isTech || isAdmin;
                case "report.review":
                    return isReviewer || isAdmin;

                // 管理员：管理模块
                case "admin.manage":
                    return isAdmin;

                default:
                    return isAdmin;
            }
        }
    }
}
