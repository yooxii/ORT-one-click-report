using ORT一键报告.Models;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 身份组申请的规则（纯函数，界面与服务共用，便于单元测试）：
    /// 1) 只能申请「技术员」「审核员」；管理员身份只能由管理员在人员管理里直接指定，不能申请；
    /// 2) 技术员申请：审核员或管理员都能通过；
    /// 3) 审核员申请：只有管理员能通过——审核员不能批出新的审核员。
    /// </summary>
    public static class RoleRequestRules
    {
        /// <summary>这个身份组能否通过申请获得（管理员不能申请）</summary>
        public static bool IsRequestable(UserRole role)
            => role == UserRole.Technician || role == UserRole.Reviewer;

        /// <summary>当前身份能否通过这条申请：管理员全部可批；审核员只能批技术员申请</summary>
        public static bool CanApprove(UserRole requestedRole, bool approverIsAdmin, bool approverIsReviewer)
        {
            if (!IsRequestable(requestedRole))
            {
                return approverIsAdmin;
            }
            return requestedRole == UserRole.Reviewer
                ? approverIsAdmin
                : (approverIsAdmin || approverIsReviewer);
        }

        /// <summary>身份组中文名（写进申请摘要等非界面文本；界面展示请用语言资源 Role_* 键）</summary>
        public static string DisplayName(UserRole role) => role switch
        {
            UserRole.Technician => "技术员",
            UserRole.Reviewer => "审核员",
            UserRole.Administrator => "管理员",
            UserRole.GeneralUser => "普通用户",
            _ => "游客"
        };
    }
}
