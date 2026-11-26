using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 身份组申请（review_requests 表 Type = "身份组申请" 的 PayloadJson 内容）。
    /// 申请人只能是「技术员」「审核员」两种更高身份组；管理员身份只能由管理员在人员管理里直接指定，
    /// 因此不在这里出现。
    /// </summary>
    public class RoleRequestPayload
    {
        /// <summary>申请的身份组（UserRole 枚举名：Technician / Reviewer）</summary>
        public string RequestedRole { get; set; }

        /// <summary>申请说明（为什么需要这个身份组，审批时参考）</summary>
        public string Note { get; set; }

        /// <summary>解析申请的身份组；认不出返回 null</summary>
        public UserRole? ParsedRole
            => Enum.TryParse(RequestedRole, out UserRole role) ? role : (UserRole?)null;
    }
}
