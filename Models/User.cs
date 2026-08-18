using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 用户角色（游客不入库，为未登录默认身份）
    /// </summary>
    public enum UserRole
    {
        /// <summary>游客（最低）：不用登录，仅可浏览计划管理</summary>
        Guest,
        /// <summary>普通用户（低）：可浏览编辑计划管理，但编辑需提交审核</summary>
        GeneralUser,
        /// <summary>技术员（中）：可直接编辑计划管理；可使用一键报告</summary>
        Technician,
        /// <summary>审核员（中）：可编辑计划管理；可审核计划表单更改请求</summary>
        Reviewer,
        /// <summary>管理员（高）：开放所有权限</summary>
        Administrator
    }

    /// <summary>
    /// 用户实体（users 表）。密码以 SHA256(Salt+密码) 散列存储。
    /// </summary>
    [Table(Name = "users")]
    [Index("uk_username", nameof(Username), true)]
    public class User
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>
        /// 登录名（唯一）
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Username { get; set; }

        /// <summary>
        /// 显示名称
        /// </summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string DisplayName { get; set; }

        /// <summary>
        /// 密码散列 SHA256(Salt+密码)
        /// </summary>
        [Column(StringLength = 128, IsNullable = false)]
        public string PasswordHash { get; set; }

        /// <summary>
        /// 密码盐
        /// </summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Salt { get; set; }

        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsActive { get; set; } = true;

        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// 用户-角色关联（user_roles 表）。一个用户可拥有多个身份。
    /// </summary>
    [Table(Name = "user_roles")]
    public class UserRoleRow
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        public long UserId { get; set; }

        /// <summary>
        /// 角色名（UserRole枚举的字符串形式：GeneralUser/Technician/Reviewer/Administrator）
        /// </summary>
        [Column(StringLength = 32, IsNullable = false)]
        public string Role { get; set; }
    }
}
