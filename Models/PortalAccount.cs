using FreeSql.DataAnnotations;
using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 公司平台（EIP 门户）账号登记（portal_accounts 表）。
    /// 用途：登录不再只认本地密码，改用公司平台账号验证；登录成功后平台会返回真实姓名，
    /// 这里把「账号（工号）→ 姓名 → 本程序身份组」的对应关系留档，
    /// 用来判断这个账号是不是本部门的人、以及进来以后是本程序的哪个身份。
    /// 密码是这台机器的开机密码（公司同步到各平台，含邮箱），按当前 Windows 用户 DPAPI 加密后留在本机库里，
    /// 供「免密进入」和邮箱发件（账号@acbel.com）复用，不会明文落库。
    /// </summary>
    [Table(Name = "portal_accounts")]
    [Index("uk_portal_account", nameof(Account), true)]
    public class PortalAccount
    {
        [Column(IsPrimary = true, IsIdentity = true)]
        public long Id { get; set; }

        /// <summary>平台账号（工号，如 22400616）</summary>
        [Column(StringLength = 64, IsNullable = false)]
        public string Account { get; set; }

        /// <summary>平台登录后返回的姓名（如 "Lucas Li(B 李志斌)"），登录成功即刷新</summary>
        [Column(StringLength = 64, IsNullable = true)]
        public string DisplayName { get; set; }

        /// <summary>邮箱（账号 + 公司邮箱域，默认 @acbel.com）</summary>
        [Column(StringLength = 128, IsNullable = true)]
        public string Email { get; set; }

        /// <summary>本程序身份组（<see cref="UserRole"/> 的字符串形式）：普通用户/技术员/审核员/管理员</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Role { get; set; }

        /// <summary>登录返回的区域（DH/DG…，用于判断是不是本部门）</summary>
        [Column(StringLength = 32, IsNullable = true)]
        public string Region { get; set; }

        /// <summary>关联的本程序用户（users 表主键；由平台账号自动建/同步）</summary>
        public long UserId { get; set; }

        /// <summary>是否允许登录（取消勾选即可停用某人，不动历史数据）</summary>
        public bool IsActive { get; set; } = true;

        /// <summary>平台密码（DPAPI 加密；仅本机本用户可解）</summary>
        [Column(StringLength = 512, IsNullable = true)]
        public string PasswordEnc { get; set; }

        /// <summary>最近一次登录成功时间</summary>
        public DateTime? LastLoginAt { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        /// <summary>身份组枚举（未知/空按普通用户处理）</summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public UserRole RoleKind
            => Enum.TryParse(Role, out UserRole role) && role != UserRole.Guest ? role : UserRole.GeneralUser;

        /// <summary>最近登录时间的文本（界面列显示用）</summary>
        [Column(IsIgnore = true)]
        [Newtonsoft.Json.JsonIgnore]
        public string LastLoginText => LastLoginAt?.ToString("yyyy/M/d HH:mm") ?? "";
    }

    /// <summary>
    /// 平台登录结果状态
    /// </summary>
    public enum PortalLoginStatus
    {
        /// <summary>登录成功</summary>
        Success,
        /// <summary>账号或密码为空</summary>
        EmptyCredentials,
        /// <summary>平台明确拒绝（账号或密码不对 / 账号被锁）</summary>
        InvalidCredentials,
        /// <summary>能登录平台，但不是本部门账号（区域不在允许列表里），也没有登记过身份</summary>
        NotDepartment,
        /// <summary>登记过但被停用</summary>
        Disabled,
        /// <summary>网络/服务端错误（连不上、超时、返回格式不对）</summary>
        ServiceError
    }

    /// <summary>
    /// 一次平台登录的结果（界面据此提示，服务层不改界面）
    /// </summary>
    public class PortalLoginOutcome
    {
        /// <summary>结果状态</summary>
        public PortalLoginStatus Status { get; set; }

        /// <summary>给用户看的提示（成功时可为空）</summary>
        public string Message { get; set; }

        /// <summary>登录成功的平台账号</summary>
        public PortalAccount Account { get; set; }

        /// <summary>平台返回的姓名（登录成功后就是本程序的用户名/显示名）</summary>
        public string DisplayName { get; set; }

        /// <summary>平台返回的区域</summary>
        public string Region { get; set; }

        /// <summary>本次登录是否顺带新建/更新了本程序账号</summary>
        public bool Enrolled { get; set; }

        public bool Success => Status == PortalLoginStatus.Success;
    }

    /// <summary>
    /// 平台登录与身份判定设置（保存在 app_settings：portal.* 键）
    /// </summary>
    public class PortalAuthSettings
    {
        /// <summary>本部门允许的区域（逗号或分号分隔，如 "DG"；空表示不做区域判断，只认已登记账号）</summary>
        public string AllowedRegions { get; set; } = "DG";

        /// <summary>公司邮箱域（登录账号 + 这个域就是邮箱账号）</summary>
        public string EmailDomain { get; set; } = "acbel.com";

        /// <summary>同部门账号首次登录时自动登记成的身份组（UserRole 名称）</summary>
        public string AutoEnrollRole { get; set; } = nameof(UserRole.Technician);

        /// <summary>同区域账号首次登录是否自动登记（关掉后只有管理员手工登记过的账号能进）</summary>
        public bool AutoEnroll { get; set; } = true;

        /// <summary>允许的区域拆成列表</summary>
        public string[] RegionList =>
            string.IsNullOrWhiteSpace(AllowedRegions)
                ? []
                : AllowedRegions.Split([',', ';', '、', ' '], StringSplitOptions.RemoveEmptyEntries);
    }
}
