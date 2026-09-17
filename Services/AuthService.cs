using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 认证服务：登录态管理、密码散列校验、默认管理员初始化。
    /// 未登录时为游客身份（Guest）。
    /// </summary>
    public class AuthService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;

        /// <summary>
        /// 当前登录用户（未登录为null）
        /// </summary>
        public User CurrentUser { get; private set; }

        /// <summary>
        /// 本次登录是"账号还没设置密码"直接进来的（登录后应提示去用户中心设置密码）
        /// </summary>
        public bool PasswordlessLogin { get; private set; }

        /// <summary>
        /// 当前用户的角色列表（未登录为空，视为游客）
        /// </summary>
        public List<UserRole> CurrentRoles { get; private set; } = [];

        /// <summary>
        /// 登录/登出时触发（供UI刷新权限状态）
        /// </summary>
        public event Action AuthChanged;

        public AuthService(DatabaseService db)
        {
            _db = db;
            EnsureDefaultAdmin();
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 首次运行无用户时创建默认管理员 admin/admin123
        /// </summary>
        private void EnsureDefaultAdmin()
        {
            try
            {
                if (_db.FreeSql.Select<User>().Count() > 0)
                {
                    return;
                }
                string salt = NewSalt();
                User admin = new()
                {
                    Username = "admin",
                    // 默认管理员显示名跟随界面语言（后续可在人员管理中修改）
                    DisplayName = LanguageService.Get("Role_Administrator"),
                    Salt = salt,
                    PasswordHash = HashPassword(salt, "admin123"),
                    IsActive = true
                };
                admin.Id = _db.FreeSql.Insert(admin).ExecuteIdentity();
                _db.FreeSql.Insert(new UserRoleRow { UserId = admin.Id, Role = nameof(UserRole.Administrator) }).ExecuteAffrows();
                _logger.Info("已创建默认管理员账号 admin/admin123，请尽快修改密码");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "初始化默认管理员失败");
            }
        }

        /// <summary>
        /// 登录验证；成功时设置当前用户与角色。
        /// 账号**还没有设置密码**时，只给用户名（密码留空）也允许登录，登录后由界面提示去设置密码；
        /// 账号已设置密码时按散列校验。
        /// </summary>
        public bool Login(string username, string password)
        {
            try
            {
                User user = _db.FreeSql.Select<User>().Where(u => u.Username == username).First();
                if (user == null || !user.IsActive)
                {
                    return false;
                }
                bool hasPassword = HasPassword(user);
                if (hasPassword)
                {
                    if (user.PasswordHash != HashPassword(user.Salt, password ?? ""))
                    {
                        return false;
                    }
                }
                else
                {
                    // 没设过密码：用户名对就放行（密码框留空即可），登录后提示设置密码
                    PasswordlessLogin = true;
                    _logger.Info($"用户 {user.Username} 尚未设置密码，按用户名直接登录");
                }
                CurrentUser = user;
                CurrentRoles = _db.FreeSql.Select<UserRoleRow>()
                    .Where(r => r.UserId == user.Id)
                    .ToList()
                    .Select(r => Enum.TryParse<UserRole>(r.Role, out UserRole role) ? role : (UserRole?)null)
                    .Where(r => r.HasValue)
                    .Select(r => r.Value)
                    .ToList();
                _logger.Info($"用户登录: {user.Username}（角色: {string.Join(",", CurrentRoles)}）");
                SaveLoginCookie(username, password ?? "");
                AuthChanged?.Invoke();
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "登录失败");
                return false;
            }
        }

        /// <summary>
        /// 登出，回到游客身份；同时完全清除本地登录 cookie
        /// </summary>
        public void Logout()
        {
            if (CurrentUser != null)
            {
                _logger.Info($"用户登出: {CurrentUser.Username}");
            }
            CurrentUser = null;
            CurrentRoles = [];
            PasswordlessLogin = false;
            ClearLoginCookie();
            AuthChanged?.Invoke();
        }

        /// <summary>
        /// 该账号是否设置过密码（没有则允许只凭用户名登录）
        /// </summary>
        public static bool HasPassword(User user)
            => user != null && !string.IsNullOrWhiteSpace(user.PasswordHash) && !string.IsNullOrWhiteSpace(user.Salt);

        /// <summary>
        /// 当前用户是否需要设置密码（账号没设过，或本次就是无密码登录进来的）
        /// </summary>
        public bool NeedsPasswordSetup => CurrentUser != null && (PasswordlessLogin || !HasPassword(CurrentUser));

        /// <summary>
        /// 校验当前登录用户的密码（用于"修改密码"前确认本人）
        /// </summary>
        public bool VerifyCurrentUserPassword(string password)
            => CurrentUser != null && HasPassword(CurrentUser)
               && CurrentUser.PasswordHash == HashPassword(CurrentUser.Salt, password ?? "");

        /// <summary>
        /// 设置/修改当前登录用户的密码（至少 6 位），返回错误信息；成功返回 null。
        /// 成功后同步更新本地 cookie（旧密码已失效），下次"继续上次登录"不会失败。
        /// </summary>
        public string SetCurrentUserPassword(string newPassword)
        {
            if (CurrentUser == null)
            {
                return "请先登录";
            }
            if (string.IsNullOrWhiteSpace(newPassword) || newPassword.Length < 6)
            {
                return "密码至少6位";
            }
            string salt = NewSalt();
            string hash = HashPassword(salt, newPassword);
            _db.FreeSql.Update<User>()
                .Set(u => u.Salt, salt)
                .Set(u => u.PasswordHash, hash)
                .Where(u => u.Id == CurrentUser.Id)
                .ExecuteAffrows();
            CurrentUser.Salt = salt;
            CurrentUser.PasswordHash = hash;
            PasswordlessLogin = false;
            SaveLoginCookie(CurrentUser.Username, newPassword);
            _logger.Info($"用户 {CurrentUser.Username} 已设置/修改密码");
            return null;
        }

        /// <summary>
        /// 修改当前登录用户的显示名（留空则界面回退显示登录名）
        /// </summary>
        public bool SetCurrentUserDisplayName(string displayName)
        {
            if (CurrentUser == null)
            {
                return false;
            }
            displayName = displayName?.Trim();
            _db.FreeSql.Update<User>()
                .Set(u => u.DisplayName, string.IsNullOrEmpty(displayName) ? null : displayName)
                .Where(u => u.Id == CurrentUser.Id)
                .ExecuteAffrows();
            CurrentUser.DisplayName = displayName;
            AuthChanged?.Invoke();
            return true;
        }

        /* ###############################  本地登录 Cookie  ################################ */

        /// <summary>
        /// 登录 cookie 文件（程序目录 Data 下），密码以 DPAPI 按当前 Windows 用户加密
        /// </summary>
        private static string CookieFile => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "auth_cookie.json");

        /// <summary>
        /// cookie 有效期：一周
        /// </summary>
        private static readonly TimeSpan CookieLifetime = TimeSpan.FromDays(7);

        private class CookieData
        {
            public string Username { get; set; }
            public string PasswordEnc { get; set; }
            public DateTime Expiry { get; set; }
        }

        /// <summary>
        /// 保存登录信息到本地 cookie（保留上一次登录，有效期一周）
        /// </summary>
        private void SaveLoginCookie(string username, string password)
        {
            try
            {
                byte[] encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(password), null, DataProtectionScope.CurrentUser);
                CookieData cookie = new()
                {
                    Username = username,
                    PasswordEnc = Convert.ToBase64String(encrypted),
                    Expiry = DateTime.Now + CookieLifetime
                };
                Directory.CreateDirectory(Path.GetDirectoryName(CookieFile));
                File.WriteAllText(CookieFile, JsonConvert.SerializeObject(cookie));
            }
            catch (Exception ex)
            {
                _logger.Warn($"保存登录 cookie 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 读取未过期的登录 cookie；不存在/已过期/损坏返回 null（过期时自动清除）
        /// </summary>
        public (string Username, string Password)? LoadValidCookie()
        {
            try
            {
                if (!File.Exists(CookieFile))
                {
                    return null;
                }
                CookieData cookie = JsonConvert.DeserializeObject<CookieData>(File.ReadAllText(CookieFile));
                if (cookie == null || string.IsNullOrWhiteSpace(cookie.Username) || string.IsNullOrWhiteSpace(cookie.PasswordEnc))
                {
                    return null;
                }
                if (cookie.Expiry < DateTime.Now)
                {
                    ClearLoginCookie();
                    return null;
                }
                byte[] decrypted = ProtectedData.Unprotect(Convert.FromBase64String(cookie.PasswordEnc), null, DataProtectionScope.CurrentUser);
                return (cookie.Username, Encoding.UTF8.GetString(decrypted));
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取登录 cookie 失败: {ex.Message}");
                ClearLoginCookie();
                return null;
            }
        }

        /// <summary>
        /// cookie 到期时间（无有效 cookie 时返回 null，供界面提示）
        /// </summary>
        public DateTime? GetCookieExpiry()
        {
            try
            {
                if (!File.Exists(CookieFile))
                {
                    return null;
                }
                CookieData cookie = JsonConvert.DeserializeObject<CookieData>(File.ReadAllText(CookieFile));
                return cookie != null && cookie.Expiry >= DateTime.Now ? cookie.Expiry : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 完全清除本地登录 cookie（注销时调用）
        /// </summary>
        public void ClearLoginCookie()
        {
            try
            {
                if (File.Exists(CookieFile))
                {
                    File.Delete(CookieFile);
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"清除登录 cookie 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 当前用户是否拥有指定角色
        /// </summary>
        public bool HasRole(UserRole role) => CurrentRoles.Contains(role);

        /// <summary>
        /// 技术员/审核员是否为邮箱为空（登录时提示完善）
        /// </summary>
        public bool NeedsEmailCompletion => CurrentUser != null
            && string.IsNullOrWhiteSpace(CurrentUser.Email)
            && (HasRole(UserRole.Technician) || HasRole(UserRole.Reviewer));

        /// <summary>
        /// 邮箱格式校验（简单校验：有且仅有一个 @，域名含点，不含空白）
        /// </summary>
        public static bool IsValidEmail(string email)
            => !string.IsNullOrWhiteSpace(email)
               && System.Text.RegularExpressions.Regex.IsMatch(email.Trim(), @"^[^@\s]+@[^@\s]+\.[^@\s]+$");

        /// <summary>
        /// 保存当前登录用户的邮箱（供登录后"完善邮箱"使用；传空表示清空邮箱）
        /// </summary>
        public bool SetCurrentUserEmail(string email)
        {
            if (CurrentUser == null)
            {
                return false;
            }
            email = email?.Trim();
            if (string.IsNullOrEmpty(email))
            {
                _db.FreeSql.Update<User>().Set(u => u.Email, (string)null).Where(u => u.Id == CurrentUser.Id).ExecuteAffrows();
                CurrentUser.Email = null;
                return true;
            }
            if (!IsValidEmail(email))
            {
                return false;
            }
            _db.FreeSql.Update<User>()
                .Set(u => u.Email, email)
                .Where(u => u.Id == CurrentUser.Id)
                .ExecuteAffrows();
            CurrentUser.Email = email;
            _logger.Info($"用户 {CurrentUser.Username} 完善邮箱: {email}");
            return true;
        }

        /// <summary>
        /// 保存当前登录用户的基本资料（显示名 + 邮箱；邮箱/显示名留空表示清空），
        /// 返回错误信息；成功返回 null
        /// </summary>
        public string SaveCurrentUserProfile(string displayName, string email)
        {
            if (CurrentUser == null)
            {
                return "请先登录";
            }
            email = email?.Trim();
            if (!string.IsNullOrEmpty(email) && !IsValidEmail(email))
            {
                return "邮箱格式不正确，请重新输入";
            }
            User row = _db.FreeSql.Select<User>().Where(u => u.Id == CurrentUser.Id).First();
            if (row == null)
            {
                return "账号不存在或已被删除";
            }
            row.DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName.Trim();
            row.Email = string.IsNullOrEmpty(email) ? null : email;
            _db.FreeSql.Update<User>().SetSource(row).Where(u => u.Id == row.Id).ExecuteAffrows();
            CurrentUser.DisplayName = row.DisplayName;
            CurrentUser.Email = row.Email;
            _logger.Info($"用户 {CurrentUser.Username} 更新资料：显示名={row.DisplayName ?? "-"}，邮箱={row.Email ?? "-"}");
            AuthChanged?.Invoke();
            return null;
        }

        /// <summary>
        /// 按测试项目负责人姓名确保存在技术员账号：
        /// 账号不存在则以初始密码 123456 创建，显示名为该姓名，登录名见 <see cref="BuildUsernameFromName"/>；
        /// 账号已存在则仅补充技术员身份，不修改其密码。
        /// </summary>
        public (bool Created, bool RoleAdded) EnsureTechnician(string name)
        {
            name = name?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                return (false, false);
            }
            try
            {
                // 优先按登录名匹配，其次按显示名匹配，避免同一人重复建号
                User user = _db.FreeSql.Select<User>().Where(u => u.Username == name).First()
                    ?? _db.FreeSql.Select<User>().Where(u => u.DisplayName == name).First();
                if (user == null)
                {
                    string username = BuildUsernameFromName(name);
                    string salt = NewSalt();
                    User created = new()
                    {
                        Username = username,
                        DisplayName = name,
                        Salt = salt,
                        PasswordHash = HashPassword(salt, DefaultTechnicianPassword),
                        IsActive = true
                    };
                    created.Id = _db.FreeSql.Insert(created).ExecuteIdentity();
                    _db.FreeSql.Insert(new UserRoleRow { UserId = created.Id, Role = nameof(UserRole.Technician) }).ExecuteAffrows();
                    _logger.Info($"按测试项目负责人创建技术员账号: {name}（登录名 {username}，初始密码 {DefaultTechnicianPassword}）");
                    return (true, false);
                }
                bool hasTechnician = _db.FreeSql.Select<UserRoleRow>()
                    .Where(r => r.UserId == user.Id && r.Role == nameof(UserRole.Technician)).Any();
                if (hasTechnician)
                {
                    return (false, false);
                }
                _db.FreeSql.Insert(new UserRoleRow { UserId = user.Id, Role = nameof(UserRole.Technician) }).ExecuteAffrows();
                _logger.Info($"已为已有账号补充技术员身份: {name}");
                return (false, true);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"确保技术员账号失败: {name}");
                return (false, false);
            }
        }

        /// <summary>
        /// 由负责人姓名生成登录名：
        /// 英文数字姓名（如 ZhangSan）直接使用；含中文等非 ASCII 字符时无法离线转拼音，
        /// 退化为递增登录名 user1、user2……（可在人员管理中改名为拼音）。
        /// </summary>
        public string BuildUsernameFromName(string name)
        {
            name = name?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                return NextIncrementalUsername();
            }
            // 仅当姓名本身只含 ASCII 字母/数字/._- 时直接作登录名
            if (name.All(c => c < 128 && (char.IsLetterOrDigit(c) || c == '.' || c == '_' || c == '-')))
            {
                return name;
            }
            return NextIncrementalUsername();
        }

        /// <summary>
        /// 生成递增登录名 userN（跳过已占用的编号）
        /// </summary>
        public string NextIncrementalUsername()
        {
            List<string> existing = _db.FreeSql.Select<User>().ToList(u => u.Username);
            for (int i = 1; i < int.MaxValue; i++)
            {
                string candidate = "user" + i;
                if (!existing.Any(u => string.Equals(u, candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    return candidate;
                }
            }
            return "user" + Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        /// <summary>
        /// 自动创建的技术员账号初始密码
        /// </summary>
        public const string DefaultTechnicianPassword = "123456";

        /// <summary>
        /// 重新从数据库载入当前登录用户（人员管理修改了自己的登录名/显示名/邮箱后调用）
        /// </summary>
        public void ReloadCurrentUser()
        {
            if (CurrentUser == null)
            {
                return;
            }
            User latest = _db.FreeSql.Select<User>().Where(u => u.Id == CurrentUser.Id).First();
            if (latest != null)
            {
                CurrentUser = latest;
                AuthChanged?.Invoke();
            }
        }

        /// <summary>
        /// 当前显示名称（未登录为"游客"）
        /// </summary>
        public string CurrentDisplayName => CurrentUser == null
            ? LanguageService.Get("Role_Guest")
            : (CurrentUser.DisplayName ?? CurrentUser.Username) + $"({string.Join("/", CurrentRoles.Select(r => LanguageService.Get("Role_" + r)))})";

        /// <summary>
        /// 用于审计字段的操作者名称
        /// </summary>
        public string CurrentOperatorName => CurrentUser?.Username ?? Environment.UserName;

        /* ###############################  密码工具  ################################ */

        /// <summary>
        /// 创建新用户（含角色），返回错误信息；成功返回null。
        /// 密码留空表示"暂不设密码"：该账号可只凭用户名登录，登录后会提示本人设置密码。
        /// </summary>
        public string CreateUser(string username, string displayName, string password, IEnumerable<UserRole> roles)
        {
            if (string.IsNullOrWhiteSpace(username))
            {
                return "用户名不能为空";
            }
            if (!string.IsNullOrEmpty(password) && password.Length < 6)
            {
                return "密码至少6位（留空表示暂不设密码）";
            }
            if (_db.FreeSql.Select<User>().Where(u => u.Username == username).Any())
            {
                return $"用户名 [{username}] 已存在";
            }
            bool passwordless = string.IsNullOrEmpty(password);
            string salt = passwordless ? "" : NewSalt();
            User user = new()
            {
                Username = username,
                DisplayName = displayName,
                Salt = salt,
                PasswordHash = passwordless ? "" : HashPassword(salt, password),
                IsActive = true
            };
            user.Id = _db.FreeSql.Insert(user).ExecuteIdentity();
            foreach (UserRole role in roles.Distinct())
            {
                _db.FreeSql.Insert(new UserRoleRow { UserId = user.Id, Role = role.ToString() }).ExecuteAffrows();
            }
            _logger.Info(passwordless
                ? $"创建用户: {username}（未设密码，可凭用户名登录后自行设置；角色: {string.Join(",", roles)}）"
                : $"创建用户: {username}（角色: {string.Join(",", roles)}）");
            return null;
        }

        /// <summary>
        /// 重置用户密码，返回错误信息；成功返回null。
        /// 新密码留空表示清除密码（该账号改为只凭用户名登录，登录后提示本人设置密码）。
        /// </summary>
        public string ResetPassword(long userId, string newPassword)
        {
            if (!string.IsNullOrEmpty(newPassword) && newPassword.Length < 6)
            {
                return "密码至少6位（留空表示清除密码）";
            }
            bool passwordless = string.IsNullOrEmpty(newPassword);
            string salt = passwordless ? "" : NewSalt();
            _db.FreeSql.Update<User>()
                .Set(u => u.Salt, salt)
                .Set(u => u.PasswordHash, passwordless ? "" : HashPassword(salt, newPassword))
                .Where(u => u.Id == userId)
                .ExecuteAffrows();
            _logger.Info(passwordless
                ? $"已清除用户 #{userId} 的密码（改为凭用户名登录）"
                : $"已重置用户 #{userId} 的密码");
            return null;
        }

        public static string HashPassword(string salt, string password)
        {
            using SHA256 sha = SHA256.Create();
            byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(salt + password));
            return Convert.ToBase64String(bytes);
        }

        private static string NewSalt() => Guid.NewGuid().ToString("N");
    }
}
