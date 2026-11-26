using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 公司平台账号登录 + 部门身份登记。
    ///
    /// 为什么要有它：
    /// 1) 每个技术员本来就有公司平台（EIP 门户）账号，登录平台成功即可证明身份，不必再单独维护一套程序密码；
    /// 2) 但平台是全公司的，所以还要判断「是不是我们部门的人」——用登录返回的区域 + 本地登记的身份组来把关；
    /// 3) 平台登录会返回真实姓名，这个姓名就直接当本程序的用户名/显示名；
    /// 4) 本地超级管理员（admin）保留，走原来的本地密码登录，平台登录不能顶替它。
    /// </summary>
    public class PortalAccountService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AuthService _auth;
        private readonly OrtBoardSettingsService _settings;
        private readonly FineReportClient _client;

        /// <summary>本次运行内缓存已登录的会话（同一账号不重复登录；token 快过期会自动重登）</summary>
        private readonly Dictionary<string, FineReportSession> _sessions = new(StringComparer.OrdinalIgnoreCase);

        public PortalAccountService(DatabaseService db, AuthService auth, OrtBoardSettingsService settings,
            FineReportClient client)
        {
            _db = db;
            _auth = auth;
            _settings = settings;
            _client = client;
        }

        /* ###############################  登录  ################################ */

        /// <summary>
        /// 用公司平台账号密码登录：验证平台 → 判断部门/身份 → 同步本程序账号 → 置为当前登录用户。
        /// 失败不抛异常，用返回的 <see cref="PortalLoginOutcome"/> 说话（界面直接显示 Message）。
        /// </summary>
        public async Task<PortalLoginOutcome> SignInAsync(string account, string password,
            CancellationToken ct = default)
        {
            account = (account ?? "").Trim();
            if (string.IsNullOrWhiteSpace(account) || string.IsNullOrEmpty(password))
            {
                return new PortalLoginOutcome
                {
                    Status = PortalLoginStatus.EmptyCredentials,
                    Message = "请输入公司平台账号（工号）和密码"
                };
            }

            OrtBoardSettings current = _settings.Get();
            string deviceUuid = _settings.EnsureDeviceUuid(current.Endpoints);
            FineReportSession session;
            try
            {
                session = await _client.LoginAsync(account, password, deviceUuid, current.Endpoints, ct)
                    .ConfigureAwait(false);
            }
            catch (PortalAuthException ex)
            {
                _logger.Warn($"平台登录失败（{ex.Status}）：{ex.Message}");
                return new PortalLoginOutcome { Status = ex.Status, Message = ex.Message };
            }
            catch (OperationCanceledException)
            {
                return new PortalLoginOutcome { Status = PortalLoginStatus.ServiceError, Message = "登录已取消" };
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "平台登录异常");
                return new PortalLoginOutcome
                {
                    Status = PortalLoginStatus.ServiceError,
                    Message = "登录时出错：" + ex.Message
                };
            }

            _sessions[account] = session;
            string displayName = string.IsNullOrWhiteSpace(session.DisplayName) ? account : session.DisplayName.Trim();
            string region = session.Region ?? "";

            PortalAccount stored = FindByAccount(account);
            PortalAuthSettings portal = _settings.GetPortal();

            // ---- 部门/身份把关 ----
            if (stored == null)
            {
                bool inDepartment = IsDepartmentRegion(region, portal);
                if (!portal.AutoEnroll || !inDepartment)
                {
                    _logger.Warn($"账号 {account}（{displayName}）不属于本部门（区域 {region}），拒绝登录");
                    return new PortalLoginOutcome
                    {
                        Status = PortalLoginStatus.NotDepartment,
                        DisplayName = displayName,
                        Region = region,
                        Message = BuildNotDepartmentMessage(account, displayName, region, portal)
                    };
                }
                stored = new PortalAccount
                {
                    Account = account,
                    DisplayName = displayName,
                    Region = region,
                    Role = NormalizeRole(portal.AutoEnrollRole),
                    IsActive = true,
                    CreatedAt = DateTime.Now
                };
                stored.Id = _db.FreeSql.Insert(stored).ExecuteIdentity();
                _logger.Info($"同部门账号首次登录，已自动登记：{account}（{displayName}）身份={stored.Role}");
            }
            else if (!stored.IsActive)
            {
                return new PortalLoginOutcome
                {
                    Status = PortalLoginStatus.Disabled,
                    DisplayName = displayName,
                    Region = region,
                    Account = stored,
                    Message = $"账号 {account} 已停用，请联系管理员"
                };
            }

            // ---- 同步平台信息 + 本程序账号 ----
            try
            {
                stored.DisplayName = displayName;
                stored.Region = region;
                stored.Email = _settings.BuildEmail(account);
                stored.PasswordEnc = Protect(password);
                stored.LastLoginAt = DateTime.Now;
                stored.Role = NormalizeRole(stored.Role);
                User user = EnsureUser(stored);
                stored.UserId = user.Id;
                _db.FreeSql.Update<PortalAccount>().SetSource(stored).Where(a => a.Id == stored.Id).ExecuteAffrows();
                _sessions[account] = session;

                if (!_auth.SignInPortalUser(user.Id, password))
                {
                    return new PortalLoginOutcome
                    {
                        Status = PortalLoginStatus.Disabled,
                        Message = "程序账号不可用（可能被停用），请联系管理员"
                    };
                }

                _logger.Info($"平台账号登录成功：{account}（{displayName}）→ 程序账号 {user.Username}，身份 {stored.Role}");
                return new PortalLoginOutcome
                {
                    Status = PortalLoginStatus.Success,
                    Account = stored,
                    DisplayName = displayName,
                    Region = region,
                    Enrolled = true
                };
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "平台登录后同步程序账号失败");
                return new PortalLoginOutcome
                {
                    Status = PortalLoginStatus.ServiceError,
                    Message = "平台登录成功，但同步本程序账号失败：" + ex.Message
                };
            }
        }

        /// <summary>区域是否算本部门（允许列表为空时不做区域判断＝只认登记过的账号，故此处返回 false）</summary>
        private static bool IsDepartmentRegion(string region, PortalAuthSettings portal)
        {
            string[] allowed = portal.RegionList;
            if (allowed.Length == 0 || string.IsNullOrWhiteSpace(region))
            {
                return false;
            }
            return allowed.Any(a => string.Equals(a.Trim(), region.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        private string BuildNotDepartmentMessage(string account, string displayName, string region,
            PortalAuthSettings portal)
        {
            string allowed = portal.RegionList.Length == 0 ? "(未设置)" : string.Join("、", portal.RegionList);
            return $"账号 {account}（{displayName}）能登录公司平台，但不属于本部门。\n\n"
                + $"登录返回的区域：{(string.IsNullOrWhiteSpace(region) ? "(空)" : region)}\n"
                + $"本部门允许的区域：{allowed}\n\n"
                + (portal.AutoEnroll
                    ? "如果你确实属于本部门，请让管理员在「设置 → 登录与身份 → 部门账号」里登记这个账号并指定身份。"
                    : "本程序当前只允许管理员登记过的账号登录，请联系管理员登记。");
        }

        /// <summary>
        /// 保证这个平台账号有一个可用的本程序用户：姓名当用户名，邮箱=账号@域，身份组取登记的身份。
        /// 已存在（按 UserId 或同名）就更新，不覆盖本地超级管理员的密码。
        /// </summary>
        private User EnsureUser(PortalAccount account)
        {
            string desired = string.IsNullOrWhiteSpace(account.DisplayName) ? account.Account : account.DisplayName.Trim();
            if (desired.Length > 64)
            {
                desired = desired.Substring(0, 64);
            }
            User user = null;
            if (account.UserId > 0)
            {
                user = _db.FreeSql.Select<User>().Where(u => u.Id == account.UserId).First();
            }
            if (user == null)
            {
                // 同名优先复用；但已被别的平台账号占着就另起一个（避免两个人共用一个程序账号）
                User same = _db.FreeSql.Select<User>().Where(u => u.Username == desired).First();
                bool taken = same != null && _db.FreeSql.Select<PortalAccount>()
                    .Where(a => a.UserId == same.Id && a.Account != account.Account).Any();
                if (same != null && !taken && !string.Equals(same.Username, "admin", StringComparison.OrdinalIgnoreCase))
                {
                    user = same;
                }
            }

            if (user == null)
            {
                string username = desired;
                int suffix = 1;
                while (_db.FreeSql.Select<User>().Where(u => u.Username == username).Any())
                {
                    suffix++;
                    username = $"{desired} ({account.Account})";
                    if (suffix > 2)
                    {
                        username = $"{desired} {suffix}";
                    }
                    if (suffix > 99)
                    {
                        username = $"{desired} {Guid.NewGuid():N}".Substring(0, 60);
                        break;
                    }
                }
                // 平台账号不设本地密码：给一个随机哈希，避免有人只填用户名就从本地登录进来
                string salt = Guid.NewGuid().ToString("N");
                user = new User
                {
                    Username = username,
                    DisplayName = desired,
                    Email = account.Email,
                    Salt = salt,
                    PasswordHash = AuthService.HashPassword(salt, Guid.NewGuid() + "!platform"),
                    IsActive = true,
                    CreatedAt = DateTime.Now
                };
                user.Id = _db.FreeSql.Insert(user).ExecuteIdentity();
                _logger.Info($"已按平台账号创建程序用户：{username}（{account.Account}）");
            }
            else
            {
                user.DisplayName = desired;
                if (!string.IsNullOrWhiteSpace(account.Email))
                {
                    user.Email = account.Email;
                }
                if (!user.IsActive)
                {
                    user.IsActive = true;
                }
                _db.FreeSql.Update<User>().SetSource(user).Where(u => u.Id == user.Id).ExecuteAffrows();
            }

            SyncRoles(user.Id, account.RoleKind);
            account.UserId = user.Id;
            return user;
        }

        /// <summary>把该用户的身分组同步成登记的那个（多身份只保留一个，和「每个账号一个身份」一致）</summary>
        private void SyncRoles(long userId, UserRole role)
        {
            List<string> existing = _db.FreeSql.Select<UserRoleRow>().Where(r => r.UserId == userId).ToList(r => r.Role);
            string target = role.ToString();
            foreach (string stale in existing.Where(r => !string.Equals(r, target, StringComparison.Ordinal)))
            {
                _db.FreeSql.Delete<UserRoleRow>()
                    .Where(r => r.UserId == userId && r.Role == stale).ExecuteAffrows();
            }
            if (!existing.Any(r => string.Equals(r, target, StringComparison.Ordinal)))
            {
                _db.FreeSql.Insert(new UserRoleRow { UserId = userId, Role = target }).ExecuteAffrows();
            }
        }

        /* ###############################  部门账号登记  ################################ */

        /// <summary>全部已登记账号（最近登录在前）</summary>
        public List<PortalAccount> ListAccounts()
            => _db.FreeSql.Select<PortalAccount>().OrderByDescending(a => a.LastLoginAt).ToList();

        /// <summary>按账号（工号）找登记</summary>
        public PortalAccount FindByAccount(string account)
        {
            account = (account ?? "").Trim();
            return string.IsNullOrWhiteSpace(account)
                ? null
                : _db.FreeSql.Select<PortalAccount>().Where(a => a.Account == account).First();
        }

        /// <summary>按本程序用户找登记（判断当前登录是不是平台账号）</summary>
        public PortalAccount FindByUser(long userId)
            => userId <= 0 ? null : _db.FreeSql.Select<PortalAccount>().Where(a => a.UserId == userId).First();

        /// <summary>按本程序用户名找登记（登录 cookie 里存的是用户名，用来判断是不是平台账号）</summary>
        public PortalAccount FindByUserName(string username)
        {
            if (string.IsNullOrWhiteSpace(username))
            {
                return null;
            }
            User user = _db.FreeSql.Select<User>().Where(u => u.Username == username.Trim()).First();
            return user == null ? null : FindByUser(user.Id);
        }

        /// <summary>
        /// 新增或更新一个部门账号（管理员用）：账号唯一，身份组必填，邮箱按账号+域自动补。
        /// 返回错误信息；成功返回 null。
        /// </summary>
        public string SaveAccount(PortalAccount value)
        {
            if (value == null || string.IsNullOrWhiteSpace(value.Account))
            {
                return "账号（工号）不能为空";
            }
            value.Account = value.Account.Trim();
            value.Role = NormalizeRole(value.Role);
            value.Email = string.IsNullOrWhiteSpace(value.Email)
                ? _settings.BuildEmail(value.Account)
                : value.Email.Trim();

            PortalAccount existing = FindByAccount(value.Account);
            if (existing == null)
            {
                value.CreatedAt = DateTime.Now;
                _db.FreeSql.Insert(value).ExecuteAffrows();
                _logger.Info($"登记部门账号：{value.Account}（身份 {value.Role}）");
                return null;
            }
            existing.DisplayName = value.DisplayName;
            existing.Role = value.Role;
            existing.Email = value.Email;
            existing.IsActive = value.IsActive;
            _db.FreeSql.Update<PortalAccount>().SetSource(existing).Where(a => a.Id == existing.Id).ExecuteAffrows();

            // 身份改了立即同步到已绑定的程序账号
            if (existing.UserId > 0)
            {
                SyncRoles(existing.UserId, existing.RoleKind);
            }
            _logger.Info($"更新部门账号：{existing.Account}（身份 {existing.Role}，{(existing.IsActive ? "启用" : "停用")}）");
            return null;
        }

        /// <summary>删除登记（程序账号保留，只是以后不能再从平台登录进来）</summary>
        public void DeleteAccount(long id)
        {
            PortalAccount account = _db.FreeSql.Select<PortalAccount>().Where(a => a.Id == id).First();
            if (account == null)
            {
                return;
            }
            _db.FreeSql.Delete<PortalAccount>().Where(a => a.Id == id).ExecuteAffrows();
            _sessions.Remove(account.Account);
            _logger.Info($"删除部门账号登记：{account.Account}");
        }

        /* ###############################  密码（DPAPI）  ################################ */

        /// <summary>取已保存的平台密码（解不开返回 null）</summary>
        public static string GetStoredPassword(PortalAccount account)
        {
            if (account == null || string.IsNullOrWhiteSpace(account.PasswordEnc))
            {
                return null;
            }
            try
            {
                byte[] data = ProtectedData.Unprotect(Convert.FromBase64String(account.PasswordEnc), null,
                    DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(data);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>DPAPI 加密（当前 Windows 用户）</summary>
        public static string Protect(string plain)
        {
            if (string.IsNullOrEmpty(plain))
            {
                return null;
            }
            try
            {
                byte[] encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(plain), null,
                    DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>最近登录过的几个账号（登录窗口的快捷入口）</summary>
        public List<PortalAccount> RecentAccounts(int count = 6)
            => _db.FreeSql.Select<PortalAccount>()
                .Where(a => a.IsActive)
                .OrderByDescending(a => a.LastLoginAt)
                .Limit(count)
                .ToList();

        /// <summary>身份组名称归一（非法值按普通用户）</summary>
        public static string NormalizeRole(string role)
            => Enum.TryParse(role, out UserRole parsed) && parsed != UserRole.Guest
                ? parsed.ToString()
                : nameof(UserRole.GeneralUser);

        /// <summary>缓存里的会话（没有返回 null）</summary>
        public FineReportSession GetCachedSession(string account)
            => _sessions.TryGetValue((account ?? "").Trim(), out FineReportSession session) ? session : null;

        /// <summary>清掉某账号的会话缓存（改密码/退出时用）</summary>
        public void DropSession(string account)
        {
            if (!string.IsNullOrWhiteSpace(account))
            {
                _sessions.Remove(account.Trim());
            }
        }
    }
}
