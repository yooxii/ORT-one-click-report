using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
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
                    DisplayName = "管理员",
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
        /// 登录验证；成功时设置当前用户与角色
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
                if (user.PasswordHash != HashPassword(user.Salt, password))
                {
                    return false;
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
        /// 登出，回到游客身份
        /// </summary>
        public void Logout()
        {
            if (CurrentUser != null)
            {
                _logger.Info($"用户登出: {CurrentUser.Username}");
            }
            CurrentUser = null;
            CurrentRoles = [];
            AuthChanged?.Invoke();
        }

        /// <summary>
        /// 当前用户是否拥有指定角色
        /// </summary>
        public bool HasRole(UserRole role) => CurrentRoles.Contains(role);

        /// <summary>
        /// 当前显示名称（未登录为"游客"）
        /// </summary>
        public string CurrentDisplayName => CurrentUser == null
            ? "游客"
            : (CurrentUser.DisplayName ?? CurrentUser.Username) + $"({string.Join("/", CurrentRoles)})";

        /// <summary>
        /// 用于审计字段的操作者名称
        /// </summary>
        public string CurrentOperatorName => CurrentUser?.Username ?? Environment.UserName;

        /* ###############################  密码工具  ################################ */

        /// <summary>
        /// 创建新用户（含角色），返回错误信息；成功返回null
        /// </summary>
        public string CreateUser(string username, string displayName, string password, IEnumerable<UserRole> roles)
        {
            if (string.IsNullOrWhiteSpace(username) || password == null || password.Length < 6)
            {
                return "用户名不能为空，密码至少6位";
            }
            if (_db.FreeSql.Select<User>().Where(u => u.Username == username).Any())
            {
                return $"用户名 [{username}] 已存在";
            }
            string salt = NewSalt();
            User user = new()
            {
                Username = username,
                DisplayName = displayName,
                Salt = salt,
                PasswordHash = HashPassword(salt, password),
                IsActive = true
            };
            user.Id = _db.FreeSql.Insert(user).ExecuteIdentity();
            foreach (UserRole role in roles.Distinct())
            {
                _db.FreeSql.Insert(new UserRoleRow { UserId = user.Id, Role = role.ToString() }).ExecuteAffrows();
            }
            _logger.Info($"创建用户: {username}（角色: {string.Join(",", roles)}）");
            return null;
        }

        /// <summary>
        /// 重置用户密码，返回错误信息；成功返回null
        /// </summary>
        public string ResetPassword(long userId, string newPassword)
        {
            if (newPassword == null || newPassword.Length < 6)
            {
                return "密码至少6位";
            }
            string salt = NewSalt();
            _db.FreeSql.Update<User>()
                .Set(u => u.Salt, salt)
                .Set(u => u.PasswordHash, HashPassword(salt, newPassword))
                .Where(u => u.Id == userId)
                .ExecuteAffrows();
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
