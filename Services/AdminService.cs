using NLog;
using OfficeOpenXml;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 用户视图：用户 + 角色列表（人员管理展示用）
    /// </summary>
    public class UserView
    {
        public long Id { get; set; }
        public string Username { get; set; }
        public string DisplayName { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public List<UserRole> Roles { get; set; } = [];
        public string RolesText => string.Join("、", Roles.Select(r => RoleDisplayName(r)));

        public static string RoleDisplayName(UserRole role) => role switch
        {
            UserRole.GeneralUser => "普通用户",
            UserRole.Technician => "技术员",
            UserRole.Reviewer => "审核员",
            UserRole.Administrator => "管理员",
            _ => role.ToString()
        };
    }

    /// <summary>
    /// 管理服务：人员管理（用户+角色）、客户管理、测试项目管理。
    /// 客户与测试项目的数据源在计划表中，可一键同步。
    /// </summary>
    public class AdminService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;

        public AdminService(DatabaseService db)
        {
            _db = db;
        }

        /* ###############################  人员管理  ################################ */

        /// <summary>
        /// 用户列表（含角色）
        /// </summary>
        public List<UserView> GetUsers()
        {
            List<User> users = _db.FreeSql.Select<User>().OrderBy(u => u.Id).ToList();
            List<UserRoleRow> roles = _db.FreeSql.Select<UserRoleRow>().ToList();
            return users.Select(u => new UserView
            {
                Id = u.Id,
                Username = u.Username,
                DisplayName = u.DisplayName,
                IsActive = u.IsActive,
                CreatedAt = u.CreatedAt,
                Roles = roles.Where(r => r.UserId == u.Id)
                    .Select(r => Enum.TryParse<UserRole>(r.Role, out UserRole role) ? role : (UserRole?)null)
                    .Where(r => r.HasValue)
                    .Select(r => r.Value)
                    .ToList()
            }).ToList();
        }

        /// <summary>
        /// 更新用户角色（管理员调动用户身份）
        /// </summary>
        public void UpdateUserRoles(long userId, IEnumerable<UserRole> roles)
        {
            _db.FreeSql.Delete<UserRoleRow>().Where(r => r.UserId == userId).ExecuteAffrows();
            foreach (UserRole role in roles.Distinct())
            {
                _db.FreeSql.Insert(new UserRoleRow { UserId = userId, Role = role.ToString() }).ExecuteAffrows();
            }
            _logger.Info($"更新用户(Id={userId})角色: {string.Join(",", roles)}");
        }

        /// <summary>
        /// 启用/禁用用户
        /// </summary>
        public void SetUserActive(long userId, bool active)
        {
            _db.FreeSql.Update<User>().Set(u => u.IsActive, active).Where(u => u.Id == userId).ExecuteAffrows();
            _logger.Info($"用户(Id={userId})已{(active ? "启用" : "禁用")}");
        }

        /// <summary>
        /// 更新显示名称
        /// </summary>
        public void UpdateDisplayName(long userId, string displayName)
        {
            _db.FreeSql.Update<User>().Set(u => u.DisplayName, displayName).Where(u => u.Id == userId).ExecuteAffrows();
        }

        /* ###############################  客户管理  ################################ */

        public List<Customer> GetCustomers()
            => _db.FreeSql.Select<Customer>().OrderBy(c => c.Name).ToList();

        /// <summary>
        /// 新增或更新客户，返回错误信息；成功返回null
        /// </summary>
        public string SaveCustomer(Customer customer)
        {
            if (string.IsNullOrWhiteSpace(customer.Name))
            {
                return "客户名称不能为空";
            }
            bool exists = _db.FreeSql.Select<Customer>()
                .Where(c => c.Name == customer.Name && c.Id != customer.Id).Any();
            if (exists)
            {
                return $"客户 [{customer.Name}] 已存在";
            }
            if (customer.Id == 0)
            {
                _db.FreeSql.Insert(customer).ExecuteAffrows();
            }
            else
            {
                _db.FreeSql.Update<Customer>().SetSource(customer).Where(c => c.Id == customer.Id).ExecuteAffrows();
            }
            return null;
        }

        public void DeleteCustomer(long id)
            => _db.FreeSql.Delete<Customer>().Where(c => c.Id == id).ExecuteAffrows();

        /// <summary>
        /// 从计划数据（plans 表的客户别）同步客户字典
        /// </summary>
        public int SyncCustomersFromPlans()
        {
            List<string> names = _db.FreeSql.Select<Plan>()
                .Where(p => p.Customer != null)
                .ToList(p => p.Customer)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct()
                .ToList();
            int added = 0;
            foreach (string name in names)
            {
                if (!_db.FreeSql.Select<Customer>().Where(c => c.Name == name).Any())
                {
                    _db.FreeSql.Insert(new Customer { Name = name }).ExecuteAffrows();
                    added++;
                }
            }
            _logger.Info($"从计划数据同步客户: 新增{added}个");
            return added;
        }

        /* ###############################  测试项目管理  ################################ */

        public List<TestItemCatalog> GetTestItems()
            => _db.FreeSql.Select<TestItemCatalog>().OrderBy(t => t.Name).ToList();

        /// <summary>
        /// 新增或更新测试项目，返回错误信息；成功返回null
        /// </summary>
        public string SaveTestItem(TestItemCatalog item)
        {
            if (string.IsNullOrWhiteSpace(item.Name))
            {
                return "测试项目名称不能为空";
            }
            bool exists = _db.FreeSql.Select<TestItemCatalog>()
                .Where(t => t.Name == item.Name && t.Id != item.Id).Any();
            if (exists)
            {
                return $"测试项目 [{item.Name}] 已存在";
            }
            if (item.Id == 0)
            {
                _db.FreeSql.Insert(item).ExecuteAffrows();
            }
            else
            {
                _db.FreeSql.Update<TestItemCatalog>().SetSource(item).Where(t => t.Id == item.Id).ExecuteAffrows();
            }
            return null;
        }

        public void DeleteTestItem(long id)
            => _db.FreeSql.Delete<TestItemCatalog>().Where(t => t.Id == id).ExecuteAffrows();

        /// <summary>
        /// 从计划表文件的 "Test Items" 工作表同步测试项目字典（项次/试验项目/试验时间/负责人/备考）
        /// </summary>
        public int SyncTestItemsFromScheduleFile(string filePath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Lucas");
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage package = new(fs);
            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "Test Items")
                ?? throw new InvalidDataException("未找到 Test Items 工作表");

            // 定位表头行（含"試驗項目"）
            int headerRow = 0;
            int colName = 0, colPeriod = 0, colOwner = 0, colRemark = 0;
            int endCol = ws.Dimension?.End.Column ?? 0;
            for (int r = 1; r <= Math.Min(ws.Dimension?.End.Row ?? 0, 10); r++)
            {
                for (int c = 1; c <= endCol; c++)
                {
                    string text = Norm(ws.Cells[r, c].Text);
                    if (text.Contains("試驗項目"))
                    {
                        headerRow = r;
                        colName = c;
                    }
                    else if (text.Contains("試驗時間"))
                    {
                        colPeriod = c;
                    }
                    else if (text.Contains("負責人"))
                    {
                        colOwner = c;
                    }
                    else if (text.Contains("備考"))
                    {
                        colRemark = c;
                    }
                }
                if (headerRow > 0)
                {
                    break;
                }
            }
            if (headerRow == 0)
            {
                throw new InvalidDataException("未找到 Test Items 表头(試驗項目)");
            }

            int added = 0, updated = 0;
            int endRow = ws.Dimension?.End.Row ?? 0;
            for (int r = headerRow + 1; r <= endRow; r++)
            {
                string name = NullIfEmpty(ws.Cells[r, colName].Text);
                if (name == null)
                {
                    continue;
                }
                TestItemCatalog existing = _db.FreeSql.Select<TestItemCatalog>().Where(t => t.Name == name).First();
                if (existing == null)
                {
                    _db.FreeSql.Insert(new TestItemCatalog
                    {
                        Name = name,
                        Period = colPeriod > 0 ? NullIfEmpty(ws.Cells[r, colPeriod].Text) : null,
                        Owner = colOwner > 0 ? NullIfEmpty(ws.Cells[r, colOwner].Text) : null,
                        Remark = colRemark > 0 ? NullIfEmpty(ws.Cells[r, colRemark].Text) : null
                    }).ExecuteAffrows();
                    added++;
                }
                else
                {
                    existing.Period = colPeriod > 0 ? NullIfEmpty(ws.Cells[r, colPeriod].Text) : existing.Period;
                    existing.Owner = colOwner > 0 ? NullIfEmpty(ws.Cells[r, colOwner].Text) : existing.Owner;
                    existing.Remark = colRemark > 0 ? NullIfEmpty(ws.Cells[r, colRemark].Text) : existing.Remark;
                    _db.FreeSql.Update<TestItemCatalog>().SetSource(existing).Where(t => t.Id == existing.Id).ExecuteAffrows();
                    updated++;
                }
            }
            _logger.Info($"从计划表文件同步测试项目: 新增{added}个, 更新{updated}个");
            return added;
        }

        private static string Norm(string s) => s?.Replace(" ", "").Replace("\n", "").Replace("\r", "") ?? "";

        private static string NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();
    }
}
