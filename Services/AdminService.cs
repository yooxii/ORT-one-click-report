using NLog;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ORT一键报告.Utils;
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
        public string Email { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public List<UserRole> Roles { get; set; } = [];
        public string RolesText => string.Join(LanguageService.Get("Common_ListSeparator"), Roles.Select(RoleDisplayName));

        /// <summary>
        /// 角色显示名（跟随界面语言本地化，资源缺失时回退为角色枚举名）
        /// </summary>
        public static string RoleDisplayName(UserRole role)
        {
            string key = "Role_" + role;
            string text = LanguageService.Get(key);
            return text == key ? role.ToString() : text;
        }
    }

    /// <summary>
    /// 管理服务：人员管理（用户+角色）、客户/测试项目/阶段字典管理、机种映射。
    /// 客户表整合客户别与产品别对应关系（Cust. Code 表）。
    /// 字典数据源均在计划表中，可随计划表导入一并同步。
    /// </summary>
    public class AdminService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AuthService _auth;

        public AdminService(DatabaseService db, AuthService auth)
        {
            _db = db;
            _auth = auth;
            EnsureDefaultStages();
        }

        /* ###############################  人员管理  ################################ */

        public List<UserView> GetUsers()
        {
            List<User> users = _db.FreeSql.Select<User>().OrderBy(u => u.Id).ToList();
            List<UserRoleRow> roles = _db.FreeSql.Select<UserRoleRow>().ToList();
            return users.Select(u => new UserView
            {
                Id = u.Id,
                Username = u.Username,
                DisplayName = u.DisplayName,
                Email = u.Email,
                IsActive = u.IsActive,
                CreatedAt = u.CreatedAt,
                Roles = roles.Where(r => r.UserId == u.Id)
                    .Select(r => Enum.TryParse<UserRole>(r.Role, out UserRole role) ? role : (UserRole?)null)
                    .Where(r => r.HasValue)
                    .Select(r => r.Value)
                    .ToList()
            }).ToList();
        }

        public void UpdateUserRoles(long userId, IEnumerable<UserRole> roles)
        {
            _db.FreeSql.Delete<UserRoleRow>().Where(r => r.UserId == userId).ExecuteAffrows();
            foreach (UserRole role in roles.Distinct())
            {
                _db.FreeSql.Insert(new UserRoleRow { UserId = userId, Role = role.ToString() }).ExecuteAffrows();
            }
            _logger.Info($"更新用户(Id={userId})角色: {string.Join(",", roles)}");
        }

        public void SetUserActive(long userId, bool active)
        {
            _db.FreeSql.Update<User>().Set(u => u.IsActive, active).Where(u => u.Id == userId).ExecuteAffrows();
            _logger.Info($"用户(Id={userId})已{(active ? "启用" : "禁用")}");
        }

        public void UpdateDisplayName(long userId, string displayName)
        {
            _db.FreeSql.Update<User>().Set(u => u.DisplayName, displayName).Where(u => u.Id == userId).ExecuteAffrows();
        }

        /// <summary>
        /// 更新用户登录名（唯一），返回错误信息；成功返回null
        /// </summary>
        public string UpdateUsername(long userId, string username)
        {
            username = username?.Trim();
            if (string.IsNullOrWhiteSpace(username))
            {
                return LanguageService.Get("Msg_UsernameRequired");
            }
            if (username.Contains(' ') || username.Contains('/'))
            {
                return LanguageService.Get("Msg_UsernameInvalid");
            }
            bool exists = GetUsers().Any(u => u.Id != userId
                && string.Equals(u.Username, username, StringComparison.OrdinalIgnoreCase));
            if (exists)
            {
                return string.Format(LanguageService.Get("Msg_UsernameExists"), username);
            }
            _db.FreeSql.Update<User>().Set(u => u.Username, username).Where(u => u.Id == userId).ExecuteAffrows();
            _logger.Info($"更新用户(Id={userId})登录名: {username}");
            return null;
        }

        /// <summary>
        /// 更新用户邮箱（技术员/审核员登录时据此判断是否需要提示完善），返回错误信息；成功返回null
        /// </summary>
        public string UpdateEmail(long userId, string email)
        {
            email = string.IsNullOrWhiteSpace(email) ? null : email.Trim();
            if (email != null && !AuthService.IsValidEmail(email))
            {
                return LanguageService.Get("Msg_EmailInvalid");
            }
            _db.FreeSql.Update<User>().Set(u => u.Email, email).Where(u => u.Id == userId).ExecuteAffrows();
            _logger.Info($"更新用户(Id={userId})邮箱: {email ?? "(清空)"}");
            return null;
        }

        /// <summary>
        /// 技术员列表（按显示名排序，供测试项目"负责人"多选使用）
        /// </summary>
        public List<UserView> GetTechnicians()
            => GetUsers()
                .Where(u => u.Roles.Contains(UserRole.Technician))
                .OrderBy(u => string.IsNullOrWhiteSpace(u.DisplayName) ? u.Username : u.DisplayName,
                         StringComparer.CurrentCultureIgnoreCase)
                .ToList();

        /// <summary>
        /// 拆分测试项目负责人文本：以"/"（含全角"／"）分隔，去空白去重
        /// </summary>
        public static List<string> SplitOwners(string owners)
            => string.IsNullOrWhiteSpace(owners)
                ? []
                : owners.Split(['/', '／'], StringSplitOptions.RemoveEmptyEntries)
                    .Select(o => o.Trim())
                    .Where(o => o.Length > 0)
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

        /// <summary>
        /// 确保指定负责人均具备技术员身份（不存在则按初始密码 123456 建号）。
        /// 返回（新建账号数, 补充身份数）。
        /// </summary>
        public (int Created, int RoleAdded) EnsureTechnicianUsers(IEnumerable<string> ownerNames)
        {
            int created = 0, roleAdded = 0;
            foreach (string name in ownerNames ?? [])
            {
                (bool isCreated, bool isRoleAdded) = _auth.EnsureTechnician(name);
                if (isCreated)
                {
                    created++;
                }
                if (isRoleAdded)
                {
                    roleAdded++;
                }
            }
            return (created, roleAdded);
        }

        /// <summary>
        /// 按全部测试项目的负责人同步技术员身份（含历史数据），
        /// 并把负责人由姓名解析为 uid 锚定（OwnerIds），同时把 Owner 文本刷成当前显示名。
        /// 返回（新建账号数, 补充身份数）。
        /// </summary>
        public (int Created, int RoleAdded) SyncTechniciansFromTestItemOwners()
        {
            List<TestItemCatalog> items = _db.FreeSql.Select<TestItemCatalog>().ToList();
            List<string> names = items
                .SelectMany(t => SplitOwners(t.Owner))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            (int created, int roleAdded) = EnsureTechnicianUsers(names);
            // 解析 uid 并刷新 Owner 文本（显示名可能已修改）
            foreach (TestItemCatalog item in items)
            {
                if (string.IsNullOrWhiteSpace(item.Owner) && string.IsNullOrWhiteSpace(item.OwnerIds))
                {
                    continue;
                }
                List<long> ids = ResolveOwnerIds(item.Owner);
                string display = BuildOwnerDisplay(ids, item.Owner);
                string idText = ids.Count > 0 ? string.Join("/", ids) : null;
                if (idText != item.OwnerIds || display != item.Owner)
                {
                    _db.FreeSql.Update<TestItemCatalog>()
                        .Set(t => t.OwnerIds, idText)
                        .Set(t => t.Owner, display)
                        .Where(t => t.Id == item.Id)
                        .ExecuteAffrows();
                }
            }
            if (created > 0 || roleAdded > 0)
            {
                _logger.Info($"按测试项目负责人同步技术员身份: 新建{created}个, 补充身份{roleAdded}个");
            }
            return (created, roleAdded);
        }

        /// <summary>
        /// 负责人姓名 → 账号Id（按显示名/登录名匹配，已存在则取其 uid）
        /// </summary>
        public List<long> ResolveOwnerIds(string ownerNames)
        {
            List<User> users = _db.FreeSql.Select<User>().ToList();
            List<long> ids = [];
            foreach (string name in SplitOwners(ownerNames))
            {
                User user = users.FirstOrDefault(u => string.Equals(u.DisplayName, name, StringComparison.CurrentCultureIgnoreCase))
                    ?? users.FirstOrDefault(u => string.Equals(u.Username, name, StringComparison.OrdinalIgnoreCase));
                if (user != null && !ids.Contains(user.Id))
                {
                    ids.Add(user.Id);
                }
            }
            return ids;
        }

        /// <summary>
        /// 负责人 uid 列表 → 显示名文本（多个以"/"分隔）；uid 缺失时回退到原始姓名文本
        /// </summary>
        public string BuildOwnerDisplay(IEnumerable<long> ownerIds, string fallbackNames = null)
        {
            List<long> ids = ownerIds?.ToList() ?? [];
            if (ids.Count == 0)
            {
                return string.IsNullOrWhiteSpace(fallbackNames) ? null : fallbackNames.Trim();
            }
            Dictionary<long, string> names = _db.FreeSql.Select<User>().ToList().ToDictionary(
                u => u.Id,
                u => string.IsNullOrWhiteSpace(u.DisplayName) ? u.Username : u.DisplayName);
            List<string> display = [];
            foreach (long id in ids)
            {
                if (names.TryGetValue(id, out string name) && !display.Contains(name))
                {
                    display.Add(name);
                }
            }
            return display.Count > 0 ? string.Join("/", display) : fallbackNames;
        }

        /// <summary>
        /// 解析负责人显示文本：优先按 OwnerIds 取当前显示名，其次回退到 Owner 文本
        /// </summary>
        public string ResolveOwnerDisplay(TestItemCatalog item)
        {
            List<long> ids = ParseOwnerIds(item.OwnerIds);
            if (ids.Count > 0)
            {
                string display = BuildOwnerDisplay(ids, item.Owner);
                if (!string.IsNullOrWhiteSpace(display))
                {
                    return display;
                }
            }
            return item.Owner;
        }

        private static List<long> ParseOwnerIds(string ownerIds)
            => string.IsNullOrWhiteSpace(ownerIds)
                ? []
                : SplitOwners(ownerIds).Select(s => long.TryParse(s, out long id) ? id : 0).Where(id => id > 0).ToList();

        /* ###############################  客户管理（整合产品别）  ################################ */

        public List<Customer> GetCustomers()
            => _db.FreeSql.Select<Customer>().OrderBy(c => c.Name).ToList();

        /// <summary>
        /// 获取产品别列表（products 表）
        /// </summary>
        public List<string> GetProducts()
            => _db.FreeSql.Select<Product>().OrderBy(p => p.Name).ToList(p => p.Name);

        /// <summary>
        /// 获取产品别列表（实体，管理窗口用）
        /// </summary>
        public List<Product> GetProductEntities()
            => _db.FreeSql.Select<Product>().OrderBy(p => p.Name).ToList();

        /// <summary>
        /// 新增或更新产品别，返回错误信息；成功返回null
        /// </summary>
        public string SaveProduct(Product product)
        {
            if (string.IsNullOrWhiteSpace(product.Name))
            {
                return "产品别名称不能为空";
            }
            bool exists = _db.FreeSql.Select<Product>()
                .Where(p => p.Name == product.Name && p.Id != product.Id).Any();
            if (exists)
            {
                return $"产品别 [{product.Name}] 已存在";
            }
            if (product.Id == 0)
            {
                _db.FreeSql.Insert(product).ExecuteAffrows();
            }
            else
            {
                _db.FreeSql.Update<Product>().SetSource(product).Where(p => p.Id == product.Id).ExecuteAffrows();
            }
            return null;
        }

        public void DeleteProduct(long id)
            => _db.FreeSql.Delete<Product>().Where(p => p.Id == id).ExecuteAffrows();

        /// <summary>
        /// 新增或更新客户（含代码），返回错误信息；成功返回null
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
        /// 从计划数据（plans 表）同步客户字典
        /// </summary>
        public int SyncCustomersFromPlans()
        {
            // 先修掉历史导入留下的公式文本脏数据（客户栏里的 IF(ISBLANK(...)) 之类）
            PlanExcelService.FixFormulaTextValues(_db);
            List<Plan> plans = _db.FreeSql.Select<Plan>().Where(p => p.Customer != null).ToList();
            int added = 0;
            foreach (IGrouping<string, Plan> group in plans.GroupBy(p => p.Customer))
            {
                string customer = group.Key;
                if (!_db.FreeSql.Select<Customer>().Where(c => c.Name == customer).Any())
                {
                    _db.FreeSql.Insert(new Customer { Name = customer }).ExecuteAffrows();
                    added++;
                }
            }
            _logger.Info($"从计划数据同步客户: 新增{added}个");
            return added;
        }

        /* ###############################  阶段管理  ################################ */

        public List<Stage> GetStages()
            => _db.FreeSql.Select<Stage>().OrderBy(s => s.Id).ToList();

        public string SaveStage(Stage stage)
        {
            if (string.IsNullOrWhiteSpace(stage.Name))
            {
                return "阶段名不能为空";
            }
            bool exists = _db.FreeSql.Select<Stage>()
                .Where(s => s.Name == stage.Name && s.Id != stage.Id).Any();
            if (exists)
            {
                return $"阶段 [{stage.Name}] 已存在";
            }
            if (stage.Id == 0)
            {
                _db.FreeSql.Insert(stage).ExecuteAffrows();
            }
            else
            {
                _db.FreeSql.Update<Stage>().SetSource(stage).Where(s => s.Id == stage.Id).ExecuteAffrows();
            }
            return null;
        }

        public void DeleteStage(long id)
            => _db.FreeSql.Delete<Stage>().Where(s => s.Id == id).ExecuteAffrows();

        private void EnsureDefaultStages()
        {
            try
            {
                string[] defaults = ["MP", "EVT", "DVT", "PVT", "RMA"];
                foreach (string name in defaults)
                {
                    if (!_db.FreeSql.Select<Stage>().Where(s => s.Name == name).Any())
                    {
                        _db.FreeSql.Insert(new Stage { Name = name }).ExecuteAffrows();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "初始化默认阶段失败");
            }
        }

        /* ###############################  测试项目管理  ################################ */

        /// <summary>
        /// 测试项目列表（负责人按 uid 解析出当前显示名填入 OwnerDisplay 供界面展示）
        /// </summary>
        public List<TestItemCatalog> GetTestItems()
        {
            List<TestItemCatalog> items = _db.FreeSql.Select<TestItemCatalog>().OrderBy(t => t.Name).ToList();
            Dictionary<long, string> userNames = _db.FreeSql.Select<User>().ToList().ToDictionary(
                u => u.Id,
                u => string.IsNullOrWhiteSpace(u.DisplayName) ? u.Username : u.DisplayName);
            foreach (TestItemCatalog item in items)
            {
                item.OwnerDisplay = BuildOwnerDisplayFromMap(item, userNames);
            }
            return items;
        }

        /// <summary>
        /// 按名称关键词与历史报告给测试项目自动归类（测试种类），只补"没有种类"和"不确定"的项目，
        /// 用户手工归类过的不动。关键词优先（历史报告里"Conducted EMI Measurement"曾被归到
        /// RELIABILITY TEST，按关键词判才是 EMC），关键词认不出时用计划索引里归并出来的分类。
        /// </summary>
        /// <returns>本次改了归类的条数</returns>
        public int AutoAssignTestItemCategories()
        {
            List<TestItemCatalog> items = _db.FreeSql.Select<TestItemCatalog>().ToList();
            if (items.Count == 0)
            {
                return 0;
            }
            Dictionary<string, string> fromHistory = _db.FreeSql.Select<PlanItemTemplate>().ToList()
                .Where(t => !string.IsNullOrWhiteSpace(t.Category))
                .GroupBy(t => PlanIndexService.NameKey(t.TestItemName))
                .Where(g => !string.IsNullOrEmpty(g.Key))
                .ToDictionary(g => g.Key, g => TestCategories.Normalize(g.Select(t => t.Category)));
            int changed = 0;
            foreach (TestItemCatalog item in items)
            {
                bool uncertain = string.IsNullOrWhiteSpace(item.Category)
                    || string.Equals(item.Category.Trim(), TestCategories.Uncertain, StringComparison.Ordinal);
                if (!uncertain)
                {
                    continue; // 已归好类（可能是用户手工改的），不覆盖
                }
                string keyword = TestCategories.Classify(item.Name);
                string category = keyword;
                if (category == TestCategories.Uncertain)
                {
                    category = fromHistory.TryGetValue(PlanIndexService.NameKey(item.Name), out string history)
                        && !string.IsNullOrEmpty(history)
                        ? history
                        : TestCategories.Uncertain;
                }
                if (string.Equals(item.Category?.Trim(), category, StringComparison.Ordinal))
                {
                    continue;
                }
                item.Category = category;
                _db.FreeSql.Update<TestItemCatalog>()
                    .Set(t => t.Category, category)
                    .Where(t => t.Id == item.Id)
                    .ExecuteAffrows();
                changed++;
            }
            return changed;
        }

        /// <summary>
        /// 保证测试项目都有测试种类（只补空的/不确定的，不动用户手工改过的）
        /// </summary>
        public int EnsureTestItemCategories() => AutoAssignTestItemCategories();

        /* ###############################  测试种类管理  ################################ */

        /// <summary>
        /// 测试种类列表（先保证内置的三种在库里，再把测试项目里已用到的种类补登记）
        /// </summary>
        public List<TestCategory> GetTestCategories()
        {
            EnsureTestCategoriesSeeded();
            return _db.FreeSql.Select<TestCategory>().OrderBy(t => t.Id).ToList();
        }

        /// <summary>
        /// 测试种类名称列表（给下拉框用）
        /// </summary>
        public List<string> GetTestCategoryNames()
            => GetTestCategories().Select(t => t.Name).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct().ToList();

        /// <summary>
        /// 首次使用时写入内置种类，并把测试项目里已有的种类补登记进字典
        /// </summary>
        public int EnsureTestCategoriesSeeded()
        {
            List<TestCategory> existing = _db.FreeSql.Select<TestCategory>().ToList();
            HashSet<string> names = new(existing.Select(t => t.Name ?? ""), StringComparer.OrdinalIgnoreCase);
            int added = 0;
            foreach ((string name, string description) in new[]
            {
                (TestCategories.Reliability, "环境/可靠性类测试（报告里显示 ENVIRONMENT TESTS）"),
                (TestCategories.Emc, "电磁兼容类测试"),
                (TestCategories.Uncertain, "还没归类，请手工指定")
            })
            {
                if (names.Add(name))
                {
                    _db.FreeSql.Insert(new TestCategory { Name = name, Description = description }).ExecuteAffrows();
                    added++;
                }
            }
            // 测试项目里已有但字典里没有的种类（例如索引归并出来的其他写法）也补进来
            foreach (string name in _db.FreeSql.Select<TestItemCatalog>().ToList()
                .Select(i => i.Category?.Trim())
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (names.Add(name))
                {
                    _db.FreeSql.Insert(new TestCategory { Name = name, Description = "由测试项目自动登记" }).ExecuteAffrows();
                    added++;
                }
            }
            return added;
        }

        /// <summary>
        /// 新增或更新测试种类，返回错误信息；成功返回 null
        /// </summary>
        public string SaveTestCategory(TestCategory item)
        {
            if (string.IsNullOrWhiteSpace(item.Name))
            {
                return "测试种类名称不能为空";
            }
            item.Name = item.Name.Trim();
            bool exists = _db.FreeSql.Select<TestCategory>()
                .Where(t => t.Name == item.Name && t.Id != item.Id).Any();
            if (exists)
            {
                return $"测试种类 [{item.Name}] 已存在";
            }
            if (item.Id == 0)
            {
                _db.FreeSql.Insert(item).ExecuteAffrows();
            }
            else
            {
                _db.FreeSql.Update<TestCategory>().SetSource(item).Where(t => t.Id == item.Id).ExecuteAffrows();
            }
            return null;
        }

        /// <summary>
        /// 测试种类改名后，把测试项目上引用旧名字的归类同步过来
        /// </summary>
        public int RenameTestCategoryOnItems(string oldName, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
            {
                return 0;
            }
            return _db.FreeSql.Update<TestItemCatalog>()
                .Set(t => t.Category, newName.Trim())
                .Where(t => t.Category == oldName.Trim())
                .ExecuteAffrows();
        }

        /// <summary>删除测试种类（已被测试项目使用的会把那些项目的种类清空，退回"不确定"）</summary>
        public void DeleteTestCategory(long id)
        {
            TestCategory item = _db.FreeSql.Select<TestCategory>().Where(t => t.Id == id).First();
            if (item == null)
            {
                return;
            }
            _db.FreeSql.Delete<TestCategory>().Where(t => t.Id == id).ExecuteAffrows();
            _db.FreeSql.Update<TestItemCatalog>()
                .Set(t => t.Category, TestCategories.Uncertain)
                .Where(t => t.Category == item.Name)
                .ExecuteAffrows();
        }

        /// <summary>
        /// 按已加载的用户字典解析负责人显示名（优先 OwnerIds，回退 Owner 文本）
        /// </summary>
        private static string BuildOwnerDisplayFromMap(TestItemCatalog item, Dictionary<long, string> userNames)
        {
            List<long> ids = ParseOwnerIds(item.OwnerIds);
            if (ids.Count > 0)
            {
                List<string> display = [];
                foreach (long id in ids)
                {
                    if (userNames.TryGetValue(id, out string name) && !display.Contains(name))
                    {
                        display.Add(name);
                    }
                }
                if (display.Count > 0)
                {
                    return string.Join("/", display);
                }
            }
            return item.Owner;
        }

        /// <summary>
        /// 新增或更新测试项目，返回错误信息；成功返回null。
        /// 负责人（多个以"/"分隔）会立即解析为 uid 锚定（OwnerIds 落库），
        /// 账号不存在时由调用方通过 <see cref="SyncTechniciansFromTestItemOwners"/> 创建。
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
            // 负责人姓名 → uid 锚定；同时把 Owner 文本规范为当前显示名
            List<long> ids = ResolveOwnerIds(item.Owner);
            item.OwnerIds = ids.Count > 0 ? string.Join("/", ids) : null;
            item.Owner = BuildOwnerDisplay(ids, item.Owner);
            item.OwnerDisplay = item.Owner;
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

        /* ###############################  机种映射  ################################ */

        public List<ModelMapping> GetModelMappings()
            => _db.FreeSql.Select<ModelMapping>().OrderBy(m => m.ModelName).ToList();

        public ModelMapping FindModelMapping(string modelName)
            => string.IsNullOrWhiteSpace(modelName) ? null
            : _db.FreeSql.Select<ModelMapping>().Where(m => m.ModelName == modelName.Trim()).First();

        /// <summary>
        /// 插入或更新代码映射（同代码不同名称时更新为最新名称）
        /// </summary>
        private void UpsertCodeMapping(string codeType, string code, string name)
        {
            if (code == null || name == null)
            {
                return;
            }
            CodeMapping existing = _db.FreeSql.Select<CodeMapping>()
                .Where(m => m.CodeType == codeType && m.Code == code).First();
            if (existing == null)
            {
                _db.FreeSql.Insert(new CodeMapping { CodeType = codeType, Code = code, Name = name }).ExecuteAffrows();
            }
            else if (existing.Name != name)
            {
                existing.Name = name;
                _db.FreeSql.Update<CodeMapping>().SetSource(existing).Where(m => m.Id == existing.Id).ExecuteAffrows();
            }
        }

        /// <summary>
        /// 代码规范化：数字代码补足两位前导零（如 "2"→"02"），非数字原样保留
        /// </summary>
        private static string NormalizeCode(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }
            string t = text.Trim();
            return int.TryParse(t, out int n) ? n.ToString("D2") : t;
        }

        /// <summary>
        /// 按机种名称查询客户别：机种名称第 8 位起的 2 位代码 → Cust. Code 表 B/C 列
        /// </summary>
        public string FindCustomerByModel(string modelName)
        {
            string code = ModelToCustomerCode(modelName);
            return code == null ? null
                : _db.FreeSql.Select<CodeMapping>().Where(m => m.CodeType == "C" && m.Code == code).First()?.Name;
        }

        /// <summary>
        /// 按机种名称查询产品别：机种名称开始的 2 位代码 → Cust. Code 表 G/H 列
        /// </summary>
        public string FindProductByModel(string modelName)
        {
            string code = ModelToProductCode(modelName);
            return code == null ? null
                : _db.FreeSql.Select<CodeMapping>().Where(m => m.CodeType == "P" && m.Code == code).First()?.Name;
        }

        /// <summary>
        /// 机种名称 → 客户代码（第 8 位起的 2 位）
        /// </summary>
        public static string ModelToCustomerCode(string modelName)
            => string.IsNullOrWhiteSpace(modelName) || modelName.Trim().Length < 9
                ? null : modelName.Trim().Substring(7, 2);

        /// <summary>
        /// 机种名称 → 产品代码（开始的 2 位）
        /// </summary>
        public static string ModelToProductCode(string modelName)
            => string.IsNullOrWhiteSpace(modelName) || modelName.Trim().Length < 2
                ? null : modelName.Trim().Substring(0, 2);

        public void SetModelMapping(string modelName, string product, string customer)
        {
            if (string.IsNullOrWhiteSpace(modelName))
            {
                return;
            }
            modelName = modelName.Trim();
            ModelMapping existing = _db.FreeSql.Select<ModelMapping>().Where(m => m.ModelName == modelName).First();
            if (existing == null)
            {
                _db.FreeSql.Insert(new ModelMapping { ModelName = modelName, Product = product, Customer = customer }).ExecuteAffrows();
            }
            else
            {
                existing.Product = product ?? existing.Product;
                existing.Customer = customer ?? existing.Customer;
                _db.FreeSql.Update<ModelMapping>().SetSource(existing).Where(m => m.Id == existing.Id).ExecuteAffrows();
            }
        }

        /* ###############################  字典同步（数据源：计划表）  ################################ */

        /// <summary>
        /// 从计划数据（plans 表）同步字典：客户 + 产品别 + 机种映射
        /// </summary>
        public (int customers, int products, int mappings) SyncCatalogsFromPlans()
        {
            List<Plan> plans = _db.FreeSql.Select<Plan>().ToList();
            int cAdded = 0, pAdded = 0, mAdded = 0;
            foreach (IGrouping<string, Plan> group in plans
                .Where(p => !string.IsNullOrWhiteSpace(p.ModelName))
                .GroupBy(p => p.ModelName))
            {
                string customer = group.Select(p => p.Customer).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
                string product = group.Select(p => p.Product).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

                if (customer != null && !_db.FreeSql.Select<Customer>().Where(c => c.Name == customer).Any())
                {
                    _db.FreeSql.Insert(new Customer { Name = customer }).ExecuteAffrows();
                    cAdded++;
                }
                if (product != null && !_db.FreeSql.Select<Product>().Where(p => p.Name == product).Any())
                {
                    _db.FreeSql.Insert(new Product { Name = product }).ExecuteAffrows();
                    pAdded++;
                }
                if (!_db.FreeSql.Select<ModelMapping>().Where(m => m.ModelName == group.Key).Any())
                {
                    _db.FreeSql.Insert(new ModelMapping { ModelName = group.Key, Product = product, Customer = customer }).ExecuteAffrows();
                    mAdded++;
                }
                else
                {
                    SetModelMapping(group.Key, product, customer);
                }
            }
            _logger.Info($"从计划数据同步字典: 客户+{cAdded}, 产品别+{pAdded}, 机种映射+{mAdded}");
            return (cAdded, pAdded, mAdded);
        }

        /// <summary>
        /// 从计划表文件的 Test Items 工作表同步测试项目字典
        /// </summary>
        public int SyncTestItemsFromScheduleFile(string filePath)
        {
            XSSFWorkbook wb = ExcelNpoi.OpenRead(filePath);
            try
            {
                ISheet ws = ExcelNpoi.SheetByName(wb, "Test Items")
                    ?? throw new InvalidDataException("未找到 Test Items 工作表");

                int headerRow = 0;
                int colName = 0, colPeriod = 0, colOwner = 0, colRemark = 0;
                int endCol = ExcelNpoi.LastColumn(ws);
                for (int r = 1; r <= Math.Min(ExcelNpoi.LastRow(ws), 10); r++)
                {
                    for (int c = 1; c <= endCol; c++)
                    {
                        string text = Norm(ExcelNpoi.CellText(ws, r, c));
                        if (text.Contains("試驗項目")) { headerRow = r; colName = c; }
                        else if (text.Contains("試驗時間")) colPeriod = c;
                        else if (text.Contains("負責人")) colOwner = c;
                        else if (text.Contains("備考")) colRemark = c;
                    }
                    if (headerRow > 0) break;
                }
                if (headerRow == 0) throw new InvalidDataException("未找到 Test Items 表头");

                int added = 0, updated = 0;
                int endRow = ExcelNpoi.LastRow(ws);
                for (int r = headerRow + 1; r <= endRow; r++)
                {
                    string name = NullIfEmpty(ExcelNpoi.CellText(ws, r, colName));
                    if (name == null) continue;
                    TestItemCatalog existing = _db.FreeSql.Select<TestItemCatalog>().Where(t => t.Name == name).First();
                    if (existing == null)
                    {
                        _db.FreeSql.Insert(new TestItemCatalog
                        {
                            Name = name,
                            Period = colPeriod > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colPeriod)) : null,
                            Owner = colOwner > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colOwner)) : null,
                            Remark = colRemark > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colRemark)) : null
                        }).ExecuteAffrows();
                        added++;
                    }
                    else
                    {
                        existing.Period = colPeriod > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colPeriod)) : existing.Period;
                        existing.Owner = colOwner > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colOwner)) : existing.Owner;
                        existing.Remark = colRemark > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colRemark)) : existing.Remark;
                        _db.FreeSql.Update<TestItemCatalog>().SetSource(existing).Where(t => t.Id == existing.Id).ExecuteAffrows();
                        updated++;
                    }
                }
                _logger.Info($"从计划表文件同步测试项目: 新增{added}个, 更新{updated}个");
                return added;
            }
            finally
            {
                wb.Close();
            }
        }

        /// <summary>
        /// 从计划表文件的 Cust. Code/Schedule 工作表同步字典：
        /// 客户（B、C 列：Cust. Code→ENDCUSTOMER）、产品别（G、H 列：Code→Product Type）、机种映射
        /// </summary>
        public (int customers, int products, int mappings) SyncCatalogsFromScheduleFile(string filePath)
        {
            XSSFWorkbook wb = ExcelNpoi.OpenRead(filePath);
            try
            {
            int cAdded = 0, pAdded = 0, mAdded = 0;

            // 1. Cust. Code 工作表：B、C 列 → 客户；G、H 列 → 产品别
            ISheet wsCode = ExcelNpoi.SheetByName(wb, "Cust. Code");
            if (wsCode != null)
            {
                int endRow = ExcelNpoi.LastRow(wsCode);
                int endCol = ExcelNpoi.LastColumn(wsCode);
                // 定位表头行（含 ENDCUSTOMER 或 ProductType）
                int headerRow = 0;
                for (int r = 1; r <= Math.Min(endRow, 10); r++)
                {
                    for (int c = 1; c <= endCol; c++)
                    {
                        string header = Norm(ExcelNpoi.CellText(wsCode, r, c));
                        if (header.Contains("ENDCUSTOMER") || header.Contains("Cust.Code"))
                        {
                            headerRow = r;
                            break;
                        }
                    }
                    if (headerRow > 0) break;
                }
                if (headerRow > 0)
                {
                    // 客户：B 列=Cust. Code，C 列=ENDCUSTOMER
                    for (int rr = headerRow + 1; rr <= endRow; rr++)
                    {
                        string code = NormalizeCode(ExcelNpoi.CellText(wsCode, rr, 2));
                        string customer = NullIfEmpty(ExcelNpoi.CellText(wsCode, rr, 3));
                        if (customer != null)
                        {
                            if (!_db.FreeSql.Select<Customer>().Where(c => c.Name == customer).Any())
                            {
                                _db.FreeSql.Insert(new Customer { Name = customer, Code = code }).ExecuteAffrows();
                                cAdded++;
                            }
                            UpsertCodeMapping("C", code, customer);
                        }
                        // 产品别：G 列=Code，H 列=Product Type
                        string productCode = NormalizeCode(ExcelNpoi.CellText(wsCode, rr, 7));
                        string product = NullIfEmpty(ExcelNpoi.CellText(wsCode, rr, 8));
                        if (product != null)
                        {
                            if (!_db.FreeSql.Select<Product>().Where(p => p.Name == product).Any())
                            {
                                _db.FreeSql.Insert(new Product { Name = product, Code = productCode }).ExecuteAffrows();
                                pAdded++;
                            }
                            UpsertCodeMapping("P", productCode, product);
                        }
                    }
                }
            }

            // 2. Schedule 工作表：机种→客户/产品别映射 + 字典补充
            ISheet ws = ExcelNpoi.SheetByName(wb, "Schedule") ?? ExcelNpoi.SheetAt(wb, 0);
            (int headerRowS, Dictionary<string, int> map) = FindScheduleHeader(ws);
            if (headerRowS > 0)
            {
                int colModel = map.TryGetValue("機種名", out int cm) ? cm : 0;
                int colProduct = map.TryGetValue("產品別", out int cp) ? cp : 0;
                int colCustomer = map.TryGetValue("客戶別", out int cc) ? cc : 0;
                int endRow = ExcelNpoi.LastRow(ws);
                for (int r = headerRowS + 1; r <= endRow; r++)
                {
                    string model = colModel > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colModel)) : null;
                    string product = colProduct > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colProduct)) : null;
                    string customer = colCustomer > 0 ? NullIfEmpty(ExcelNpoi.CellText(ws, r, colCustomer)) : null;
                    if (model == null) continue;
                    if (customer != null && !_db.FreeSql.Select<Customer>().Where(c => c.Name == customer).Any())
                    {
                        _db.FreeSql.Insert(new Customer { Name = customer }).ExecuteAffrows();
                        cAdded++;
                    }
                    if (product != null && !_db.FreeSql.Select<Product>().Where(p => p.Name == product).Any())
                    {
                        _db.FreeSql.Insert(new Product { Name = product }).ExecuteAffrows();
                        pAdded++;
                    }
                    if (!_db.FreeSql.Select<ModelMapping>().Where(m => m.ModelName == model).Any())
                    {
                        _db.FreeSql.Insert(new ModelMapping { ModelName = model, Product = product, Customer = customer }).ExecuteAffrows();
                        mAdded++;
                    }
                    else
                    {
                        SetModelMapping(model, product, customer);
                    }
                }
            }
            _logger.Info($"从计划表文件同步字典: 客户+{cAdded}, 产品别+{pAdded}, 机种映射+{mAdded}");
            return (cAdded, pAdded, mAdded);
            }
            finally
            {
                wb.Close();
            }
        }

        private static (int, Dictionary<string, int>) FindScheduleHeader(ISheet ws)
        {
            int endRow = Math.Min(ExcelNpoi.LastRow(ws), 10);
            int endCol = ExcelNpoi.LastColumn(ws);
            for (int r = 1; r <= endRow; r++)
            {
                Dictionary<string, int> map = [];
                for (int c = 1; c <= endCol; c++)
                {
                    string key = Norm(ExcelNpoi.CellText(ws, r, c));
                    if (key.Contains("機種名")) map["機種名"] = c;
                    else if (key.Contains("產品別")) map["產品別"] = c;
                    else if (key.Contains("客戶別")) map["客戶別"] = c;
                }
                if (map.Count > 0) return (r, map);
            }
            return (0, []);
        }

        private static string Norm(string s) => s?.Replace(" ", "").Replace("\n", "").Replace("\r", "") ?? "";
        private static string NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();
    }
}
