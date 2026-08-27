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
    /// 管理服务：人员管理（用户+角色）、客户/测试项目/阶段字典管理、机种映射。
    /// 客户表整合客户别与产品别对应关系（Cust. Code 表）。
    /// 字典数据源均在计划表中，可随计划表导入一并同步。
    /// </summary>
    public class AdminService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;

        public AdminService(DatabaseService db)
        {
            _db = db;
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

        public List<TestItemCatalog> GetTestItems()
            => _db.FreeSql.Select<TestItemCatalog>().OrderBy(t => t.Name).ToList();

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
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage package = new(fs);
            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "Test Items")
                ?? throw new InvalidDataException("未找到 Test Items 工作表");

            int headerRow = 0;
            int colName = 0, colPeriod = 0, colOwner = 0, colRemark = 0;
            int endCol = ws.Dimension?.End.Column ?? 0;
            for (int r = 1; r <= Math.Min(ws.Dimension?.End.Row ?? 0, 10); r++)
            {
                for (int c = 1; c <= endCol; c++)
                {
                    string text = Norm(ws.Cells[r, c].Text);
                    if (text.Contains("試驗項目")) { headerRow = r; colName = c; }
                    else if (text.Contains("試驗時間")) colPeriod = c;
                    else if (text.Contains("負責人")) colOwner = c;
                    else if (text.Contains("備考")) colRemark = c;
                }
                if (headerRow > 0) break;
            }
            if (headerRow == 0) throw new InvalidDataException("未找到 Test Items 表头");

            int added = 0, updated = 0;
            int endRow = ws.Dimension?.End.Row ?? 0;
            for (int r = headerRow + 1; r <= endRow; r++)
            {
                string name = NullIfEmpty(ws.Cells[r, colName].Text);
                if (name == null) continue;
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

        /// <summary>
        /// 从计划表文件的 Cust. Code/Schedule 工作表同步字典：
        /// 客户（B、C 列：Cust. Code→ENDCUSTOMER）、产品别（G、H 列：Code→Product Type）、机种映射
        /// </summary>
        public (int customers, int products, int mappings) SyncCatalogsFromScheduleFile(string filePath)
        {
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ExcelPackage package = new(fs);
            int cAdded = 0, pAdded = 0, mAdded = 0;

            // 1. Cust. Code 工作表：B、C 列 → 客户；G、H 列 → 产品别
            ExcelWorksheet wsCode = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "Cust. Code");
            if (wsCode != null)
            {
                int endRow = wsCode.Dimension?.End.Row ?? 0;
                int endCol = wsCode.Dimension?.End.Column ?? 0;
                // 定位表头行（含 ENDCUSTOMER 或 ProductType）
                int headerRow = 0;
                for (int r = 1; r <= Math.Min(endRow, 10); r++)
                {
                    for (int c = 1; c <= endCol; c++)
                    {
                        string header = Norm(wsCode.Cells[r, c].Text);
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
                        string code = NormalizeCode(wsCode.Cells[rr, 2].Text);
                        string customer = NullIfEmpty(wsCode.Cells[rr, 3].Text);
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
                        string productCode = NormalizeCode(wsCode.Cells[rr, 7].Text);
                        string product = NullIfEmpty(wsCode.Cells[rr, 8].Text);
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
            ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "Schedule")
                ?? package.Workbook.Worksheets[0];
            (int headerRowS, Dictionary<string, int> map) = FindScheduleHeader(ws);
            if (headerRowS > 0)
            {
                int colModel = map.TryGetValue("機種名", out int cm) ? cm : 0;
                int colProduct = map.TryGetValue("產品別", out int cp) ? cp : 0;
                int colCustomer = map.TryGetValue("客戶別", out int cc) ? cc : 0;
                int endRow = ws.Dimension?.End.Row ?? 0;
                for (int r = headerRowS + 1; r <= endRow; r++)
                {
                    string model = colModel > 0 ? NullIfEmpty(ws.Cells[r, colModel].Text) : null;
                    string product = colProduct > 0 ? NullIfEmpty(ws.Cells[r, colProduct].Text) : null;
                    string customer = colCustomer > 0 ? NullIfEmpty(ws.Cells[r, colCustomer].Text) : null;
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

        private static (int, Dictionary<string, int>) FindScheduleHeader(ExcelWorksheet ws)
        {
            int endRow = Math.Min(ws.Dimension?.End.Row ?? 0, 10);
            int endCol = ws.Dimension?.End.Column ?? 0;
            for (int r = 1; r <= endRow; r++)
            {
                Dictionary<string, int> map = [];
                for (int c = 1; c <= endCol; c++)
                {
                    string key = Norm(ws.Cells[r, c].Text);
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
