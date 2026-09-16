using FreeSql;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 测试计划维护服务：机种+阶段计划的查询、编辑、排序，测试项模板维护，
    /// 以及"待人工确认差异"的采纳/保留。索引生成本身由 <see cref="PlanIndexService"/> 负责。
    /// </summary>
    public class TestPlanService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;

        public TestPlanService(DatabaseService db, IPermissionService permission)
        {
            _db = db;
            _permission = permission;
        }

        /* ###############################  查询  ################################ */

        /// <summary>
        /// 全部测试计划（带明细数/差异数，按机种+阶段排序）
        /// </summary>
        public List<TestPlan> GetPlans()
        {
            List<TestPlan> plans = _db.FreeSql.Select<TestPlan>()
                .OrderBy(p => p.ModelName)
                .OrderBy(p => p.Stage)
                .ToList();
            List<TestPlanItem> items = _db.FreeSql.Select<TestPlanItem>().ToList();
            foreach (TestPlan plan in plans)
            {
                List<TestPlanItem> own = items.Where(i => i.PlanId == plan.Id).ToList();
                plan.ItemCount = own.Count;
                plan.OverrideCount = own.Count(i => i.HasOverride);
            }
            return plans;
        }

        /// <summary>
        /// 计划的明细（已装配测试项模板，便于直接取"实际生效"文本）
        /// </summary>
        public List<TestPlanItem> GetItems(long planId)
        {
            List<TestPlanItem> items = _db.FreeSql.Select<TestPlanItem>()
                .Where(i => i.PlanId == planId)
                .OrderBy(i => i.OrderNo)
                .ToList();
            AttachTemplates(items);
            return items;
        }

        /// <summary>
        /// 全部测试项模板（按测试项目名排序）
        /// </summary>
        public List<PlanItemTemplate> GetTemplates()
            => _db.FreeSql.Select<PlanItemTemplate>().OrderBy(t => t.TestItemName).ToList();

        /// <summary>
        /// 待人工确认的差异：有差异字段且尚未确认的计划明细（含计划/模板信息）
        /// </summary>
        public List<TestPlanItem> GetPendingDiffs()
        {
            List<TestPlanItem> items = _db.FreeSql.Select<TestPlanItem>()
                .Where(i => i.OverriddenFields != null && i.OverriddenFields != "" && !i.Confirmed)
                .OrderBy(i => i.PlanId)
                .OrderBy(i => i.OrderNo)
                .ToList();
            AttachTemplates(items);
            return items;
        }

        /// <summary>
        /// 计划明细所属计划的显示名（机种 阶段），用于"待确认差异"列表
        /// </summary>
        public Dictionary<long, string> GetPlanTitles()
            => _db.FreeSql.Select<TestPlan>().ToList()
                .ToDictionary(p => p.Id, p => $"{p.ModelName} {PlanStage.Display(p.Stage)}");

        /// <summary>给明细装配模板（一次取全部模板做字典，避免逐条查库）</summary>
        private void AttachTemplates(List<TestPlanItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return;
            }
            Dictionary<long, PlanItemTemplate> templates = _db.FreeSql.Select<PlanItemTemplate>()
                .ToList()
                .ToDictionary(t => t.Id);
            foreach (TestPlanItem item in items)
            {
                if (item.TemplateId.HasValue && templates.TryGetValue(item.TemplateId.Value, out PlanItemTemplate tpl))
                {
                    item.Template = tpl;
                }
                item.RefreshOverriddenFields();
            }
        }

        /* ###############################  计划维护  ################################ */

        /// <summary>
        /// 新增/更新计划（机种+阶段唯一）；返回计划Id
        /// </summary>
        public long SavePlan(TestPlan plan, string user)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            if (string.IsNullOrWhiteSpace(plan.ModelName))
            {
                throw new InvalidOperationException("机种名称不能为空");
            }
            plan.ModelName = plan.ModelName.Trim();
            plan.Stage = PlanStage.Normalize(plan.Stage);
            TestPlan existing = _db.FreeSql.Select<TestPlan>()
                .Where(p => p.ModelName == plan.ModelName && p.Stage == plan.Stage)
                .First();
            if (existing != null && existing.Id != plan.Id)
            {
                throw new InvalidOperationException($"已存在 {plan.ModelName} {PlanStage.Display(plan.Stage)} 的测试计划");
            }
            plan.UpdatedBy = user;
            plan.UpdatedAt = DateTime.Now;
            if (plan.Id <= 0)
            {
                plan.CreatedBy = user;
                plan.CreatedAt = DateTime.Now;
                plan.Source ??= "Manual";
                return _db.FreeSql.Insert(plan).ExecuteIdentity();
            }
            _db.FreeSql.Update<TestPlan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
            return plan.Id;
        }

        /// <summary>删除计划及其明细</summary>
        public void DeletePlan(long planId)
        {
            _db.FreeSql.Delete<TestPlanItem>().Where(i => i.PlanId == planId).ExecuteAffrows();
            _db.FreeSql.Delete<TestPlan>().Where(p => p.Id == planId).ExecuteAffrows();
        }

        /* ###############################  明细维护  ################################ */

        /// <summary>
        /// 新增/更新一条计划明细。
        /// 差异字段与模板相同则自动清空（回到"沿用模板"），并同步差异字段列表。
        /// </summary>
        public long SaveItem(TestPlanItem item, string user)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }
            if (string.IsNullOrWhiteSpace(item.TestItemName))
            {
                throw new InvalidOperationException("测试项目不能为空");
            }
            item.TestItemName = item.TestItemName.Trim();
            item.Category = string.IsNullOrWhiteSpace(item.Category) ? null : item.Category.Trim();
            EnsureTestItemRegistered(item.TestItemName, item.EffectivePeriod, user);
            NormalizeOverrides(item);
            item.UpdatedBy = user;
            item.UpdatedAt = DateTime.Now;
            if (item.Id <= 0)
            {
                if (item.OrderNo <= 0)
                {
                    int max = _db.FreeSql.Select<TestPlanItem>().Where(i => i.PlanId == item.PlanId)
                        .ToList().Select(i => i.OrderNo).DefaultIfEmpty(0).Max();
                    item.OrderNo = max + 1;
                }
                item.RefreshOverriddenFields();
                return _db.FreeSql.Insert(item).ExecuteIdentity();
            }
            item.RefreshOverriddenFields();
            _db.FreeSql.Update<TestPlanItem>().SetSource(item).Where(i => i.Id == item.Id).ExecuteAffrows();
            return item.Id;
        }

        /// <summary>删除一条计划明细</summary>
        public void DeleteItem(long itemId)
            => _db.FreeSql.Delete<TestPlanItem>().Where(i => i.Id == itemId).ExecuteAffrows();

        /// <summary>
        /// 按给定顺序重排计划明细（拖动排序后调用）
        /// </summary>
        public void SaveOrder(long planId, IEnumerable<long> itemIdsInOrder)
        {
            int order = 1;
            foreach (long id in itemIdsInOrder)
            {
                _db.FreeSql.Update<TestPlanItem>()
                    .Set(i => i.OrderNo, order)
                    .Where(i => i.Id == id && i.PlanId == planId)
                    .ExecuteAffrows();
                order++;
            }
        }

        /// <summary>
        /// 把明细的差异清空，改回完全沿用模板
        /// </summary>
        public void ApplyTemplateToItem(long itemId, string user)
        {
            TestPlanItem item = _db.FreeSql.Select<TestPlanItem>().Where(i => i.Id == itemId).First();
            if (item == null)
            {
                return;
            }
            item.SamplingPlan = null;
            item.TestCondition = null;
            item.PassCriterion = null;
            item.Remark = null;
            item.Period = null;
            item.SourceVariants = null;
            item.Confirmed = true;
            item.UpdatedBy = user;
            item.UpdatedAt = DateTime.Now;
            item.RefreshOverriddenFields();
            _db.FreeSql.Update<TestPlanItem>().SetSource(item).Where(i => i.Id == item.Id).ExecuteAffrows();
        }

        /// <summary>
        /// 确认（或取消确认）一条差异：确认后计划索引重建不会覆盖该条明细
        /// </summary>
        public void SetConfirmed(long itemId, bool confirmed, string user)
        {
            TestPlanItem item = _db.FreeSql.Select<TestPlanItem>().Where(i => i.Id == itemId).First();
            if (item == null)
            {
                return;
            }
            item.Confirmed = confirmed;
            item.UpdatedBy = user;
            item.UpdatedAt = DateTime.Now;
            _db.FreeSql.Update<TestPlanItem>().Set(i => i.Confirmed, confirmed)
                .Set(i => i.UpdatedBy, user)
                .Set(i => i.UpdatedAt, DateTime.Now)
                .Where(i => i.Id == itemId)
                .ExecuteAffrows();
        }

        /// <summary>
        /// 把模板的取值写回该明细（"采纳模板"），即清除差异
        /// </summary>
        public void AdoptTemplate(long itemId, string user) => ApplyTemplateToItem(itemId, user);

        /* ###############################  模板维护  ################################ */

        /// <summary>
        /// 新增/更新测试项模板（测试项目名唯一）。用户手工维护的模板会打上 IsManual，索引重建不覆盖文本。
        /// </summary>
        public long SaveTemplate(PlanItemTemplate template, bool manual, string user)
        {
            if (template == null || string.IsNullOrWhiteSpace(template.TestItemName))
            {
                throw new InvalidOperationException("测试项目不能为空");
            }
            template.TestItemName = template.TestItemName.Trim();
            PlanItemTemplate existing = _db.FreeSql.Select<PlanItemTemplate>()
                .Where(t => t.TestItemName == template.TestItemName)
                .First();
            if (existing != null && existing.Id != template.Id)
            {
                throw new InvalidOperationException($"已存在测试项模板：{template.TestItemName}");
            }
            template.IsManual = manual || template.IsManual;
            template.UpdatedBy = user;
            template.UpdatedAt = DateTime.Now;
            if (template.Id <= 0)
            {
                template.CreatedBy = user;
                template.CreatedAt = DateTime.Now;
                return _db.FreeSql.Insert(template).ExecuteIdentity();
            }
            _db.FreeSql.Update<PlanItemTemplate>().SetSource(template).Where(t => t.Id == template.Id).ExecuteAffrows();
            return template.Id;
        }

        /// <summary>删除测试项模板（引用它的明细会失去模板，差异字段仍在，界面可再指定）</summary>
        public void DeleteTemplate(long id)
        {
            _db.FreeSql.Update<TestPlanItem>()
                .Set(i => i.TemplateId, null)
                .Where(i => i.TemplateId == id)
                .ExecuteAffrows();
            _db.FreeSql.Delete<PlanItemTemplate>().Where(t => t.Id == id).ExecuteAffrows();
        }

        /* ###############################  测试项目登记  ################################ */

        /// <summary>
        /// 确保测试项目在"测试项目字典"里已登记（未登记则自动新增），返回是否新增
        /// </summary>
        public bool EnsureTestItemRegistered(string name, string period, string user)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }
            name = name.Trim();
            TestItemCatalog existing = _db.FreeSql.Select<TestItemCatalog>().Where(t => t.Name == name).First();
            if (existing != null)
            {
                return false;
            }
            _db.FreeSql.Insert(new TestItemCatalog
            {
                Name = name,
                Period = period,
                Remark = string.IsNullOrWhiteSpace(user) ? "由测试计划自动登记" : $"由 {user} 的测试计划自动登记"
            }).ExecuteAffrows();
            _logger.Info($"测试项目字典自动登记：{name}");
            return true;
        }

        /// <summary>
        /// 把测试计划里出现过的所有测试项目补登记到字典，返回新增数量
        /// </summary>
        public int SyncTestItemsCatalog(string user)
        {
            int added = 0;
            HashSet<string> known = _db.FreeSql.Select<TestItemCatalog>().ToList()
                .Select(t => t.Name?.Trim())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);
            foreach (TestPlanItem item in _db.FreeSql.Select<TestPlanItem>().ToList())
            {
                string name = item.TestItemName?.Trim();
                if (string.IsNullOrWhiteSpace(name) || !known.Add(name))
                {
                    continue;
                }
                if (EnsureTestItemRegistered(name, item.EffectivePeriod, user))
                {
                    added++;
                }
            }
            return added;
        }

        /// <summary>
        /// 差异字段与模板相同的部分自动清空（保持"能省则省"的存储），
        /// 并同步 <see cref="TestPlanItem.OverriddenFields"/>
        /// </summary>
        private static void NormalizeOverrides(TestPlanItem item)
        {
            PlanItemTemplate tpl = item.Template;
            if (tpl == null && item.TemplateId.HasValue)
            {
                tpl = null; // 调用方未装配模板时不冒险清空
            }
            if (tpl != null)
            {
                if (SameText(item.SamplingPlan, tpl.SamplingPlan)) item.SamplingPlan = null;
                if (SameText(item.TestCondition, tpl.TestCondition)) item.TestCondition = null;
                if (SameText(item.PassCriterion, tpl.PassCriterion)) item.PassCriterion = null;
                if (SameText(item.Remark, tpl.Remark)) item.Remark = null;
                if (SameText(item.Period, tpl.Period)) item.Period = null;
            }
            item.RefreshOverriddenFields();
        }

        /// <summary>归一化后比较两段文本是否等价（空与空视为相同）</summary>
        private static bool SameText(string a, string b)
            => OrtPlanParser.Normalize(a) == OrtPlanParser.Normalize(b);
    }
}
