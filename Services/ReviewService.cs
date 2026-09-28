using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 审核工作流服务：请求提交 / 列表查询 / 通过（应用更改）/ 驳回。
    /// 当前支持"计划表单"类型的 新增/编辑/删除 更改请求（预留"报告"类型）。
    /// </summary>
    public class ReviewService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly MailNotifier _notifier;

        public const string TypePlan = "计划表单";
        public const string StatusPending = "待审核";
        public const string StatusApproved = "已通过";
        public const string StatusRejected = "已驳回";

        public ReviewService(DatabaseService db, MailNotifier notifier)
        {
            _db = db;
            _notifier = notifier;
        }

        /* ###############################  提交  ################################ */

        /// <summary>
        /// 提交计划表单更改请求（普通用户编辑时调用）
        /// </summary>
        public void SubmitPlanRequest(string action, Plan payload, long? targetId, string requester)
        {
            string summary = action switch
            {
                "新增" => $"新增计划: {payload.ModelName ?? "-"} / {payload.JobNo ?? "-"}",
                "编辑" => $"编辑计划(Id={targetId}): {payload.ModelName ?? "-"} / {payload.JobNo ?? "-"}",
                "删除" => $"删除计划(Id={targetId}): {payload?.ModelName ?? "-"} / {payload?.JobNo ?? "-"}",
                _ => $"{action}计划"
            };
            ReviewRequest request = new()
            {
                Type = TypePlan,
                Action = action,
                TargetId = targetId,
                Summary = summary,
                PayloadJson = payload == null ? null : JsonConvert.SerializeObject(payload),
                RequesterName = requester,
                AssigneeName = PickPendingReviewer(),
                Status = StatusPending,
                CreatedAt = DateTime.Now
            };
            _db.FreeSql.Insert(request).ExecuteAffrows();
            _logger.Info($"提交审核请求: {summary} (请求人: {requester}, 待审核人: {request.AssigneeName ?? "-"})");
            _notifier?.NotifyReviewSubmitted(request);
        }

        /// <summary>
        /// 提交领退表更改请求（普通用户编辑时调用），PayloadJson 存 Requisition 序列化
        /// </summary>
        public void SubmitRequisitionRequest(string action, Requisition payload, long? targetId, string requester)
        {
            string summary = action switch
            {
                "新增" => $"新增领退: {payload.ModelName ?? "-"} / {payload.RequisitionNo ?? "-"}",
                "编辑" => $"编辑领退(Id={targetId}): {payload.ModelName ?? "-"} / {payload.RequisitionNo ?? "-"}",
                "报废" => $"报废领退(Id={targetId}): {payload.ModelName ?? "-"} / {payload.RequisitionNo ?? "-"} "
                          + $"数量={payload.ScrapQty ?? "-"} 日期={payload.ScrapDate:yyyy/M/d} "
                          + (payload.ScrapSnFilePath != null ? "序列号=文件" : $"序列号={payload.ScrapSnText?.Count(c => c == '\n') + 1 ?? 0}个"),
                _ => $"{action}领退"
            };
            ReviewRequest request = new()
            {
                Type = "领退表单",
                Action = action,
                TargetId = targetId,
                Summary = summary,
                PayloadJson = payload == null ? null : JsonConvert.SerializeObject(payload),
                RequesterName = requester,
                AssigneeName = PickPendingReviewer(),
                Status = StatusPending,
                CreatedAt = DateTime.Now
            };
            _db.FreeSql.Insert(request).ExecuteAffrows();
            _logger.Info($"提交领退审核请求: {summary} (请求人: {requester}, 待审核人: {request.AssigneeName ?? "-"})");
            _notifier?.NotifyReviewSubmitted(request);
        }

        /// <summary>
        /// 指派"当前待审核人"：在启用且有邮箱的审核员中，选当前待办最少的一位（并列取账号 Id 最小）。
        /// 没有可用审核员时返回 null（发送时退回全部审核员）。
        /// </summary>
        public string PickPendingReviewer()
        {
            try
            {
                List<User> reviewers = _db.FreeSql.Select<User>().Where(u => u.IsActive).ToList();
                List<long> reviewerIds = _db.FreeSql.Select<UserRoleRow>()
                    .Where(r => r.Role == nameof(UserRole.Reviewer)).ToList(r => r.UserId);
                reviewers = reviewers
                    .Where(u => reviewerIds.Contains(u.Id) && !string.IsNullOrWhiteSpace(u.Email))
                    .OrderBy(u => u.Id)
                    .ToList();
                if (reviewers.Count == 0)
                {
                    return null;
                }
                if (reviewers.Count == 1)
                {
                    return reviewers[0].Username;
                }
                // 统计各审核员当前的待审核量（按已指派人名汇总）
                Dictionary<string, int> pending = [];
                foreach (ReviewRequest item in _db.FreeSql.Select<ReviewRequest>().Where(r => r.Status == StatusPending).ToList())
                {
                    string key = item.AssigneeName ?? "";
                    pending[key] = pending.TryGetValue(key, out int count) ? count + 1 : 1;
                }
                return reviewers
                    .OrderBy(u => pending.TryGetValue(u.Username, out int count) ? count : 0)
                    .ThenBy(u => u.Id)
                    .First()
                    .Username;
            }
            catch (Exception ex)
            {
                _logger.Warn($"指派待审核人失败: {ex.Message}");
                return null;
            }
        }

        /* ###############################  查询  ################################ */

        /// <summary>
        /// 请求列表（默认全部，可按状态过滤）
        /// </summary>
        public List<ReviewRequest> GetRequests(string status = null)
        {
            var query = _db.FreeSql.Select<ReviewRequest>();
            if (!string.IsNullOrEmpty(status))
            {
                query = query.Where(r => r.Status == status);
            }
            return query.OrderByDescending(r => r.Id).ToList();
        }

        /// <summary>
        /// 待审核数量（主界面提示用）
        /// </summary>
        public long PendingCount()
            => _db.FreeSql.Select<ReviewRequest>().Where(r => r.Status == StatusPending).Count();

        /* ###############################  审核  ################################ */

        /// <summary>
        /// 通过请求并应用更改；返回错误信息，成功返回null
        /// </summary>
        public string Approve(long requestId, string reviewerName, string comment)
        {
            ReviewRequest request = _db.FreeSql.Select<ReviewRequest>().Where(r => r.Id == requestId).First();
            if (request == null)
            {
                return "请求不存在";
            }
            if (request.Status != StatusPending)
            {
                return $"请求已是 [{request.Status}] 状态，无法重复审核";
            }
            try
            {
                ApplyPlanChange(request);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"应用审核请求(Id={requestId})更改失败");
                return $"应用更改失败: {ex.Message}";
            }
            request.Status = StatusApproved;
            request.ReviewerName = reviewerName;
            request.ReviewComment = comment;
            request.ReviewedAt = DateTime.Now;
            _db.FreeSql.Update<ReviewRequest>().SetSource(request).Where(r => r.Id == requestId).ExecuteAffrows();
            _logger.Info($"审核通过: Id={requestId} ({request.Summary}) 审核人: {reviewerName}");
            _notifier?.NotifyReviewDecided(request, true);
            return null;
        }

        /// <summary>
        /// 驳回请求；返回错误信息，成功返回null
        /// </summary>
        public string Reject(long requestId, string reviewerName, string comment)
        {
            ReviewRequest request = _db.FreeSql.Select<ReviewRequest>().Where(r => r.Id == requestId).First();
            if (request == null)
            {
                return "请求不存在";
            }
            if (request.Status != StatusPending)
            {
                return $"请求已是 [{request.Status}] 状态，无法重复审核";
            }
            request.Status = StatusRejected;
            request.ReviewerName = reviewerName;
            request.ReviewComment = comment;
            request.ReviewedAt = DateTime.Now;
            _db.FreeSql.Update<ReviewRequest>().SetSource(request).Where(r => r.Id == requestId).ExecuteAffrows();
            _logger.Info($"审核驳回: Id={requestId} ({request.Summary}) 审核人: {reviewerName} 意见: {comment}");
            _notifier?.NotifyReviewDecided(request, false);
            return null;
        }

        /// <summary>
        /// 应用计划表单更改：新增→Insert，编辑→Update，删除→Delete
        /// </summary>
        private void ApplyPlanChange(ReviewRequest request)
        {
            if (request.Type == "领退表单")
            {
                ApplyRequisitionChange(request);
                return;
            }
            if (request.Type != TypePlan)
            {
                throw new NotSupportedException($"暂不支持的请求类型: {request.Type}");
            }
            switch (request.Action)
            {
                case "新增":
                    Plan newPlan = JsonConvert.DeserializeObject<Plan>(request.PayloadJson);
                    newPlan.Id = 0;
                    _db.FreeSql.Insert(newPlan).ExecuteAffrows();
                    break;

                case "编辑":
                    if (request.TargetId == null)
                    {
                        throw new InvalidOperationException("编辑请求缺少目标记录Id");
                    }
                    Plan edited = JsonConvert.DeserializeObject<Plan>(request.PayloadJson);
                    edited.Id = request.TargetId.Value;
                    edited.UpdatedAt = DateTime.Now;
                    edited.UpdatedBy = request.ReviewerName;
                    _db.FreeSql.Update<Plan>().SetSource(edited).Where(p => p.Id == edited.Id).ExecuteAffrows();
                    break;

                case "删除":
                    if (request.TargetId == null)
                    {
                        throw new InvalidOperationException("删除请求缺少目标记录Id");
                    }
                    _db.FreeSql.Delete<Plan>().Where(p => p.Id == request.TargetId.Value).ExecuteAffrows();
                    break;

                default:
                    throw new NotSupportedException($"暂不支持的操作: {request.Action}");
            }
        }

        /// <summary>
        /// 应用领退表更改
        /// </summary>
        private void ApplyRequisitionChange(ReviewRequest request)
        {
            switch (request.Action)
            {
                case "新增":
                    Requisition newReq = JsonConvert.DeserializeObject<Requisition>(request.PayloadJson);
                    newReq.Id = 0;
                    _db.FreeSql.Insert(newReq).ExecuteAffrows();
                    break;

                case "编辑":
                    if (request.TargetId == null)
                    {
                        throw new InvalidOperationException("编辑请求缺少目标记录Id");
                    }
                    Requisition edited = JsonConvert.DeserializeObject<Requisition>(request.PayloadJson);
                    edited.Id = request.TargetId.Value;
                    edited.UpdatedAt = DateTime.Now;
                    edited.UpdatedBy = request.ReviewerName;
                    _db.FreeSql.Update<Requisition>().SetSource(edited).Where(r => r.Id == edited.Id).ExecuteAffrows();
                    break;

                case "报废":
                    if (request.TargetId == null)
                    {
                        throw new InvalidOperationException("报废请求缺少目标记录Id");
                    }
                    Requisition scrap = JsonConvert.DeserializeObject<Requisition>(request.PayloadJson);
                    // 通过前复核序列号：领用清单中途被改导致与报废清单不一致时阻止（请求保持待审核）
                    VerifyScrapSnOrThrow(scrap, request.TargetId.Value);
                    // 只更新报废字段与审计字段，不整行覆盖（提交后领用侧可能又改过其他字段）
                    _db.FreeSql.Update<Requisition>()
                        .Set(r => r.ScrapNo, scrap.ScrapNo)
                        .Set(r => r.ScrapQty, scrap.ScrapQty)
                        .Set(r => r.ScrapDate, scrap.ScrapDate)
                        .Set(r => r.ScrapSnText, scrap.ScrapSnText)
                        .Set(r => r.ScrapSnFilePath, scrap.ScrapSnFilePath)
                        .Set(r => r.UpdatedBy, request.ReviewerName)
                        .Set(r => r.UpdatedAt, DateTime.Now)
                        .Where(r => r.Id == request.TargetId.Value)
                        .ExecuteAffrows();
                    break;

                default:
                    throw new NotSupportedException($"暂不支持的操作: {request.Action}");
            }
        }

        /// <summary>
        /// 报废审核通过前的序列号复核：当前领用清单与请求里的报废清单必须完全一致。
        /// 领用清单或报废清单取不到（未登记/附件无法识别）时跳过复核——提交登记时已校验过。
        /// </summary>
        private void VerifyScrapSnOrThrow(Requisition payload, long targetId)
        {
            Requisition current = _db.FreeSql.Select<Requisition>().Where(r => r.Id == targetId).First();
            if (current == null)
            {
                throw new InvalidOperationException($"领退记录不存在(Id={targetId})");
            }
            List<string> scrapList = SnVerification.ParseText(payload.ScrapSnText);
            if (scrapList.Count == 0 && !string.IsNullOrWhiteSpace(payload.ScrapSnFilePath))
            {
                string scrapFile = _db.ResolveAttachmentPath(payload.ScrapSnFilePath);
                if (!string.IsNullOrWhiteSpace(scrapFile) && File.Exists(scrapFile))
                {
                    try
                    {
                        scrapList = SnVerification.ExtractFromFile(scrapFile);
                    }
                    catch
                    {
                        scrapList = [];
                    }
                }
            }
            List<string> reqList = SnVerification.ExtractFromRequisition(current.SN, _db.ResolveAttachmentPath(current.SnFilePath));
            if (reqList.Count == 0 || scrapList.Count == 0)
            {
                return; // 无法复核，不阻挡（提交时已校验）
            }
            SnVerification.SnCompareResult result = SnVerification.Compare(reqList, scrapList);
            if (!result.Ok)
            {
                throw new InvalidOperationException(
                    $"序列号复核未通过：当前领用清单（{result.RequisitionCount}个）与报废清单（{result.ScrapCount}个）不一致，请驳回后重新登记");
            }
        }
    }
}
