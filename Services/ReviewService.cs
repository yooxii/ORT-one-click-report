using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;

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

        public const string TypePlan = "计划表单";
        public const string StatusPending = "待审核";
        public const string StatusApproved = "已通过";
        public const string StatusRejected = "已驳回";

        public ReviewService(DatabaseService db)
        {
            _db = db;
        }

        /* ###############################  提交  ################################ */

        /// <summary>
        /// 提交计划表单更改请求（普通用户编辑时调用）
        /// </summary>
        public void SubmitPlanRequest(string action, Plan payload, long? targetId, string requester)
        {
            string summary = action switch
            {
                "新增" => $"新增计划: {payload.ModelName ?? "-"} / {payload.JobNo ?? payload.RequisitionNo ?? "-"}",
                "编辑" => $"编辑计划(Id={targetId}): {payload.ModelName ?? "-"} / {payload.JobNo ?? payload.RequisitionNo ?? "-"}",
                "删除" => $"删除计划(Id={targetId}): {payload?.ModelName ?? "-"} / {payload?.JobNo ?? payload?.RequisitionNo ?? "-"}",
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
                Status = StatusPending,
                CreatedAt = DateTime.Now
            };
            _db.FreeSql.Insert(request).ExecuteAffrows();
            _logger.Info($"提交审核请求: {summary} (请求人: {requester})");
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
            return null;
        }

        /// <summary>
        /// 应用计划表单更改：新增→Insert，编辑→Update，删除→Delete
        /// </summary>
        private void ApplyPlanChange(ReviewRequest request)
        {
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
    }
}
