using NLog;
using ORT一键报告.Models;
using System;
using System.Globalization;

namespace ORT一键报告.Services
{
    /// <summary>
    /// ORT 看板取数设置 + 平台登录设置的读写（都落在数据库 app_settings 键值对里）。
    /// 用「一项一个键」而不是整段 JSON：app_settings.Value 只有 512 字符，
    /// 链接配置拼成 JSON 会超长，逐项存既安全又方便单点恢复默认。
    /// </summary>
    public class OrtBoardSettingsService
    {
        /* app_settings 键：ORT 看板取数 */
        private const string K_QueryWholeMonth = "ort.query.wholeMonth";
        private const string K_QueryDateFrom = "ort.query.dateFrom";
        private const string K_QueryDateTo = "ort.query.dateTo";
        private const string K_QueryMoType = "ort.query.moType";
        private const string K_EpLoginUrl = "ort.ep.loginUrl";
        private const string K_EpRedirectUrl = "ort.ep.redirectUrl";
        private const string K_EpReportHost = "ort.ep.reportHost";
        private const string K_EpPortalOrigin = "ort.ep.portalOrigin";
        private const string K_EpEntryAccessPath = "ort.ep.entryAccessPath";
        private const string K_EpConfigPath = "ort.ep.configPath";
        private const string K_EpLoadContentPath = "ort.ep.loadContentPath";
        private const string K_EpTargetEntryId = "ort.ep.targetEntryId";
        private const string K_EpWidgetHint = "ort.ep.widgetHint";
        private const string K_EpDeviceUuid = "ort.ep.deviceUuid";
        private const string K_EpUserRegion = "ort.ep.userRegion";
        private const string K_EpUserLanguage = "ort.ep.userLanguage";
        private const string K_EpDashboardWidth = "ort.ep.dashboardWidth";
        private const string K_EpDashboardHeight = "ort.ep.dashboardHeight";
        private const string K_EpLoadRefererHeight = "ort.ep.loadRefererHeight";
        private const string K_EpPageIndex = "ort.ep.pageIndex";
        private const string K_EpDateFromKey = "ort.ep.dateFromKey";
        private const string K_EpDateToKey = "ort.ep.dateToKey";
        private const string K_EpMoTypeKey = "ort.ep.moTypeKey";
        private const string K_EpDateLabelKey = "ort.ep.dateLabelKey";
        private const string K_EpDateLabelText = "ort.ep.dateLabelText";
        private const string K_EpMoTypeLabelKey = "ort.ep.moTypeLabelKey";
        private const string K_EpMoTypeLabelText = "ort.ep.moTypeLabelText";

        /* app_settings 键：平台登录 / 部门身份 */
        private const string K_PortalRegions = "portal.allowedRegions";
        private const string K_PortalEmailDomain = "portal.emailDomain";
        private const string K_PortalAutoRole = "portal.autoEnrollRole";
        private const string K_PortalAutoEnroll = "portal.autoEnroll";

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly AppSettingsService _settings;

        public OrtBoardSettingsService(AppSettingsService settings)
        {
            _settings = settings;
        }

        /* ###############################  ORT 看板取数设置  ################################ */

        /// <summary>
        /// 读取已保存的取数设置（缺的项用默认值补齐，所以旧库/手工改过的库都能读）
        /// </summary>
        public OrtBoardSettings Get()
        {
            OrtBoardEndpoints def = OrtBoardEndpoints.CreateDefault();
            OrtBoardEndpoints ep = new()
            {
                LoginUrl = Text(K_EpLoginUrl, def.LoginUrl),
                RedirectUrl = Text(K_EpRedirectUrl, def.RedirectUrl),
                ReportHost = Text(K_EpReportHost, def.ReportHost),
                PortalOrigin = Text(K_EpPortalOrigin, def.PortalOrigin),
                EntryAccessPath = Text(K_EpEntryAccessPath, def.EntryAccessPath),
                ConfigPath = Text(K_EpConfigPath, def.ConfigPath),
                LoadContentPath = Text(K_EpLoadContentPath, def.LoadContentPath),
                TargetEntryId = Text(K_EpTargetEntryId, ""),
                WidgetHint = Text(K_EpWidgetHint, ""),
                DeviceUuid = Text(K_EpDeviceUuid, ""),
                UserRegion = Text(K_EpUserRegion, def.UserRegion),
                UserLanguage = Text(K_EpUserLanguage, def.UserLanguage),
                DashboardWidth = Int(K_EpDashboardWidth, def.DashboardWidth),
                DashboardHeight = Int(K_EpDashboardHeight, def.DashboardHeight),
                LoadRefererHeight = Int(K_EpLoadRefererHeight, def.LoadRefererHeight),
                PageIndex = Int(K_EpPageIndex, def.PageIndex),
                DateFromKey = Text(K_EpDateFromKey, def.DateFromKey),
                DateToKey = Text(K_EpDateToKey, def.DateToKey),
                MoTypeKey = Text(K_EpMoTypeKey, def.MoTypeKey),
                DateLabelKey = Text(K_EpDateLabelKey, def.DateLabelKey),
                DateLabelText = Text(K_EpDateLabelText, def.DateLabelText),
                MoTypeLabelKey = Text(K_EpMoTypeLabelKey, def.MoTypeLabelKey),
                MoTypeLabelText = Text(K_EpMoTypeLabelText, def.MoTypeLabelText)
            };
            OrtBoardQuery query = new()
            {
                WholeMonth = Bool(K_QueryWholeMonth, true),
                DateFrom = Text(K_QueryDateFrom, ""),
                DateTo = Text(K_QueryDateTo, ""),
                MoType = Text(K_QueryMoType, "")
            };
            if (string.IsNullOrWhiteSpace(query.DateFrom) || string.IsNullOrWhiteSpace(query.DateTo))
            {
                // 第一次用（或早先只存了半截）：按当前整月口径补上
                query.ApplyWholeMonth();
            }
            else if (query.WholeMonth)
            {
                // 勾着「整月」时，日期跟着当前月份走，避免隔月还用上个月的区间
                string expectedFrom = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).ToString("yyyyMMdd");
                if (!string.Equals(query.DateFrom, expectedFrom, StringComparison.Ordinal))
                {
                    query.ApplyWholeMonth();
                }
            }
            return new OrtBoardSettings { Query = query, Endpoints = ep };
        }

        /// <summary>
        /// 保存取数设置（逐项写库）
        /// </summary>
        public void Save(OrtBoardSettings value)
        {
            if (value?.Endpoints == null || value.Query == null)
            {
                return;
            }
            OrtBoardEndpoints ep = value.Endpoints;
            OrtBoardQuery q = value.Query;
            try
            {
                _settings.SetBool(K_QueryWholeMonth, q.WholeMonth);
                _settings.SetText(K_QueryDateFrom, q.DateFrom ?? "");
                _settings.SetText(K_QueryDateTo, q.DateTo ?? "");
                _settings.SetText(K_QueryMoType, q.MoType ?? "");
                _settings.SetText(K_EpLoginUrl, ep.LoginUrl ?? "");
                _settings.SetText(K_EpRedirectUrl, ep.RedirectUrl ?? "");
                _settings.SetText(K_EpReportHost, (ep.ReportHost ?? "").TrimEnd('/'));
                _settings.SetText(K_EpPortalOrigin, ep.PortalOrigin ?? "");
                _settings.SetText(K_EpEntryAccessPath, ep.EntryAccessPath ?? "");
                _settings.SetText(K_EpConfigPath, ep.ConfigPath ?? "");
                _settings.SetText(K_EpLoadContentPath, ep.LoadContentPath ?? "");
                _settings.SetText(K_EpTargetEntryId, (ep.TargetEntryId ?? "").Trim());
                _settings.SetText(K_EpWidgetHint, (ep.WidgetHint ?? "").Trim());
                _settings.SetText(K_EpDeviceUuid, (ep.DeviceUuid ?? "").Trim());
                _settings.SetText(K_EpUserRegion, ep.UserRegion ?? "");
                _settings.SetText(K_EpUserLanguage, ep.UserLanguage ?? "");
                _settings.SetText(K_EpDashboardWidth, ep.DashboardWidth.ToString(CultureInfo.InvariantCulture));
                _settings.SetText(K_EpDashboardHeight, ep.DashboardHeight.ToString(CultureInfo.InvariantCulture));
                _settings.SetText(K_EpLoadRefererHeight, ep.LoadRefererHeight.ToString(CultureInfo.InvariantCulture));
                _settings.SetText(K_EpPageIndex, ep.PageIndex.ToString(CultureInfo.InvariantCulture));
                _settings.SetText(K_EpDateFromKey, ep.DateFromKey ?? "");
                _settings.SetText(K_EpDateToKey, ep.DateToKey ?? "");
                _settings.SetText(K_EpMoTypeKey, ep.MoTypeKey ?? "");
                _settings.SetText(K_EpDateLabelKey, ep.DateLabelKey ?? "");
                _settings.SetText(K_EpDateLabelText, ep.DateLabelText ?? "");
                _settings.SetText(K_EpMoTypeLabelKey, ep.MoTypeLabelKey ?? "");
                _settings.SetText(K_EpMoTypeLabelText, ep.MoTypeLabelText ?? "");
                _logger.Info($"已保存 ORT 看板取数设置：日期 {q.DateFrom}~{q.DateTo}，MoType='{q.MoType}'，"
                    + $"entry={Short(ep.TargetEntryId)}，widget='{ep.WidgetHint}'");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存 ORT 看板取数设置失败");
            }
        }

        /// <summary>
        /// 恢复默认：链接配置回到默认值，TargetEntryId / WidgetHint 一并清空（它们不是默认值）。
        /// 返回恢复后的设置（已写库）。
        /// </summary>
        public OrtBoardSettings ResetAdvancedToDefault()
        {
            OrtBoardSettings current = Get();
            current.Endpoints = current.Endpoints.ResetToDefault();
            Save(current);
            return current;
        }

        /// <summary>本机设备指纹：没配就生成一个固定的 GUID 存下来（用同一个才像同一台机器）</summary>
        public string EnsureDeviceUuid(OrtBoardEndpoints ep)
        {
            if (ep == null)
            {
                return Guid.NewGuid().ToString();
            }
            if (!string.IsNullOrWhiteSpace(ep.DeviceUuid))
            {
                return ep.DeviceUuid.Trim();
            }
            string uuid = Guid.NewGuid().ToString();
            ep.DeviceUuid = uuid;
            _settings.SetText(K_EpDeviceUuid, uuid);
            _logger.Info($"已为平台登录生成本机设备指纹: {uuid}");
            return uuid;
        }

        /* ###############################  平台登录 / 部门身份  ################################ */

        /// <summary>读取平台登录与身份判定设置</summary>
        public PortalAuthSettings GetPortal()
        {
            string role = Text(K_PortalAutoRole, nameof(UserRole.Technician));
            if (!Enum.TryParse(role, out UserRole _))
            {
                role = nameof(UserRole.Technician);
            }
            return new PortalAuthSettings
            {
                AllowedRegions = Text(K_PortalRegions, "DG"),
                EmailDomain = Text(K_PortalEmailDomain, "acbel.com"),
                AutoEnrollRole = role,
                AutoEnroll = Bool(K_PortalAutoEnroll, true)
            };
        }

        /// <summary>保存平台登录与身份判定设置</summary>
        public void SavePortal(PortalAuthSettings value)
        {
            if (value == null)
            {
                return;
            }
            _settings.SetText(K_PortalRegions, value.AllowedRegions ?? "");
            _settings.SetText(K_PortalEmailDomain, value.EmailDomain ?? "acbel.com");
            _settings.SetText(K_PortalAutoRole, value.AutoEnrollRole ?? nameof(UserRole.Technician));
            _settings.SetBool(K_PortalAutoEnroll, value.AutoEnroll);
        }

        /// <summary>邮箱账号 = 登录账号 + 邮箱域</summary>
        public string BuildEmail(string account)
        {
            PortalAuthSettings portal = GetPortal();
            string domain = (portal.EmailDomain ?? "acbel.com").Trim().TrimStart('@');
            return string.IsNullOrWhiteSpace(account) ? null : account.Trim() + "@" + domain;
        }

        /* ###############################  小工具  ################################ */

        private string Text(string key, string fallback)
        {
            string value = _settings.GetText(key, null);
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }

        private int Int(string key, int fallback)
        {
            string value = _settings.GetText(key, null);
            return int.TryParse(value, out int result) && result > 0 ? result : fallback;
        }

        private bool Bool(string key, bool fallback) => _settings.GetBool(key, fallback);

        private static string Short(string text)
            => string.IsNullOrWhiteSpace(text) ? "(未填)" : (text.Length <= 12 ? text : text.Substring(0, 8) + "…");
    }
}
