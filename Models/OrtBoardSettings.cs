using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// ORT 看板取数设置（保存在 app_settings：ort.* 键）。
    /// 分两层：
    /// 1) <see cref="Query"/> —— 普通选项：日期区间、MoType，日常只需要改这里；
    /// 2) <see cref="Endpoints"/> —— 高级选项：目标报表（TargetEntryId/WidgetHint）与全部链接配置。
    /// 高级项的默认值 = 已经在用的那套链接配置（<see cref="OrtBoardEndpoints.CreateDefault"/>），
    /// 界面提供「恢复默认」一键回到它们；TargetEntryId / WidgetHint 属于要动态获取的东西，
    /// **故意不给默认值**，只给出参考提示，等实测确认后再由使用人填写。
    /// </summary>
    public class OrtBoardSettings
    {
        /// <summary>普通选项：取数条件</summary>
        public OrtBoardQuery Query { get; set; } = new OrtBoardQuery();

        /// <summary>高级选项：目标报表与全部链接配置</summary>
        public OrtBoardEndpoints Endpoints { get; set; } = new OrtBoardEndpoints();

        /// <summary>深拷贝（界面编辑用，取消时不污染已保存的实例）</summary>
        public OrtBoardSettings Clone() => new()
        {
            Query = new OrtBoardQuery
            {
                WholeMonth = Query.WholeMonth,
                DateFrom = Query.DateFrom,
                DateTo = Query.DateTo,
                MoType = Query.MoType
            },
            Endpoints = Endpoints.Clone()
        };

        /// <summary>取数条件是否具备（日期区间能认出来）</summary>
        public bool IsQueryReady(out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(Query.DateFrom) || string.IsNullOrWhiteSpace(Query.DateTo))
            {
                error = "请先选好取数日期区间（普通选项）";
                return false;
            }
            if (!TryParseYmd(Query.DateFrom, out DateTime from) || !TryParseYmd(Query.DateTo, out DateTime to))
            {
                error = "日期格式应是 yyyyMMdd（例如 20261101）";
                return false;
            }
            if (from > to)
            {
                error = "开始日期不能晚于结束日期";
                return false;
            }
            return true;
        }

        /// <summary>目标报表是否配置好了（高级选项）</summary>
        public bool IsTargetReady(out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(Endpoints.TargetEntryId))
            {
                error = "还没填目标报表 TargetEntryId（高级选项），可以粘贴看板链接自动提取";
                return false;
            }
            return true;
        }

        /// <summary>yyyyMMdd 文本解析</summary>
        public static bool TryParseYmd(string text, out DateTime value)
        {
            value = default;
            text = (text ?? "").Trim();
            if (text.Length != 8)
            {
                return false;
            }
            return DateTime.TryParseExact(text, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out value);
        }

        /// <summary>DateTime → yyyyMMdd</summary>
        public static string ToYmd(DateTime value) => value.ToString("yyyyMMdd");
    }

    /// <summary>
    /// 普通选项：取数条件（日期区间 + MoType）
    /// </summary>
    public class OrtBoardQuery
    {
        /// <summary>整月取数（勾上时日期区间自动取「本月 1 日 ~ 今天」）</summary>
        public bool WholeMonth { get; set; } = true;

        /// <summary>开始日期（yyyyMMdd）</summary>
        public string DateFrom { get; set; } = DateTime.Now.ToString("yyyyMM") + "01";

        /// <summary>结束日期（yyyyMMdd）</summary>
        public string DateTo { get; set; } = DateTime.Now.ToString("yyyyMMdd");

        /// <summary>MoType（类型；空 = 全部）</summary>
        public string MoType { get; set; } = "";

        /// <summary>把「整月」口径套到当前时间（勾了整月时调用）</summary>
        public void ApplyWholeMonth(DateTime? now = null)
        {
            DateTime date = now ?? DateTime.Now;
            DateFrom = new DateTime(date.Year, date.Month, 1).ToString("yyyyMMdd");
            DateTo = date.ToString("yyyyMMdd");
        }
    }

    /// <summary>
    /// 高级选项：目标报表与全部链接配置。
    /// 下面各项的默认值就是当前已经在跑通的那套配置（<see cref="CreateDefault"/>）。
    /// </summary>
    public class OrtBoardEndpoints
    {
        /// <summary>门户登录接口（POST JSON：userId / password / deviceUuid）</summary>
        public string LoginUrl { get; set; } = "https://eip.acbel.com/api/auth/login";

        /// <summary>换取帆软 ssoToken 跳转地址的接口（POST，Authorization: Bearer &lt;accessToken&gt;）</summary>
        public string RedirectUrl { get; set; } = "https://eip.acbel.com/api/redirectToFineReport";

        /// <summary>帆软报表站点根地址（bireport）</summary>
        public string ReportHost { get; set; } = "https://bireport.acbel.com";

        /// <summary>门户 Origin（发请求头用）</summary>
        public string PortalOrigin { get; set; } = "https://eip.acbel.com";

        /// <summary>看板入口路径（{0}=TargetEntryId，{1}=宽度，{2}=高度）</summary>
        public string EntryAccessPath { get; set; } = "/decision/v10/entry/access/{0}?dashboardType=5&width={1}&height={2}";

        /// <summary>表单配置路径（取控件清单）</summary>
        public string ConfigPath { get; set; } = "/decision/view/form";

        /// <summary>取数路径（load/content）</summary>
        public string LoadContentPath { get; set; } = "/decision/view/fit/form/load/content";

        /// <summary>目标报表 entryId。注意：需要动态获取，这里**没有默认值**，由使用人填写/实测后确认</summary>
        public string TargetEntryId { get; set; } = "";

        /// <summary>数据控件名提示（大小写不敏感；留空＝取 config 里第一个非标题控件）。同样没有默认值</summary>
        public string WidgetHint { get; set; } = "";

        /// <summary>设备指纹（留空＝本机自动生成一个并固定下来）</summary>
        public string DeviceUuid { get; set; } = "";

        /// <summary>请求头 User-Region（少了会 500）</summary>
        public string UserRegion { get; set; } = "DG";

        /// <summary>请求头 User-Language（少了会 500）</summary>
        public string UserLanguage { get; set; } = "zh_CN";

        /// <summary>看板宽度参数</summary>
        public int DashboardWidth { get; set; } = 1412;

        /// <summary>看板高度参数</summary>
        public int DashboardHeight { get; set; } = 705;

        /// <summary>取数请求 Referer 里用的高度（浏览器抓包是 162）</summary>
        public int LoadRefererHeight { get; set; } = 162;

        /// <summary>页码（load/content 的 pageIndex）</summary>
        public int PageIndex { get; set; } = 1;

        /// <summary>开始日期参数键</summary>
        public string DateFromKey { get; set; } = "RUN_DATE";

        /// <summary>结束日期参数键</summary>
        public string DateToKey { get; set; } = "RUN_DATE2";

        /// <summary>类型（MoType）参数键</summary>
        public string MoTypeKey { get; set; } = "MO_TYPE";

        /// <summary>日期控件标签键（占位，值一般是「日期:」）</summary>
        public string DateLabelKey { get; set; } = "LABEL0_C";

        /// <summary>日期控件标签文本</summary>
        public string DateLabelText { get; set; } = "日期:";

        /// <summary>类型控件标签键（占位，值一般是「类型:」）</summary>
        public string MoTypeLabelKey { get; set; } = "LABEL0_C_C";

        /// <summary>类型控件标签文本</summary>
        public string MoTypeLabelText { get; set; } = "类型:";

        /// <summary>按默认值重建一份（「恢复默认」用：链接配置回到当前可用的一套）</summary>
        public static OrtBoardEndpoints CreateDefault() => new OrtBoardEndpoints();

        /// <summary>深拷贝</summary>
        public OrtBoardEndpoints Clone() => (OrtBoardEndpoints)MemberwiseClone();

        /// <summary>
        /// 恢复默认：链接配置全部回到默认，
        /// **TargetEntryId / WidgetHint 清空**（它们要动态获取，不算默认值）。
        /// </summary>
        public OrtBoardEndpoints ResetToDefault()
        {
            OrtBoardEndpoints defaults = CreateDefault();
            defaults.TargetEntryId = "";
            defaults.WidgetHint = "";
            return defaults;
        }

        /// <summary>当前入口 URL（TargetEntryId 未填时返回 null）</summary>
        public string BuildEntryUrl()
        {
            if (string.IsNullOrWhiteSpace(TargetEntryId))
            {
                return null;
            }
            string path = string.IsNullOrWhiteSpace(EntryAccessPath)
                ? "/decision/v10/entry/access/{0}"
                : EntryAccessPath;
            try
            {
                return ReportHost.TrimEnd('/') + string.Format(path, TargetEntryId.Trim(), DashboardWidth, DashboardHeight);
            }
            catch (FormatException)
            {
                return ReportHost.TrimEnd('/') + "/decision/v10/entry/access/" + TargetEntryId.Trim();
            }
        }

        /// <summary>
        /// 从一段看板/报表链接里提取 entryId：
        /// 支持 ".../decision/v10/entry/access/{id}?..." 与 ".../entry/access/{id}" 两种写法。
        /// </summary>
        public static string ExtractEntryId(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return null;
            }
            string text = url.Trim();
            int index = text.IndexOf("/entry/access/", StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                return null;
            }
            string rest = text.Substring(index + "/entry/access/".Length);
            int end = rest.IndexOfAny(['?', '&', '#', '/', ' ']);
            string id = (end < 0 ? rest : rest.Substring(0, end)).Trim();
            return string.IsNullOrEmpty(id) ? null : id;
        }
    }
}
