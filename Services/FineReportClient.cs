using Newtonsoft.Json.Linq;
using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 一次平台登录（SSO）拿到的会话：门户 accessToken + 帆软 fine_auth_token + 平台返回的姓名/区域。
    /// </summary>
    public class FineReportSession
    {
        /// <summary>登录账号（工号）</summary>
        public string Account { get; set; }

        /// <summary>平台返回的姓名（本程序就用它当用户名/显示名）</summary>
        public string DisplayName { get; set; }

        /// <summary>平台返回的区域列表（第一个为主区域）</summary>
        public List<string> Regions { get; set; } = [];

        /// <summary>门户 accessToken（Bearer）</summary>
        public string AccessToken { get; set; }

        /// <summary>帆软 fine_auth_token（调报表接口用，同时也是 Cookie）</summary>
        public string FineToken { get; set; }

        /// <summary>fine_auth_token 过期时间（解析 JWT 的 exp，失败为 MinValue）</summary>
        public DateTime FineTokenExpiry { get; set; }

        /// <summary>
        /// 报表票据没拿到时的原因：登录只校验门户（eip.acbel.com），
        /// 帆软站点维护/故障时票据会缺，登录仍然算成功，等取数时再把这条原因显给用户。
        /// </summary>
        public string ReportTicketError { get; set; }

        /// <summary>主区域</summary>
        public string Region => Regions.FirstOrDefault();

        /// <summary>token 是否还有效（留 5 分钟余量）</summary>
        public bool IsTokenUsable => !string.IsNullOrEmpty(FineToken)
            && FineTokenExpiry > DateTime.Now.AddMinutes(5);
    }

    /// <summary>
    /// 帆软报表取数结果：控件清单 + 还原成行列的单元格矩阵。
    /// </summary>
    public class FineReportTable
    {
        /// <summary>config 返回的全部控件名</summary>
        public List<string> Widgets { get; set; } = [];

        /// <summary>实际取数用的控件名</summary>
        public string Widget { get; set; }

        /// <summary>还原后的行列（行优先，行内按列号顺序，已按合并单元格补齐）</summary>
        public List<List<string>> Rows { get; set; } = [];

        /// <summary>原始响应长度（排查用）</summary>
        public int RawLength { get; set; }

        /// <summary>取数前的表单会话号</summary>
        public string SessionId { get; set; }
    }

    /// <summary>
    /// 公司平台 → 帆软报表的最小客户端：登录换 token、建立表单会话、取数、探测控件名。
    /// 全部地址/请求头都来自 <see cref="OrtBoardEndpoints"/>（设置界面可改，默认即当前可用的一套）。
    /// 这里只做协议，不碰数据库和界面。
    /// </summary>
    public class FineReportClient
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>默认 UA（和浏览器抓包一致，服务端会看）</summary>
        private const string UserAgent =
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/158.0.0.0 Safari/537.36 Edg/158.0.0.0";

        static FineReportClient()
        {
            // .NET 4.8 默认不一定开 TLS1.2，内网 HTTPS 会直接握手失败
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch
            {
                // 老系统上设不了就算了，交给系统默认
            }
        }

        /* ###############################  登录  ################################ */

        /// <summary>
        /// 走完三步登录：门户登录 → 换帆软 ssoToken 跳转地址 → 访问它拿 fine_auth_token。
        /// 失败抛 <see cref="PortalAuthException"/>（带面向用户的提示）。
        /// </summary>
        public async Task<FineReportSession> LoginAsync(string account, string password, string deviceUuid,
            OrtBoardEndpoints ep, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(account) || string.IsNullOrWhiteSpace(password))
            {
                throw new PortalAuthException(PortalLoginStatus.EmptyCredentials, "请输入公司平台账号和密码");
            }
            ep ??= new OrtBoardEndpoints();
            string uuid = string.IsNullOrWhiteSpace(deviceUuid) ? Guid.NewGuid().ToString() : deviceUuid.Trim();
            CookieContainer cookies = new();
            using HttpClient client = BuildPortalClient(cookies, ep);

            // ---- Step 1: 门户登录 ----
            string loginJson = new JObject
            {
                ["userId"] = account.Trim(),
                ["password"] = password,
                ["deviceUuid"] = uuid
            }.ToString(Newtonsoft.Json.Formatting.None);

            _logger.Info($"平台登录开始：账号={account}，deviceUuid={uuid}");
            HttpResponseMessage loginResp;
            string loginText;
            try
            {
                using StringContent content = new(loginJson, Encoding.UTF8, "application/json");
                loginResp = await client.PostAsync(ep.LoginUrl, content, ct).ConfigureAwait(false);
                loginText = await loginResp.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "平台登录请求失败");
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    $"连不上公司平台（{HostOf(ep.LoginUrl)}）：{ex.Message}");
            }

            if (!loginResp.IsSuccessStatusCode)
            {
                _logger.Warn($"平台登录被拒：HTTP {(int)loginResp.StatusCode} {Truncate(loginText, 300)}");
                throw new PortalAuthException(PortalLoginStatus.InvalidCredentials,
                    (int)loginResp.StatusCode == 401 || (int)loginResp.StatusCode == 400
                        ? "账号或密码不正确，或者账号已被锁定"
                        : $"公司平台返回 HTTP {(int)loginResp.StatusCode}，请稍后再试");
            }

            JObject login;
            try
            {
                login = JObject.Parse(loginText);
            }
            catch (Exception)
            {
                _logger.Warn("平台登录响应不是 JSON：" + Truncate(loginText, 300));
                throw new PortalAuthException(PortalLoginStatus.ServiceError, "公司平台返回的内容无法识别，可能被网络代理拦截");
            }

            string access = login["loginAccess"]?.Value<string>();
            string accessToken = login["accessToken"]?.Value<string>();
            if (!string.IsNullOrEmpty(access) && !string.Equals(access, "LOGIN_SUCCESS", StringComparison.OrdinalIgnoreCase))
            {
                _logger.Warn($"平台登录未通过：loginAccess={access}");
                throw new PortalAuthException(PortalLoginStatus.InvalidCredentials,
                    "账号或密码不正确，或者账号已被锁定");
            }
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new PortalAuthException(PortalLoginStatus.ServiceError, "公司平台登录成功但没返回访问令牌");
            }

            FineReportSession session = new()
            {
                Account = account.Trim(),
                DisplayName = (login["userName"]?.Value<string>() ?? "").Trim(),
                AccessToken = accessToken
            };
            if (login["regionList"] is JArray regions)
            {
                session.Regions = regions.Select(r => r?.Value<string>())
                    .Where(r => !string.IsNullOrWhiteSpace(r)).ToList();
            }
            if (session.Regions.Count == 0 && login["roleInfoList"] is JArray roles)
            {
                session.Regions = roles.Select(r => r?["region"]?.Value<string>())
                    .Where(r => !string.IsNullOrWhiteSpace(r)).Distinct().ToList();
            }
            _logger.Info($"平台登录成功：{session.DisplayName}（区域 {string.Join(",", session.Regions)}）");

            // ---- Step 2/3: 换帆软票据（只影响取数，不影响登录）----
            // 登录身份只认门户（eip.acbel.com）：门户通过就代表本人。
            // 帆软报表站点维护/故障时这里会失败——以前会把整次登录带崩，现在只在会话里记下原因，
            // 等用户点「从平台取数」时再把「报表站点暂时不可用」讲清楚。
            try
            {
                string redirectUrl = await RequestSsoRedirectAsync(client, account.Trim(), accessToken, uuid, ep, ct)
                    .ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(redirectUrl) || !redirectUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    throw new PortalAuthException(PortalLoginStatus.ServiceError,
                        "换取报表票据失败：返回的不是跳转地址，" + Truncate(redirectUrl, 200));
                }

                // ---- Step 3: 访问跳转地址，帆软会 Set-Cookie 下发 fine_auth_token ----
                string fineToken = await AcceptSsoTokenAsync(client, cookies, redirectUrl, ep, ct).ConfigureAwait(false);
                session.FineToken = fineToken;
                session.FineTokenExpiry = JwtExpiry(fineToken);
                _logger.Info($"帆软票据获取成功，有效期至 {session.FineTokenExpiry:yyyy-MM-dd HH:mm:ss}");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                session.ReportTicketError = ex is PortalAuthException ? ex.Message : "换取报表票据失败：" + ex.Message;
                _logger.Warn($"登录成功，但没拿到帆软报表票据（不影响登录）：{session.ReportTicketError}");
            }
            return session;
        }

        /// <summary>Step 2：带门户 accessToken 换 ssoToken 跳转地址（空 body + 业务头，少一个都会 500）</summary>
        private async Task<string> RequestSsoRedirectAsync(HttpClient client, string account, string accessToken,
            string deviceUuid, OrtBoardEndpoints ep, CancellationToken ct)
        {
            using HttpRequestMessage request = new(HttpMethod.Post, ep.RedirectUrl);
            request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + accessToken);
            // 服务端会从这几个 Cookie 里取 userId / deviceUuid，缺了直接 500
            request.Headers.TryAddWithoutValidation("Cookie",
                $"userId={account}; deviceUuid={deviceUuid}; dontAskTrust=false; isTrust=false");
            request.Headers.TryAddWithoutValidation("Origin", ep.PortalOrigin);
            request.Headers.TryAddWithoutValidation("Referer", ep.PortalOrigin.TrimEnd('/') + "/");
            request.Headers.TryAddWithoutValidation("Accept", "*/*");
            request.Headers.TryAddWithoutValidation("Accept-Language", "zh-TW,zh;q=0.9,en;q=0.8");
            request.Headers.TryAddWithoutValidation("User-Region", ep.UserRegion);
            request.Headers.TryAddWithoutValidation("User-Language", ep.UserLanguage);
            request.Headers.TryAddWithoutValidation("DeviceUuid", deviceUuid);
            request.Headers.TryAddWithoutValidation("Sec-Fetch-Dest", "empty");
            request.Headers.TryAddWithoutValidation("Sec-Fetch-Mode", "cors");
            request.Headers.TryAddWithoutValidation("Sec-Fetch-Site", "same-origin");
            // 裸 body：空内容却不声明 Content-Type（声明成 JSON 会被服务端判成格式错误 → 500）
            request.Content = new ByteArrayContent(Array.Empty<byte>());

            HttpResponseMessage resp;
            string body;
            try
            {
                resp = await client.SendAsync(request, ct).ConfigureAwait(false);
                body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "换取报表票据失败");
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    $"换取报表票据失败（{HostOf(ep.RedirectUrl)}）：{ex.Message}");
            }

            if (!resp.IsSuccessStatusCode)
            {
                _logger.Warn($"换取报表票据失败：HTTP {(int)resp.StatusCode} {Truncate(body, 300)}");
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    $"换取报表票据失败 HTTP {(int)resp.StatusCode}，" + Truncate(body, 200));
            }
            _logger.Info("已拿到 ssoToken 跳转地址");
            return (body ?? "").Trim();
        }

        /// <summary>Step 3：访问 ssoToken 地址，从 Cookie 容器里取 fine_auth_token</summary>
        private async Task<string> AcceptSsoTokenAsync(HttpClient client, CookieContainer cookies, string url,
            OrtBoardEndpoints ep, CancellationToken ct)
        {
            using HttpRequestMessage request = new(HttpMethod.Get, url);
            request.Headers.TryAddWithoutValidation("Origin", ep.ReportHost);
            request.Headers.TryAddWithoutValidation("Referer", ep.ReportHost.TrimEnd('/') + "/decision");
            try
            {
                using HttpResponseMessage resp = await client.SendAsync(request, ct).ConfigureAwait(false);
                await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "访问 ssoToken 地址失败");
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    $"访问报表站点失败（{HostOf(ep.ReportHost)}）：{ex.Message}");
            }

            Cookie token = cookies.GetCookies(new Uri(ep.ReportHost)).Cast<Cookie>()
                .FirstOrDefault(c => string.Equals(c.Name, "fine_auth_token", StringComparison.OrdinalIgnoreCase));
            if (token == null || string.IsNullOrWhiteSpace(token.Value))
            {
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    "报表站点没有下发 fine_auth_token（报表站点地址或 SSO 流程配置可能不对）");
            }
            return token.Value;
        }

        /* ###############################  取数  ################################ */

        /// <summary>
        /// 建立表单会话并取数：打开看板入口 → 读控件清单 → POST load/content。
        /// </summary>
        public async Task<FineReportTable> FetchAsync(FineReportSession session, OrtBoardEndpoints ep,
            OrtBoardQuery query, CancellationToken ct = default)
        {
            if (session == null || string.IsNullOrWhiteSpace(session.FineToken))
            {
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    "帆软报表站点暂时不可用，还不能取数"
                    + (string.IsNullOrWhiteSpace(session?.ReportTicketError) ? "" : "：" + session.ReportTicketError)
                    + "（登录与看板配置不受影响，等报表恢复后再试）");
            }
            ep ??= new OrtBoardEndpoints();
            query ??= new OrtBoardQuery();
            string entryUrl = ep.BuildEntryUrl();
            if (string.IsNullOrWhiteSpace(entryUrl))
            {
                throw new InvalidOperationException("还没配置目标报表 TargetEntryId（设置 → ORT看板取数 → 高级）");
            }

            CookieContainer cookies = new();
            Uri host = new(ep.ReportHost);
            cookies.Add(host, new Cookie("fine_auth_token", session.FineToken, "/"));
            cookies.Add(host, new Cookie("tenantId", "default", "/"));
            using HttpClient client = BuildReportClient(cookies, session.FineToken);

            // ---- 4a 打开看板入口，建立这张表的会话 ----
            string sessionId;
            try
            {
                using HttpRequestMessage entryRequest = new(HttpMethod.Get, entryUrl);
                entryRequest.Headers.TryAddWithoutValidation("Referer", ep.ReportHost.TrimEnd('/') + "/decision");
                using HttpResponseMessage entryResp = await client.SendAsync(entryRequest, ct).ConfigureAwait(false);
                await entryResp.Content.ReadAsStringAsync().ConfigureAwait(false);
                sessionId = cookies.GetCookies(host).Cast<Cookie>()
                    .FirstOrDefault(c => string.Equals(c.Name, "sessionID", StringComparison.OrdinalIgnoreCase))?.Value;
                if (string.IsNullOrWhiteSpace(sessionId))
                {
                    sessionId = Guid.NewGuid().ToString();
                    cookies.Add(host, new Cookie("sessionID", sessionId, "/"));
                }
                _logger.Info($"看板表单会话：sessionID={sessionId}（entry {ep.TargetEntryId}）");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "打开看板入口失败");
                throw new PortalAuthException(PortalLoginStatus.ServiceError, "打开看板入口失败：" + ex.Message);
            }

            // ---- 4b 读控件清单，挑出真正的数据控件 ----
            List<string> widgets = await RequestWidgetsAsync(client, ep, sessionId, ct).ConfigureAwait(false);
            string widget = PickWidget(widgets, ep.WidgetHint);
            if (string.IsNullOrWhiteSpace(widget))
            {
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    "这张报表没有可用的数据控件，控件清单：" + string.Join(", ", widgets));
            }

            // ---- 4c 取数 ----
            // 大看板会分页（响应里 totalPage > 1）：按页码逐页取回来再拼成一张表，
            // 否则只拿到第一页，后面的行都丢了
            string loadUrl = ep.ReportHost.TrimEnd('/') + ep.LoadContentPath + "?_=" + NowMilliseconds();
            List<List<string>> grid = [];
            int rawLength = 0;
            int startPage = Math.Max(1, ep.PageIndex);
            int totalPage = startPage;
            string previousPage = null;
            string lastText = null;
            for (int page = startPage; page <= totalPage && page < startPage + MaxPages; page++)
            {
                string text = await PostLoadContentAsync(client, ep, loadUrl, widget, query, page, sessionId, ct)
                    .ConfigureAwait(false);
                rawLength += text.Length;
                lastText = text;
                // 服务器忽略 pageIndex、翻回来还是同一页时不重复拼
                if (page > startPage && string.Equals(text, previousPage, StringComparison.Ordinal))
                {
                    break;
                }
                previousPage = text;
                grid.AddRange(ParseGrid(text));
                totalPage = Math.Max(page, ReadTotalPage(text));
            }
            if (grid.Count == 0)
            {
                _logger.Warn($"取数响应里没解析出单元格（原始 {rawLength} 字节）：{Truncate(lastText, 200)}");
            }

            FineReportTable table = new()
            {
                Widgets = widgets,
                Widget = widget,
                SessionId = sessionId,
                RawLength = rawLength,
                Rows = grid
            };
            _logger.Info($"取数成功：控件={widget}，原始 {rawLength} 字节，还原 {table.Rows.Count} 行");
            return table;
        }

        /// <summary>最多翻多少页（防止服务端 totalPage 异常时死循环）</summary>
        private const int MaxPages = 50;

        /// <summary>
        /// POST 一次 load/content 取一页数据，返回响应文本。
        /// HTTP 失败或响应里带 errorCode 时抛 <see cref="PortalAuthException"/>。
        /// </summary>
        private async Task<string> PostLoadContentAsync(HttpClient client, OrtBoardEndpoints ep, string loadUrl,
            string widget, OrtBoardQuery query, int pageIndex, string sessionId, CancellationToken ct)
        {
            string body = BuildLoadBody(ep, widget, query, pageIndex);
            using HttpRequestMessage loadRequest = new(HttpMethod.Post, loadUrl);
            using StringContent content = new(body, Encoding.UTF8);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded") { CharSet = "UTF-8" };
            loadRequest.Content = content;
            loadRequest.Headers.TryAddWithoutValidation("X-Requested-With", "XMLHttpRequest");
            loadRequest.Headers.TryAddWithoutValidation("sessionID", sessionId);
            loadRequest.Headers.TryAddWithoutValidation("Origin", ep.ReportHost);
            loadRequest.Headers.TryAddWithoutValidation("Referer",
                $"{ep.ReportHost.TrimEnd('/')}/decision/v10/entry/access/{ep.TargetEntryId}?width={ep.DashboardWidth}&height={ep.LoadRefererHeight}");

            string text;
            try
            {
                using HttpResponseMessage resp = await client.SendAsync(loadRequest, ct).ConfigureAwait(false);
                text = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode)
                {
                    throw new PortalAuthException(PortalLoginStatus.ServiceError,
                        $"取数失败 HTTP {(int)resp.StatusCode}：" + Truncate(text, 200));
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (PortalAuthException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "取数请求失败");
                throw new PortalAuthException(PortalLoginStatus.ServiceError, "取数请求失败：" + ex.Message);
            }

            string errorCode = TryField(text, "errorCode");
            if (!string.IsNullOrEmpty(errorCode))
            {
                string errorMsg = TryField(text, "errorMsg");
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    $"报表返回错误 {errorCode}：{errorMsg}" + HintForError(errorCode, errorMsg));
            }
            return text;
        }

        /// <summary>
        /// 只读控件清单（设置界面「读取控件列表」用）：打开看板入口 → config。
        /// 这就是 WidgetHint 的动态获取方式。
        /// </summary>
        public async Task<List<string>> ProbeWidgetsAsync(FineReportSession session, OrtBoardEndpoints ep,
            CancellationToken ct = default)
        {
            if (session == null || string.IsNullOrWhiteSpace(session.FineToken))
            {
                throw new PortalAuthException(PortalLoginStatus.ServiceError,
                    "帆软报表站点暂时不可用，还读不到控件列表"
                    + (string.IsNullOrWhiteSpace(session?.ReportTicketError) ? "" : "：" + session.ReportTicketError));
            }
            string entryUrl = ep.BuildEntryUrl();
            if (string.IsNullOrWhiteSpace(entryUrl))
            {
                throw new InvalidOperationException("请先填写目标报表 TargetEntryId");
            }
            CookieContainer cookies = new();
            Uri host = new(ep.ReportHost);
            cookies.Add(host, new Cookie("fine_auth_token", session.FineToken, "/"));
            cookies.Add(host, new Cookie("tenantId", "default", "/"));
            using HttpClient client = BuildReportClient(cookies, session.FineToken);
            using (HttpRequestMessage entryRequest = new(HttpMethod.Get, entryUrl))
            {
                entryRequest.Headers.TryAddWithoutValidation("Referer", ep.ReportHost.TrimEnd('/') + "/decision");
                using HttpResponseMessage entryResp = await client.SendAsync(entryRequest, ct).ConfigureAwait(false);
                await entryResp.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            string sessionId = cookies.GetCookies(host).Cast<Cookie>()
                .FirstOrDefault(c => string.Equals(c.Name, "sessionID", StringComparison.OrdinalIgnoreCase))?.Value
                ?? Guid.NewGuid().ToString();
            return await RequestWidgetsAsync(client, ep, sessionId, ct).ConfigureAwait(false);
        }

        /// <summary>读 config，拿控件名清单</summary>
        private async Task<List<string>> RequestWidgetsAsync(HttpClient client, OrtBoardEndpoints ep,
            string sessionId, CancellationToken ct)
        {
            string url = $"{ep.ReportHost.TrimEnd('/')}{ep.ConfigPath}?_={NowMilliseconds()}"
                + $"&op=form_refresh&cmd=config&sessionID={Uri.EscapeDataString(sessionId)}";
            string body;
            try
            {
                using HttpResponseMessage resp = await client.GetAsync(url, ct).ConfigureAwait(false);
                body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode)
                {
                    throw new PortalAuthException(PortalLoginStatus.ServiceError,
                        $"读取控件列表失败 HTTP {(int)resp.StatusCode}");
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (PortalAuthException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new PortalAuthException(PortalLoginStatus.ServiceError, "读取控件列表失败：" + ex.Message);
            }

            List<string> names = [];
            try
            {
                if (JObject.Parse(body)["reports"] is JArray reports)
                {
                    names = reports.Select(r => r?["widgetName"]?.Value<string>())
                        .Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "解析控件清单失败：" + Truncate(body, 200));
            }
            _logger.Info("控件清单：[" + string.Join(", ", names) + "]");
            return names;
        }

        /// <summary>
        /// 拼 load/content 的请求体：
        /// widgetName=…&amp;pageIndex=1&amp;__parameters__=&lt;双重编码的 JSON&gt;&amp;noCache=&amp;simpleJson=false&amp;arrayJson=false
        /// 注意 __parameters__ 必须**编码两次**（%257B…），其余字段明文，不能整体再过一次表单编码，
        /// 否则 pageIndex 会被吃掉、后端报 "Error page index -1"。
        /// <paramref name="pageIndex"/> 为空时用高级选项里配置的页码（默认第 1 页）。
        /// </summary>
        public static string BuildLoadBody(OrtBoardEndpoints ep, string widget, OrtBoardQuery query, int? pageIndex = null)
        {
            JObject parameters = new()
            {
                [ep.DateLabelKey] = ep.DateLabelText,
                [ep.DateToKey] = query.DateTo,
                [ep.DateFromKey] = query.DateFrom,
                [ep.MoTypeLabelKey] = ep.MoTypeLabelText,
                [ep.MoTypeKey] = query.MoType ?? ""
            };
            string once = Uri.EscapeDataString(parameters.ToString(Newtonsoft.Json.Formatting.None));
            string twice = Uri.EscapeDataString(once);
            return $"widgetName={Uri.EscapeDataString(widget)}"
                + $"&pageIndex={(pageIndex ?? ep.PageIndex)}"
                + $"&__parameters__={twice}"
                + "&noCache="
                + "&simpleJson=false"
                + "&arrayJson=false";
        }

        /// <summary>
        /// 挑数据控件：按提示忽略大小写匹配；没有提示就取第一个不像标题/标签的控件。
        /// （帆软 widgetName 大小写敏感，所以这里必须用 config 返回的真实名字）
        /// </summary>
        public static string PickWidget(IEnumerable<string> widgets, string hint)
        {
            List<string> names = widgets?.Where(n => !string.IsNullOrWhiteSpace(n)).ToList() ?? [];
            if (names.Count == 0)
            {
                return null;
            }
            if (!string.IsNullOrWhiteSpace(hint))
            {
                string hit = names.FirstOrDefault(n => string.Equals(n.Trim(), hint.Trim(), StringComparison.OrdinalIgnoreCase));
                if (hit != null)
                {
                    return hit;
                }
            }
            return names.FirstOrDefault(n => !n.Any(ch => ch > 0x2E80)
                && !n.ToLowerInvariant().Contains("title") && !n.ToLowerInvariant().Contains("label"))
                ?? names[0];
        }

        /// <summary>
        /// 把帆软返回的稀疏单元格矩阵还原成行列文本。
        /// 每个单元格是 { position:{x,y,columnSpan,rowSpan}, value:"{\"type\":\"text\",\"value\":\"…\"}" }，
        /// 合并单元格要按 span 铺开，否则后面的行会整体错位。
        /// <para>
        /// 注意 pageResult 的真实形状是「行数组的数组」：<c>[[单元格…], [单元格…]]</c>，
        /// 不是「页 → 行 → 单元格」的三层结构；早期按三层解析会一个单元格都取不到（表现为「取到数据但 0 行」）。
        /// 这里递归收集带 position 的单元格，嵌套层数多少都能认。
        /// </para>
        /// </summary>
        public static List<List<string>> ParseGrid(string json)
        {
            List<List<string>> rows = [];
            JToken pageResult;
            try
            {
                pageResult = JObject.Parse(json)["pageResult"];
            }
            catch (Exception)
            {
                return rows;
            }
            if (pageResult is not JArray)
            {
                return rows;
            }

            // (行, 列) → 文本；pageResult 的嵌套层数不固定（一页时是「行 → 单元格」，
            // 多页时是「页 → 行 → 单元格」），所以递归找带 position 的单元格，不硬编码层数
            SortedDictionary<int, SortedDictionary<int, string>> cells = [];
            int maxRow = -1, maxCol = -1;
            CollectCells(pageResult, cells, ref maxRow, ref maxCol);

            for (int y = 0; y <= maxRow; y++)
            {
                List<string> line = [];
                cells.TryGetValue(y, out SortedDictionary<int, string> source);
                for (int x = 0; x <= maxCol; x++)
                {
                    string text = null;
                    source?.TryGetValue(x, out text);
                    line.Add((text ?? "").Trim());
                }
                rows.Add(line);
            }
            // 去掉整行都空的尾巴（帆软常在末尾带空行）
            while (rows.Count > 0 && rows[rows.Count - 1].All(string.IsNullOrWhiteSpace))
            {
                rows.RemoveAt(rows.Count - 1);
            }
            return rows;
        }

        /// <summary>
        /// 递归收集单元格：遇到带 position 的对象就按 x/y 落位（合并单元格按 rowSpan/columnSpan 铺开），
        /// 遇到数组就往下走。这样「行数组的数组」和「页 → 行 → 单元格」两种形状都能解析。
        /// </summary>
        private static void CollectCells(JToken token, SortedDictionary<int, SortedDictionary<int, string>> cells,
            ref int maxRow, ref int maxCol)
        {
            if (token is JArray array)
            {
                foreach (JToken child in array)
                {
                    CollectCells(child, cells, ref maxRow, ref maxCol);
                }
                return;
            }
            if (token is not JObject cell)
            {
                return;
            }
            JToken position = cell["position"];
            int y = position?["y"]?.Value<int>() ?? -1;
            int x = position?["x"]?.Value<int>() ?? -1;
            if (y < 0 || x < 0)
            {
                return;   // 不是单元格（比如别处的普通对象）
            }
            int rowSpan = Math.Max(1, position?["rowSpan"]?.Value<int>() ?? 1);
            int colSpan = Math.Max(1, position?["columnSpan"]?.Value<int>() ?? 1);
            string text = CellText(cell["value"]);
            for (int dy = 0; dy < rowSpan; dy++)
            {
                for (int dx = 0; dx < colSpan; dx++)
                {
                    int ry = y + dy, rx = x + dx;
                    if (!cells.TryGetValue(ry, out SortedDictionary<int, string> line))
                    {
                        line = [];
                        cells[ry] = line;
                    }
                    line[rx] = text;
                    if (ry > maxRow)
                    {
                        maxRow = ry;
                    }
                    if (rx > maxCol)
                    {
                        maxCol = rx;
                    }
                }
            }
        }

        /// <summary>
        /// 响应里的总页数（大看板会分页；没有这个字段或认不出就当 1 页）。
        /// </summary>
        public static int ReadTotalPage(string json)
        {
            try
            {
                JToken token = JObject.Parse(json)["totalPage"];
                if (token == null || token.Type == JTokenType.Null)
                {
                    return 1;
                }
                int pages = token.Type == JTokenType.Integer ? token.Value<int>() : 0;
                if (pages <= 0 && !int.TryParse(token.ToString(), out pages))
                {
                    pages = 0;
                }
                return pages > 0 ? pages : 1;
            }
            catch (Exception)
            {
                return 1;
            }
        }

        /// <summary>
        /// 单元格 value 是「JSON 套 JSON」：取里层的 value 字段；不是 JSON 就原样返回。
        /// </summary>
        private static string CellText(JToken value)
        {
            if (value == null || value.Type == JTokenType.Null)
            {
                return "";
            }
            string raw = value.Type == JTokenType.String ? value.Value<string>() : value.ToString();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return "";
            }
            string trimmed = raw.Trim();
            if (!trimmed.StartsWith("{", StringComparison.Ordinal))
            {
                return raw;
            }
            try
            {
                JToken inner = JObject.Parse(trimmed)["value"];
                if (inner == null)
                {
                    return raw;
                }
                return inner.Type == JTokenType.String ? inner.Value<string>() : inner.ToString();
            }
            catch (Exception)
            {
                return raw;
            }
        }

        /* ###############################  工具  ################################ */

        private static HttpClient BuildPortalClient(CookieContainer cookies, OrtBoardEndpoints ep)
        {
            HttpClientHandler handler = new()
            {
                UseCookies = true,
                CookieContainer = cookies,
                AllowAutoRedirect = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };
            HttpClient client = new(handler) { Timeout = TimeSpan.FromSeconds(40) };
            client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", UserAgent);
            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "zh-TW,zh;q=0.9,en;q=0.8");
            client.DefaultRequestHeaders.TryAddWithoutValidation("Referer", ep.PortalOrigin.TrimEnd('/') + "/");
            client.DefaultRequestHeaders.TryAddWithoutValidation("Origin", ep.PortalOrigin);
            return client;
        }

        private static HttpClient BuildReportClient(CookieContainer cookies, string fineToken)
        {
            HttpClientHandler handler = new()
            {
                UseCookies = true,
                CookieContainer = cookies,
                AllowAutoRedirect = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };
            HttpClient client = new(handler) { Timeout = TimeSpan.FromSeconds(90) };
            client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", UserAgent);
            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json, text/plain, */*");
            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "zh-TW,zh;q=0.9,en;q=0.8");
            // 帆软这类接口两头都要带：Cookie 里一份，Authorization 头里也一份
            client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", "Bearer " + fineToken);
            return client;
        }

        /// <summary>解析 JWT 的 exp（失败返回 MinValue）</summary>
        public static DateTime JwtExpiry(string token)
        {
            try
            {
                string payload = token.Split('.')[1].Replace('-', '+').Replace('_', '/');
                switch (payload.Length % 4)
                {
                    case 2: payload += "=="; break;
                    case 3: payload += "="; break;
                }
                string json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                long? exp = JObject.Parse(json)["exp"]?.Value<long?>();
                return exp == null ? DateTime.MinValue : DateTimeOffset.FromUnixTimeSeconds(exp.Value).LocalDateTime;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }

        private static string TryField(string json, string key)
        {
            try
            {
                return JObject.Parse(json)[key]?.Value<string>();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string HintForError(string code, string message)
        {
            if (code == "12400007")
            {
                return "（这张表的会话没建起来：核对目标报表 TargetEntryId 是否对得上这张看板）";
            }
            if (code == "400" && (message ?? "").Contains("page index"))
            {
                return "（pageIndex 异常：核对高级选项里的页码与参数键名）";
            }
            return "";
        }

        private static string HostOf(string url)
        {
            try
            {
                return new Uri(url).Host;
            }
            catch (Exception)
            {
                return url;
            }
        }

        private static long NowMilliseconds() => DateTimeOffset.Now.ToUnixTimeMilliseconds();

        private static string Truncate(string text, int length)
            => string.IsNullOrEmpty(text) ? "(空)" : (text.Length <= length ? text : text.Substring(0, length) + "…");
    }

    /// <summary>
    /// 平台/报表流程的可预期失败：带状态码和给用户看的提示，界面直接显示 Message。
    /// </summary>
    public class PortalAuthException : Exception
    {
        public PortalAuthException(PortalLoginStatus status, string message) : base(message)
        {
            Status = status;
        }

        /// <summary>失败归类</summary>
        public PortalLoginStatus Status { get; }
    }
}
