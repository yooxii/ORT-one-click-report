using Newtonsoft.Json.Linq;
using NLog;
using ORT一键报告.Models;
using System;

namespace ORT一键报告.Services.Sim44
{
    /// <summary>
    /// SIM44（MES 回线工令）服务：统一管理「登录态 + 连接设置 + 凭据」，
    /// 界面只跟这一层打交道，不直接碰桥接进程。
    /// 登录态是进程级的（程序开着就一直有效），密码按用户意愿加密存本机。
    /// </summary>
    public sealed class Sim44Service : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Sim44BridgeClient _client = new();
        private readonly object _gate = new();

        /// <summary>当前是否已登录数据库</summary>
        public bool IsLoggedIn => !string.IsNullOrEmpty(_client.CurrentUser);

        /// <summary>当前登录的数据库账号</summary>
        public string CurrentUser => _client.CurrentUser;

        /// <summary>桥接程序路径（诊断用）</summary>
        public string BridgePath => _client.BridgePath;

        /// <summary>读取本机连接设置（读不到用默认值）</summary>
        public static Sim44Settings ReadSettings()
        {
            try
            {
                return LocalSettingsStore.Read()?.Sim44 ?? new Sim44Settings();
            }
            catch
            {
                return new Sim44Settings();
            }
        }

        /// <summary>保存本机连接设置（只动 Sim44 这一段）</summary>
        public static void SaveSettings(Action<Sim44Settings> change)
        {
            LocalSettingsStore.Update(s =>
            {
                s.Sim44 ??= new Sim44Settings();
                change?.Invoke(s.Sim44);
            });
        }

        /// <summary>环境自检：桥接程序在不在、32 位 Oracle 客户端能不能加载</summary>
        public (bool ok, string message) SelfTest()
        {
            return _client.SelfTest();
        }

        /* ###############################  登录  ################################ */

        /// <summary>
        /// 登录结果分类，界面据此决定提示什么。
        /// </summary>
        public enum LoginOutcome
        {
            /// <summary>成功</summary>
            Success,
            /// <summary>账号或密码不对</summary>
            BadCredential,
            /// <summary>密码已过期，必须先在 MES/数据库那边改</summary>
            PasswordExpired,
            /// <summary>连不上数据库（不在内网 / 服务未开）</summary>
            Network,
            /// <summary>环境问题（缺桥接程序、缺 Oracle 客户端）</summary>
            Environment,
            /// <summary>登录成功，但这个账号访问不到 SIM44 表（缺授权/同义词）</summary>
            NoTable,
            /// <summary>其它错误</summary>
            Other,
        }

        /// <summary>
        /// 用「账号 + 明文密码」登录。remember 为 true 时把密码 DPAPI 加密存到本机。
        /// </summary>
        public (LoginOutcome outcome, string user, string message) Login(string user, string password, bool remember)
        {
            lock (_gate)
            {
                Sim44Settings settings = ReadSettings();
                try
                {
                    string actual = _client.Login(user, password, settings.DataSource, settings.Table);
                    DateTime now = DateTime.Now;
                    Sim44Credential existing = Sim44CredentialStore.Read();
                    bool sameUser = existing != null &&
                                    string.Equals(existing.Username, actual, StringComparison.OrdinalIgnoreCase);
                    if (remember)
                    {
                        // 换了账号就重新算有效期起点（新账号的密码周期未知，按今天算）
                        DateTime? changedAt = sameUser ? existing.LastChangedAt : now;
                        Sim44CredentialStore.Save(actual, password, changedAt, now);
                        SaveSettings(s => { s.LastUser = actual; });
                    }
                    else if (sameUser)
                    {
                        Sim44CredentialStore.Touch(actual, now);
                    }
                    _logger.Info($"SIM44 登录成功：{actual} @ {settings.DataSource}"
                                 + (string.IsNullOrWhiteSpace(_client.LastLoginDiagnostics)
                                     ? ""
                                     : "（表解析：" + _client.LastLoginDiagnostics + "）"));
                    return (LoginOutcome.Success, actual, "");
                }
                catch (Sim44Exception ex)
                {
                    _logger.Warn($"SIM44 登录失败（{ex.Code}）：{ex.Message}"
                                 + (string.IsNullOrWhiteSpace(ex.Detail) ? "" : "｜SQL：" + ex.Detail));
                    LoginOutcome outcome = ex.Code switch
                    {
                        "auth" => LoginOutcome.BadCredential,
                        "expired" => LoginOutcome.PasswordExpired,
                        "net" => LoginOutcome.Network,
                        "notable" => LoginOutcome.NoTable,
                        "nobridge" or "crashed" or "badreply" => LoginOutcome.Environment,
                        _ => LoginOutcome.Other,
                    };
                    return (outcome, "", ex.Message);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "SIM44 登录异常");
                    return (LoginOutcome.Other, "", ex.Message);
                }
            }
        }

        /// <summary>
        /// 用本机保存的凭据自动登录。没有凭据 / 解不开密码时返回 false（调用方再弹登录框）。
        /// </summary>
        public bool TryAutoLogin(out string message)
        {
            message = "";
            Sim44Credential cred = Sim44CredentialStore.Read();
            if (cred == null || string.IsNullOrWhiteSpace(cred.Username))
            {
                return false;
            }
            string password = Sim44CredentialStore.DecryptPassword(cred.PasswordEnc);
            if (password == null)
            {
                message = "本机保存的密码已无法解密（可能换了电脑或 Windows 账号）";
                return false;
            }
            (LoginOutcome outcome, _, string error) = Login(cred.Username, password, remember: true);
            if (outcome == LoginOutcome.Success)
            {
                return true;
            }
            message = error;
            return false;
        }

        /// <summary>退出登录（清连接，不清本机凭据）</summary>
        public void Logout()
        {
            lock (_gate)
            {
                _client.Logout();
            }
        }

        /// <summary>本机凭据状态（用户名 / 上次改密 / 剩余天数）</summary>
        public Sim44CredentialStatus GetCredentialStatus()
        {
            Sim44Credential cred = Sim44CredentialStore.Read();
            Sim44Settings settings = ReadSettings();
            int? daysLeft = Sim44CredentialStore.DaysUntilExpiry(cred?.LastChangedAt);
            return new Sim44CredentialStatus
            {
                HasCredential = cred != null && !string.IsNullOrWhiteSpace(cred.Username),
                User = cred?.Username ?? settings.LastUser ?? "",
                LastChangedAt = cred?.LastChangedAt,
                LastLoginAt = cred?.LastLoginAt,
                DaysLeft = daysLeft,
                HasSavedPassword = cred != null && !string.IsNullOrWhiteSpace(cred.PasswordEnc),
            };
        }

        /// <summary>忘记密码时清掉本机凭据</summary>
        public void ForgetCredential()
        {
            lock (_gate)
            {
                _client.Logout();
            }
            Sim44CredentialStore.Clear();
        }

        /* ###############################  数据操作  ################################ */

        /// <summary>
        /// 查询 SIM44。未登录时抛 <see cref="Sim44Exception"/>（Code=nologin）。
        /// </summary>
        public JArray Query(string moNo, string partNo, string lineName, DateTime? dateFrom, DateTime? dateTo,
            int limit = 200)
        {
            RequireLogin();
            return _client.Query(moNo, partNo, lineName,
                dateFrom?.ToString("yyyy/MM/dd") ?? "", dateTo?.ToString("yyyy/MM/dd") ?? "", limit);
        }

        /// <summary>按工令取现有记录（用于判断「新增/修改」与查重）</summary>
        public JArray GetByMoNo(string moNo)
        {
            RequireLogin();
            return _client.GetByMoNo(moNo);
        }

        /// <summary>新增一笔工令</summary>
        public int Insert(Sim44Row row)
        {
            RequireLogin();
            return _client.Insert(row.ToJson());
        }

        /// <summary>修改一笔工令（按工令定位）</summary>
        public int Update(Sim44Row row)
        {
            RequireLogin();
            return _client.Update(row.ToJson());
        }

        /// <summary>机种名/料号查 BMM04</summary>
        public JArray LookupPart(string text, int limit = 50)
        {
            RequireLogin();
            return _client.LookupPart(text, limit);
        }

        /// <summary>线别查 SIM11</summary>
        public JArray LookupLine(string text, int limit = 50)
        {
            RequireLogin();
            return _client.LookupLine(text, limit);
        }

        /// <summary>
        /// 改数据库密码：成功后会同步更新本机保存的密码与「上次改密日期」（90 天重新计时）。
        /// </summary>
        public (bool ok, string message) ChangePassword(string oldPassword, string newPassword)
        {
            RequireLogin();
            try
            {
                _client.ChangePassword(oldPassword, newPassword);
                Sim44CredentialStore.MarkPasswordChanged(CurrentUser, newPassword, DateTime.Now);
                _logger.Info($"SIM44 账号 {CurrentUser} 已修改数据库密码");
                return (true, "");
            }
            catch (Sim44Exception ex)
            {
                _logger.Warn($"SIM44 改密失败（{ex.Code}）：{ex.Message}");
                return (false, ex.Message);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "SIM44 改密异常");
                return (false, ex.Message);
            }
        }

        private void RequireLogin()
        {
            if (!IsLoggedIn)
            {
                throw new Sim44Exception("nologin", "尚未登录数据库");
            }
        }

        public void Dispose()
        {
            _client.Dispose();
        }
    }

    /// <summary>本机凭据状态（界面显示「密码还有多少天到期」用）</summary>
    public sealed class Sim44CredentialStatus
    {
        public bool HasCredential { get; set; }
        public bool HasSavedPassword { get; set; }
        public string User { get; set; } = "";
        public DateTime? LastChangedAt { get; set; }
        public DateTime? LastLoginAt { get; set; }

        /// <summary>距离 90 天到期还有几天（负数=已超期；null=不知道上次改密时间）</summary>
        public int? DaysLeft { get; set; }

        /// <summary>是否已到期（公司要求 90 天变更）</summary>
        public bool Expired => DaysLeft is < 0;

        /// <summary>是否快到到期（7 天内）</summary>
        public bool ExpiringSoon => DaysLeft is >= 0 and <= 7;
    }

    /// <summary>回线工令（SIM44 一行）的界面模型</summary>
    public sealed class Sim44Row
    {
        public string MoNo { get; set; } = "";
        public string LineNo { get; set; } = "";
        public string LineName { get; set; } = "";
        public string StartDate { get; set; } = "";
        public string PartNo { get; set; } = "";
        public string MoQty { get; set; } = "";
        public string RmaNo { get; set; } = "";
        public string RmaSeq { get; set; } = "";
        public string UpdUser { get; set; } = "";
        public string UpdDate { get; set; } = "";
        public string UpdTime { get; set; } = "";

        public JObject ToJson()
        {
            return new JObject
            {
                ["moNo"] = MoNo ?? "",
                ["lineNo"] = LineNo ?? "",
                ["lineName"] = LineName ?? "",
                ["startDate"] = StartDate ?? "",
                ["partNo"] = PartNo ?? "",
                ["moQty"] = MoQty ?? "",
                ["rmaNo"] = RmaNo ?? "",
                ["rmaSeq"] = RmaSeq ?? "",
                ["updUser"] = UpdUser ?? "",
            };
        }

        public static Sim44Row FromJson(JToken token)
        {
            return new Sim44Row
            {
                MoNo = token.Value<string>("moNo") ?? "",
                LineNo = token.Value<string>("lineNo") ?? "",
                LineName = token.Value<string>("lineName") ?? "",
                StartDate = token.Value<string>("startDate") ?? "",
                PartNo = token.Value<string>("partNo") ?? "",
                MoQty = token.Value<string>("moQty") ?? "",
                RmaNo = token.Value<string>("rmaNo") ?? "",
                RmaSeq = token.Value<string>("rmaSeq") ?? "",
                UpdUser = token.Value<string>("updUser") ?? "",
                UpdDate = token.Value<string>("updDate") ?? "",
                UpdTime = token.Value<string>("updTime") ?? "",
            };
        }

        /// <summary>界面上识别这一行的简短标签</summary>
        public string Summary => MoNo + " / " + (string.IsNullOrEmpty(LineName) ? LineNo : LineName) + " / " + StartDate;
    }
}
