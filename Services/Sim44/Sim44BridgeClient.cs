using NLog;
using Newtonsoft.Json.Linq;
using ORT一键报告.Models;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ORT一键报告.Services.Sim44
{
    /// <summary>
    /// SIM44 桥接进程客户端。
    /// 为什么要多一个进程：MES 的 Oracle 8.1.7 客户端（C:\oracle\ora81）只有 32 位，
    /// 而本程序是 AnyCPU（64 位 Windows 上跑 64 位），进程内根本加载不了 32 位 OCI。
    /// 所以把「连 Oracle」这件事交给随程序发布的 MesBridge.exe（x86，模块 sim44），
    /// 双方用 stdin/stdout 一行一条 JSON 通信。
    /// </summary>
    public sealed class Sim44BridgeClient : IDisposable
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly object _gate = new();

        private Process _process;
        private StreamWriter _stdin;
        private StreamReader _stdout;

        /// <summary>桥接程序文件名（随主程序发布在 MesBridge 子目录）</summary>
        public const string BridgeFileName = "MesBridge.exe";

        /// <summary>桥接模块名：sim44 = MES 回线工令（同一支 exe 里以后还会加别的模块）</summary>
        public const string BridgeArguments = "sim44";

        /// <summary>当前登录的数据库账号（未登录为空）</summary>
        public string CurrentUser { get; private set; }

        /// <summary>桥接程序实际路径（找不到时为预期路径，用于报错提示）</summary>
        public string BridgePath => ResolveBridgePath();

        /// <summary>
        /// 按「程序目录\MesBridge\、程序目录\、本机设置里的覆盖路径」顺序找桥接程序。
        /// </summary>
        private static string ResolveBridgePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            [
                Path.Combine(baseDir, "MesBridge", BridgeFileName),
                Path.Combine(baseDir, BridgeFileName),
            ];
            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            // 允许在设置里指定（比如放到共用目录）
            try
            {
                string custom = LocalSettingsStore.Read()?.Sim44?.BridgePath;
                if (!string.IsNullOrWhiteSpace(custom) && File.Exists(custom))
                {
                    return custom;
                }
            }
            catch
            {
                // 设置读不到不阻塞（下面会给出明确错误）
            }
            return candidates[0];
        }

        /// <summary>
        /// 离线自检：桥接程序在不在、能不能加载 32 位 Oracle 客户端。
        /// 不连数据库、不用账号，随时可点。
        /// </summary>
        public (bool ok, string message) SelfTest()
        {
            string path = ResolveBridgePath();
            if (!File.Exists(path))
            {
                return (false, "找不到桥接程序：" + path);
            }
            try
            {
                ProcessStartInfo psi = new()
                {
                    FileName = path,
                    Arguments = BridgeArguments + " --selftest",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = new UTF8Encoding(false),
                    StandardErrorEncoding = new UTF8Encoding(false),
                };
                using Process process = Process.Start(psi);
                string output = process.StandardOutput.ReadToEnd();
                if (!process.WaitForExit(20000))
                {
                    try { process.Kill(); } catch { }
                    return (false, "桥接程序自检超时");
                }
                JObject result = JObject.Parse(output.Trim());
                bool ok = result.Value<bool?>("ok") == true;
                return (ok, result.Value<string>("message") ?? "");
            }
            catch (Exception ex)
            {
                return (false, "自检失败：" + ex.Message);
            }
        }

        /// <summary>启动桥接进程（已在运行则复用）</summary>
        private void EnsureStarted()
        {
            if (_process != null && !_process.HasExited)
            {
                return;
            }
            string path = ResolveBridgePath();
            if (!File.Exists(path))
            {
                throw new Sim44Exception("nobridge",
                    "找不到 MES 桥接程序（MesBridge.exe）。请确认安装包完整，或重新解压发布包。路径：" + path);
            }
            ProcessStartInfo psi = new()
            {
                FileName = path,
                Arguments = BridgeArguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = new UTF8Encoding(false),
                StandardErrorEncoding = new UTF8Encoding(false),
                WorkingDirectory = Path.GetDirectoryName(path) ?? AppDomain.CurrentDomain.BaseDirectory,
            };
            _process = Process.Start(psi);
            if (_process == null)
            {
                throw new Sim44Exception("nobridge", "无法启动 SIM44 桥接程序：" + path);
            }
            // .NET Framework 的 ProcessStartInfo 没有 StandardInputEncoding，
            // 直接包一层 UTF-8 写入器，保证中文在两端一致（桥接进程读 stdin 也用 UTF-8）。
            _stdin = new StreamWriter(_process.StandardInput.BaseStream, new UTF8Encoding(false))
            {
                AutoFlush = true,
            };
            _stdout = _process.StandardOutput;
            // 错误流不读会阻塞子进程，起个后台任务丢掉
            _ = Task.Run(() =>
            {
                try
                {
                    string line;
                    while ((line = _process.StandardError.ReadLine()) != null)
                    {
                        _logger.Debug("[MesBridge] " + line);
                    }
                }
                catch
                {
                    // 进程退出时的正常异常
                }
            });
            _logger.Info($"SIM44 桥接进程已启动：{path}（PID={_process.Id}）");
        }

        /// <summary>
        /// 发一条请求并等响应（同一条管道，必须串行）。
        /// 请求/响应都会写进日志（密码字段打码），出错时连 SQL 一起记下来，方便离线排查。
        /// </summary>
        private JObject Send(JObject request, int timeoutMs)
        {
            lock (_gate)
            {
                EnsureStarted();
                string line = request.ToString(Newtonsoft.Json.Formatting.None);
                _logger.Debug("→ " + MaskSecrets(line));
                System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();
                _stdin.WriteLine(line);
                Task<string> readTask = _stdout.ReadLineAsync();
                if (!readTask.Wait(timeoutMs))
                {
                    Kill();
                    throw new Sim44Exception("timeout", "数据库无响应（超时 " + timeoutMs / 1000 + " 秒），请检查网络后重试");
                }
                string response = readTask.Result;
                watch.Stop();
                if (response == null)
                {
                    int exitCode = -1;
                    try { exitCode = _process.HasExited ? _process.ExitCode : -1; } catch { }
                    Kill();
                    throw new Sim44Exception("crashed", "SIM44 桥接程序异常退出（exit=" + exitCode + "）");
                }
                JObject reply;
                try
                {
                    reply = JObject.Parse(response);
                }
                catch (Exception ex)
                {
                    throw new Sim44Exception("badreply", "桥接程序返回了无法解析的内容：" + ex.Message);
                }
                if (reply.Value<bool?>("ok") != true)
                {
                    string code = reply.Value<string>("code") ?? "error";
                    string message = reply.Value<string>("message") ?? "未知错误";
                    string sql = reply.Value<string>("sql") ?? "";
                    _logger.Warn($"← 失败〔{code}〕{watch.ElapsedMilliseconds}ms：{message}"
                                 + (string.IsNullOrWhiteSpace(sql) ? "" : "｜SQL：" + sql));
                    throw new Sim44Exception(code, message) { Detail = sql };
                }
                _logger.Debug($"← 成功 {watch.ElapsedMilliseconds}ms：" + MaskSecrets(reply.ToString(Newtonsoft.Json.Formatting.None)));
                return reply;
            }
        }

        /// <summary>日志里把密码字段打码（绝不让密码进日志）</summary>
        private static string MaskSecrets(string json)
        {
            if (string.IsNullOrEmpty(json))
            {
                return "";
            }
            return System.Text.RegularExpressions.Regex.Replace(json,
                "\"(pass|oldPass|newPass)\"\\s*:\\s*\"[^\"]*\"",
                "\"$1\":\"***\"");
        }

        /* ###############################  业务命令  ################################ */

        /// <summary>上次登录时桥接程序报告的「表解析结果」（SIM44 → 实际表名 等），供日志与界面诊断用</summary>
        public string LastLoginDiagnostics { get; private set; } = "";

        /// <summary>登录数据库（成功后桥接进程保持连接）</summary>
        public string Login(string user, string password, string dataSource, string table, int timeoutMs = 30000)
        {
            JObject request = new()
            {
                ["op"] = "login",
                ["user"] = user ?? "",
                ["pass"] = password ?? "",
                ["dataSource"] = dataSource ?? "",
                ["table"] = table ?? "",
            };
            JObject reply = Send(request, timeoutMs);
            CurrentUser = reply.Value<string>("user") ?? user;
            LastLoginDiagnostics = reply["tables"] is JArray tables
                ? string.Join("；", tables.Select(t => t.ToString()))
                : "";
            return CurrentUser;
        }

        /// <summary>退出登录（保留进程，下次登录更快）</summary>
        public void Logout()
        {
            CurrentUser = "";
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _ = Send(new JObject { ["op"] = "logout" }, 10000);
                }
            }
            catch
            {
                // 退出失败无所谓，下次登录会重新建立连接
            }
        }

        /// <summary>按条件查询 SIM44</summary>
        public JArray Query(string moNo, string partNo, string lineName, string dateFrom, string dateTo,
            int limit = 200, int timeoutMs = 30000)
        {
            JObject reply = Send(new JObject
            {
                ["op"] = "query",
                ["moNo"] = moNo ?? "",
                ["partNo"] = partNo ?? "",
                ["lineName"] = lineName ?? "",
                ["dateFrom"] = dateFrom ?? "",
                ["dateTo"] = dateTo ?? "",
                ["limit"] = limit,
            }, timeoutMs);
            return reply["rows"] as JArray ?? [];
        }

        /// <summary>按工令取现有记录</summary>
        public JArray GetByMoNo(string moNo, int timeoutMs = 30000)
        {
            JObject reply = Send(new JObject { ["op"] = "get", ["moNo"] = moNo ?? "" }, timeoutMs);
            return reply["rows"] as JArray ?? [];
        }

        /// <summary>新增一条</summary>
        public int Insert(JObject row, int timeoutMs = 30000)
        {
            JObject reply = Send(new JObject { ["op"] = "insert", ["row"] = row }, timeoutMs);
            return reply.Value<int?>("affected") ?? 0;
        }

        /// <summary>按工令修改一条</summary>
        public int Update(JObject row, int timeoutMs = 30000)
        {
            JObject reply = Send(new JObject { ["op"] = "update", ["row"] = row }, timeoutMs);
            return reply.Value<int?>("affected") ?? 0;
        }

        /// <summary>机种名/料号查 BMM04</summary>
        public JArray LookupPart(string text, int limit = 50, int timeoutMs = 30000)
        {
            JObject reply = Send(new JObject { ["op"] = "lookupPart", ["text"] = text ?? "", ["limit"] = limit },
                timeoutMs);
            return reply["rows"] as JArray ?? [];
        }

        /// <summary>线别查 SIM11</summary>
        public JArray LookupLine(string text, int limit = 50, int timeoutMs = 30000)
        {
            JObject reply = Send(new JObject { ["op"] = "lookupLine", ["text"] = text ?? "", ["limit"] = limit },
                timeoutMs);
            return reply["rows"] as JArray ?? [];
        }

        /// <summary>改数据库密码</summary>
        public void ChangePassword(string oldPassword, string newPassword, int timeoutMs = 30000)
        {
            _ = Send(new JObject
            {
                ["op"] = "changePassword",
                ["oldPass"] = oldPassword ?? "",
                ["newPass"] = newPassword ?? "",
            }, timeoutMs);
        }

        private void Kill()
        {
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _process.Kill();
                }
            }
            catch
            {
                // 杀掉失败不影响后续（下次调用会重新拉起）
            }
            _process?.Dispose();
            _process = null;
            _stdin = null;
            _stdout = null;
            CurrentUser = "";
        }

        public void Dispose()
        {
            lock (_gate)
            {
                try
                {
                    if (_process != null && !_process.HasExited)
                    {
                        _stdin?.WriteLine("{\"op\":\"quit\"}");
                        if (!_process.WaitForExit(2000))
                        {
                            _process.Kill();
                        }
                    }
                }
                catch
                {
                    // 退出流程不抛
                }
                finally
                {
                    _process?.Dispose();
                    _process = null;
                    _stdin = null;
                    _stdout = null;
                    CurrentUser = "";
                }
            }
        }
    }

    /// <summary>SIM44 访问异常（Code 对应桥接程序给出的分类：auth / net / expired / dup / notfound …）</summary>
    public sealed class Sim44Exception : Exception
    {
        public Sim44Exception(string code, string message) : base(message)
        {
            Code = code ?? "";
        }

        public string Code { get; }

        /// <summary>出错时桥接程序回传的 SQL（只用于写日志/排障，不弹给用户）</summary>
        public string Detail { get; set; } = "";

        /// <summary>账号或密码不对</summary>
        public bool IsAuth => Code == "auth" || Code == "expired";

        /// <summary>连不上数据库服务器</summary>
        public bool IsNetwork => Code == "net";

        /// <summary>没有桥接程序 / 客户端环境不对</summary>
        public bool IsEnvironment => Code is "nobridge" or "noadapter" or "crashed" or "badreply" or "timeout";
    }
}
