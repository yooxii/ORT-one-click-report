using System;
using System.Collections.Generic;
using System.Data.OracleClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace Sim44Bridge
{
    /// <summary>
    /// SIM44 桥接进程（32 位）。主程序用 CreateProcess 拉起它，通过 stdin/stdout 一行一条 JSON 通信：
    ///   {"op":"ping"}、{"op":"login","user":..,"pass":..,"dataSource":..}、{"op":"query",...}、
    ///   {"op":"insert","row":{...}}、{"op":"update","row":{...}}、{"op":"lookupPart","text":..}、
    ///   {"op":"lookupLine","text":..}、{"op":"changePassword","oldPass":..,"newPass":..}、{"op":"quit"}
    /// 单独运行加 --selftest 可以离线检查本机 Oracle 客户端能不能被 32 位进程加载。
    /// </summary>
    internal static class Program
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        private static int Main(string[] args)
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            Console.InputEncoding = Encoding.UTF8;

            if (args.Length > 0 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
            {
                (bool ok, string code, string message) = Sim44Repository.SelfTest();
                Write(new Dictionary<string, object>
                {
                    ["ok"] = ok,
                    ["code"] = code,
                    ["message"] = message,
                    ["bitness"] = Environment.Is64BitProcess ? "x64" : "x86",
                    ["clr"] = Environment.Version.ToString(),
                });
                return ok ? 0 : 3;
            }

            string dataSource = Sim44Repository.DefaultDataSource;
            string table = null;
            Sim44Repository repo = null;
            try
            {
                string line;
                while ((line = Console.In.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    Dictionary<string, object> request;
                    try
                    {
                        request = Json.Deserialize<Dictionary<string, object>>(line);
                    }
                    catch (Exception ex)
                    {
                        WriteError("badjson", "请求不是合法 JSON：" + ex.Message);
                        continue;
                    }
                    if (request == null)
                    {
                        WriteError("badjson", "请求为空");
                        continue;
                    }
                    string op = (BridgeProtocol.Str(request, "op") ?? "").ToLowerInvariant();
                    if (op == "quit" || op == "exit")
                    {
                        break;
                    }
                    try
                    {
                        Handle(op, request, ref repo, ref dataSource, ref table);
                    }
                    catch (BridgeException ex)
                    {
                        WriteError(ex.Code, ex.Message);
                    }
                    catch (OracleException ex)
                    {
                        (string code, string message) = ClassifyOracle(ex);
                        WriteError(code, message);
                    }
                    catch (Exception ex)
                    {
                        WriteError("unknown", ex.GetType().Name + ": " + ex.Message);
                    }
                }
            }
            finally
            {
                repo?.Dispose();
            }
            return 0;
        }

        private static void Handle(string op, Dictionary<string, object> request,
            ref Sim44Repository repo, ref string dataSource, ref string table)
        {
            switch (op)
            {
                case "ping":
                    Write(new Dictionary<string, object>
                    {
                        ["ok"] = true,
                        ["bitness"] = Environment.Is64BitProcess ? "x64" : "x86",
                        ["dataSource"] = dataSource,
                        ["user"] = repo?.CurrentUser ?? "",
                    });
                    return;

                case "selftest":
                {
                    (bool ok, string code, string message) = Sim44Repository.SelfTest();
                    Write(new Dictionary<string, object>
                    {
                        ["ok"] = ok,
                        ["code"] = code,
                        ["message"] = message,
                        ["bitness"] = Environment.Is64BitProcess ? "x64" : "x86",
                    });
                    return;
                }

                case "configure":
                {
                    string ds = BridgeProtocol.Str(request, "dataSource");
                    string tb = BridgeProtocol.Str(request, "table");
                    if (!string.IsNullOrEmpty(ds))
                    {
                        dataSource = ds;
                    }
                    if (!string.IsNullOrEmpty(tb))
                    {
                        table = tb;
                    }
                    repo?.Dispose();
                    repo = new Sim44Repository(dataSource, table);
                    Write(new Dictionary<string, object>
                    {
                        ["ok"] = true,
                        ["dataSource"] = dataSource,
                        ["table"] = table ?? "SIM44",
                    });
                    return;
                }

                case "login":
                {
                    string user = BridgeProtocol.Str(request, "user");
                    string pass = BridgeProtocol.Str(request, "pass");
                    string ds = BridgeProtocol.Str(request, "dataSource");
                    string tb = BridgeProtocol.Str(request, "table");
                    if (!string.IsNullOrEmpty(ds))
                    {
                        dataSource = ds;
                    }
                    if (!string.IsNullOrEmpty(tb))
                    {
                        table = tb;
                    }
                    if (repo == null)
                    {
                        repo = new Sim44Repository(dataSource, table);
                    }
                    string actualUser;
                    try
                    {
                        actualUser = repo.Open(user, pass);
                    }
                    catch (OracleException ex)
                    {
                        (string code, string message) = ClassifyOracle(ex);
                        WriteError(code, message);
                        return;
                    }
                    Write(new Dictionary<string, object>
                    {
                        ["ok"] = true,
                        ["user"] = actualUser,
                        ["dataSource"] = dataSource,
                    });
                    return;
                }

                case "logout":
                    repo?.Dispose();
                    repo = new Sim44Repository(dataSource, table);
                    Write(new Dictionary<string, object> { ["ok"] = true });
                    return;

                case "query":
                {
                    RequireLogin(repo);
                    BridgeProtocol.Sim44Filter filter = new BridgeProtocol.Sim44Filter
                    {
                        MoNo = BridgeProtocol.Str(request, "moNo"),
                        PartNo = BridgeProtocol.Str(request, "partNo"),
                        LineName = BridgeProtocol.Str(request, "lineName"),
                        DateFrom = NormalizeDate(BridgeProtocol.Str(request, "dateFrom")),
                        DateTo = NormalizeDate(BridgeProtocol.Str(request, "dateTo")),
                        Limit = BridgeProtocol.Int(request, "limit", 200),
                    };
                    List<BridgeProtocol.Sim44Row> rows = repo.Query(filter);
                    List<Dictionary<string, object>> payload = new List<Dictionary<string, object>>();
                    foreach (BridgeProtocol.Sim44Row row in rows)
                    {
                        payload.Add(row.ToDictionary());
                    }
                    Write(new Dictionary<string, object>
                    {
                        ["ok"] = true,
                        ["rows"] = payload,
                        ["count"] = payload.Count,
                    });
                    return;
                }

                case "get":
                {
                    RequireLogin(repo);
                    string mo = BridgeProtocol.Str(request, "moNo");
                    List<BridgeProtocol.Sim44Row> rows = repo.GetByMoNo(mo);
                    List<Dictionary<string, object>> payload = new List<Dictionary<string, object>>();
                    foreach (BridgeProtocol.Sim44Row row in rows)
                    {
                        payload.Add(row.ToDictionary());
                    }
                    Write(new Dictionary<string, object>
                    {
                        ["ok"] = true,
                        ["rows"] = payload,
                        ["count"] = payload.Count,
                    });
                    return;
                }

                case "insert":
                {
                    RequireLogin(repo);
                    BridgeProtocol.Sim44Row row = BridgeProtocol.RowFrom(request);
                    row.StartDate = NormalizeDate(row.StartDate);
                    if (string.IsNullOrEmpty(row.UpdUser))
                    {
                        row.UpdUser = repo.CurrentUser;
                    }
                    int affected = repo.Insert(row);
                    Write(new Dictionary<string, object> { ["ok"] = true, ["affected"] = affected });
                    return;
                }

                case "update":
                {
                    RequireLogin(repo);
                    BridgeProtocol.Sim44Row row = BridgeProtocol.RowFrom(request);
                    row.StartDate = NormalizeDate(row.StartDate);
                    if (string.IsNullOrEmpty(row.UpdUser))
                    {
                        row.UpdUser = repo.CurrentUser;
                    }
                    int affected = repo.Update(row);
                    Write(new Dictionary<string, object> { ["ok"] = true, ["affected"] = affected });
                    return;
                }

                case "lookupPart":
                {
                    RequireLogin(repo);
                    List<Dictionary<string, object>> rows = repo.LookupPart(BridgeProtocol.Str(request, "text"),
                        BridgeProtocol.Int(request, "limit", 50));
                    Write(new Dictionary<string, object> { ["ok"] = true, ["rows"] = rows, ["count"] = rows.Count });
                    return;
                }

                case "lookupLine":
                {
                    RequireLogin(repo);
                    List<Dictionary<string, object>> rows = repo.LookupLine(BridgeProtocol.Str(request, "text"),
                        BridgeProtocol.Int(request, "limit", 50));
                    Write(new Dictionary<string, object> { ["ok"] = true, ["rows"] = rows, ["count"] = rows.Count });
                    return;
                }

                case "changePassword":
                {
                    RequireLogin(repo);
                    string oldPass = BridgeProtocol.Str(request, "oldPass");
                    string newPass = BridgeProtocol.Str(request, "newPass");
                    repo.ChangePassword(oldPass, newPass);
                    Write(new Dictionary<string, object> { ["ok"] = true, ["user"] = repo.CurrentUser });
                    return;
                }

                default:
                    WriteError("badop", "不认识的命令：" + op);
                    return;
            }
        }

        private static void RequireLogin(Sim44Repository repo)
        {
            if (repo == null || !repo.IsOpen)
            {
                throw new BridgeException("nologin", "尚未登录数据库，请先输入账号密码");
            }
        }

        /// <summary>日期统一成 MES 认的 yyyy/MM/dd（START_DATE 是字符串列，不是 DATE）</summary>
        private static string NormalizeDate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }
            string[] formats = { "yyyy/M/d", "yyyy/MM/dd", "yyyy-M-d", "yyyy-MM-dd", "yyyyMMdd", "M/d/yyyy" };
            if (DateTime.TryParseExact(text.Trim(), formats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out DateTime date))
            {
                return date.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
            }
            if (DateTime.TryParse(text.Trim(), out date))
            {
                return date.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
            }
            throw new BridgeException("invalid", "日期格式无法识别：" + text);
        }

        /// <summary>把 Oracle 错误翻成主程序能分别处理的小类</summary>
        private static (string code, string message) ClassifyOracle(OracleException ex)
        {
            int number = ex.Code;
            string message = ex.Message.Replace("\r", " ").Replace("\n", " ").Trim();
            if (number == 1017)
            {
                return ("auth", "账号或密码不正确（" + message + "）");
            }
            switch (number)
            {
                case 28000:
                case 28001:
                    return ("expired", "密码已过期，请先修改密码（" + message + "）");
                case 28002:
                    return ("expiring", "密码即将过期，建议尽快修改（" + message + "）");
                case 12154:
                case 12170:
                case 12535:
                case 12541:
                case 12545:
                case 12560:
                    return ("net", "连不上数据库服务器（" + message + "），请确认在公司内网/已连接产线网络");
                case 942:
                    return ("notable", "表不存在或当前账号没有权限（" + message + "）");
                case 1:
                    return ("dup", "数据重复（" + message + "）");
            }
            return ("sql", message);
        }

        private static void WriteError(string code, string message)
        {
            Write(new Dictionary<string, object>
            {
                ["ok"] = false,
                ["code"] = code,
                ["message"] = message,
            });
        }

        private static void Write(Dictionary<string, object> payload)
        {
            string json = Json.Serialize(payload);
            Console.Out.WriteLine(json);
            Console.Out.Flush();
        }
    }
}
