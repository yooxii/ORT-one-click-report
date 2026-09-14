using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 轻量 SMTP 客户端：支持 明文 / STARTTLS(587) / 隐式 SSL(465) 三种安全方式、
    /// 可选忽略服务器证书错误、LOGIN/PLAIN 认证、UTF-8 主题与正文。
    /// （.NET 自带 SmtpClient 不支持隐式 SSL，故自行实现，便于"服务器/种类"完全可配。）
    /// </summary>
    public class SmtpMailClient
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private readonly string _host;
        private readonly int _port;
        private readonly string _security;
        private readonly bool _ignoreCertErrors;
        private readonly bool _useDefaultCredentials;
        private readonly string _username;
        private readonly string _password;
        private readonly int _timeoutMs;

        public SmtpMailClient(MailSettings settings)
        {
            _host = settings.Host;
            _port = settings.Port;
            _security = string.IsNullOrWhiteSpace(settings.Security) ? "None" : settings.Security;
            _ignoreCertErrors = settings.IgnoreCertErrors;
            _useDefaultCredentials = settings.UseDefaultCredentials;
            _username = settings.Username;
            _password = settings.Password;
            _timeoutMs = Math.Max(5, settings.TimeoutSeconds) * 1000;
        }

        private bool IsImplicitSsl => string.Equals(_security, "Ssl", StringComparison.OrdinalIgnoreCase);

        private bool IsStartTls => string.Equals(_security, "StartTls", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// 发送一封邮件（失败抛异常，由调用方记录到 mail_logs）
        /// </summary>
        public void Send(MailMessage message)
        {
            using Session session = Connect();
            string capabilities = session.Command("EHLO " + (Dns.GetHostName() ?? "localhost"), 250, "EHLO");
            if (IsStartTls)
            {
                if (!capabilities.Contains("STARTTLS", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("邮件服务器不支持 STARTTLS，请改用 SSL(465) 或明文(25)方式");
                }
                session.Command("STARTTLS", 220, "STARTTLS");
                session.UpgradeToTls(_host);
                capabilities = session.Command("EHLO " + (Dns.GetHostName() ?? "localhost"), 250, "EHLO(TLS)");
            }
            Authenticate(session, capabilities);
            session.Command($"MAIL FROM:<{message.From}>", 250, "MAIL FROM");
            foreach (string to in message.To)
            {
                session.Command($"RCPT TO:<{to}>", 250, "RCPT TO " + to);
            }
            foreach (string cc in message.Cc)
            {
                session.Command($"RCPT TO:<{cc}>", 250, "RCPT TO(Cc) " + cc);
            }
            session.Command("DATA", 354, "DATA");
            session.WriteRaw(BuildData(message));
            session.Expect(250, "邮件内容");
            session.TryQuit();
            _logger.Info($"邮件已发送至 {string.Join(";", message.To)}，主题: {message.Subject}");
        }

        /// <summary>
        /// 连接/加密/认证自检（不发送邮件），返回 null 表示正常
        /// </summary>
        public string TestConnection()
        {
            try
            {
                using Session session = Connect();
                string capabilities = session.Command("EHLO " + (Dns.GetHostName() ?? "localhost"), 250, "EHLO");
                if (IsStartTls)
                {
                    if (!capabilities.Contains("STARTTLS", StringComparison.OrdinalIgnoreCase))
                    {
                        return "服务器不支持 STARTTLS";
                    }
                    session.Command("STARTTLS", 220, "STARTTLS");
                    session.UpgradeToTls(_host);
                    capabilities = session.Command("EHLO " + (Dns.GetHostName() ?? "localhost"), 250, "EHLO(TLS)");
                }
                Authenticate(session, capabilities);
                session.TryQuit();
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private Session Connect()
        {
            TcpClient client = new();
            try
            {
                client.ReceiveTimeout = _timeoutMs;
                client.SendTimeout = _timeoutMs;
                IAsyncResult connect = client.BeginConnect(_host, _port, null, null);
                if (!connect.AsyncWaitHandle.WaitOne(_timeoutMs))
                {
                    throw new TimeoutException($"连接邮件服务器超时: {_host}:{_port}");
                }
                client.EndConnect(connect);
                Stream stream = client.GetStream();
                stream.ReadTimeout = _timeoutMs;
                stream.WriteTimeout = _timeoutMs;
                Session session = new(stream, client, _timeoutMs, _ignoreCertErrors);
                if (IsImplicitSsl)
                {
                    session.UpgradeToTls(_host);   // 465：连上即 TLS
                }
                session.Expect(220, "服务器问候");
                return session;
            }
            catch
            {
                client.Dispose();
                throw;
            }
        }

        private void Authenticate(Session session, string capabilities)
        {
            if (_useDefaultCredentials || string.IsNullOrWhiteSpace(_username))
            {
                return;
            }
            if (!capabilities.Contains("AUTH", StringComparison.OrdinalIgnoreCase))
            {
                _logger.Warn("邮件服务器未声明 AUTH 能力，跳过认证");
                return;
            }
            if (capabilities.Contains("PLAIN", StringComparison.OrdinalIgnoreCase))
            {
                string token = Convert.ToBase64String(Encoding.UTF8.GetBytes("\0" + _username + "\0" + (_password ?? "")));
                session.Command("AUTH PLAIN " + token, 235, "认证");
                return;
            }
            session.Command("AUTH LOGIN", 334, "认证(1/2)");
            session.Command(Convert.ToBase64String(Encoding.UTF8.GetBytes(_username)), 334, "认证(2/2)");
            session.Command(Convert.ToBase64String(Encoding.UTF8.GetBytes(_password ?? "")), 235, "认证");
        }

        /// <summary>
        /// 组装 DATA 内容（头部 + 正文；主题/显示名按 RFC2047 编码，正文按需 Base64）
        /// </summary>
        private static string BuildData(MailMessage message)
        {
            StringBuilder sb = new();
            sb.Append("From: ").Append(FormatAddress(message.From, message.FromName)).Append("\r\n");
            sb.Append("To: ").Append(string.Join(", ", message.To)).Append("\r\n");
            if (message.Cc.Count > 0)
            {
                sb.Append("Cc: ").Append(string.Join(", ", message.Cc)).Append("\r\n");
            }
            sb.Append("Subject: ").Append(EncodeHeader(message.Subject ?? "")).Append("\r\n");
            sb.Append("Date: ").Append(DateTime.Now.ToString("r")).Append("\r\n");
            sb.Append("MIME-Version: 1.0\r\n");
            string body = message.Body ?? "";
            if (message.IsHtml || ContainsNonAscii(body))
            {
                string contentType = message.IsHtml ? "text/html" : "text/plain";
                sb.Append($"Content-Type: {contentType}; charset=utf-8\r\n");
                sb.Append("Content-Transfer-Encoding: base64\r\n\r\n");
                sb.Append(WrapBase64(Convert.ToBase64String(Encoding.UTF8.GetBytes(body))));
            }
            else
            {
                sb.Append("Content-Type: text/plain; charset=utf-8\r\n");
                sb.Append("Content-Transfer-Encoding: 8bit\r\n\r\n");
                sb.Append(body.Replace("\r\n", "\n").Replace("\n", "\r\n"));
            }
            sb.Append("\r\n.\r\n");
            return sb.ToString();
        }

        private static bool ContainsNonAscii(string text)
        {
            foreach (char c in text)
            {
                if (c > 127)
                {
                    return true;
                }
            }
            return false;
        }

        private static string WrapBase64(string base64)
        {
            StringBuilder sb = new();
            for (int i = 0; i < base64.Length; i += 76)
            {
                sb.Append(base64, i, Math.Min(76, base64.Length - i)).Append("\r\n");
            }
            return sb.ToString();
        }

        /// <summary>
        /// RFC2047 编码（含非 ASCII 时用 Base64 编码字，避免中文主题乱码）
        /// </summary>
        private static string EncodeHeader(string text)
            => ContainsNonAscii(text)
                ? "=?UTF-8?B?" + Convert.ToBase64String(Encoding.UTF8.GetBytes(text)) + "?="
                : text;

        private static string FormatAddress(string address, string displayName)
            => string.IsNullOrWhiteSpace(displayName) ? address : $"{EncodeHeader(displayName)} <{address}>";

        /// <summary>
        /// SMTP 会话：统一管理流/读写器，并支持中途升级为 TLS
        /// </summary>
        private sealed class Session : IDisposable
        {
            private static readonly UTF8Encoding Utf8NoBom = new(false);
            private readonly int _timeoutMs;
            private readonly bool _ignoreCertErrors;
            private readonly TcpClient _client;
            private Stream _stream;
            private StreamReader _reader;
            private StreamWriter _writer;

            public Session(Stream stream, TcpClient client, int timeoutMs, bool ignoreCertErrors)
            {
                _stream = stream;
                _client = client;
                _timeoutMs = timeoutMs;
                _ignoreCertErrors = ignoreCertErrors;
                ResetTextWrappers();
            }

            private void ResetTextWrappers()
            {
                _reader = new StreamReader(_stream, Utf8NoBom, false, 4096);
                _writer = new StreamWriter(_stream, Utf8NoBom, 4096, false) { AutoFlush = true, NewLine = "\r\n" };
            }

            /// <summary>
            /// 升级为 TLS（STARTTLS 或隐式 SSL）
            /// </summary>
            public void UpgradeToTls(string host)
            {
                _writer.Dispose();
                _reader.Dispose();
                SslStream ssl = new(_stream, true, _ignoreCertErrors ? AcceptAllCertificates : null);
                ssl.ReadTimeout = _timeoutMs;
                ssl.WriteTimeout = _timeoutMs;
                ssl.AuthenticateAsClient(host, null, SslProtocols.Tls12 | SslProtocols.Tls11 | SslProtocols.Tls, false);
                _stream = ssl;
                ResetTextWrappers();
            }

            public void WriteRaw(string text) => _writer.Write(text);

            public void WriteLine(string text) => _writer.WriteLine(text);

            /// <summary>
            /// 发送命令并校验响应码，返回完整响应文本
            /// </summary>
            public string Command(string command, int expected, string stage)
            {
                WriteLine(command);
                return Expect(expected, stage);
            }

            public string Expect(int expected, string stage) => ReadMultiline(expected, stage);

            public void TryQuit()
            {
                try
                {
                    Command("QUIT", 221, "QUIT");
                }
                catch
                {
                    // 退出失败无需处理
                }
            }

            private string ReadMultiline(int expected, string stage)
            {
                StringBuilder sb = new();
                while (true)
                {
                    string line = _reader.ReadLine();
                    if (line == null)
                    {
                        throw new IOException($"邮件服务器在[{stage}]阶段断开连接");
                    }
                    sb.AppendLine(line);
                    if (line.Length < 4 || line[3] != '-')
                    {
                        if (!int.TryParse(line.Substring(0, 3), out int code) || code != expected)
                        {
                            throw new InvalidOperationException($"邮件服务器在[{stage}]阶段返回 {line.Trim()}（期望 {expected}）");
                        }
                        return sb.ToString();
                    }
                }
            }

            private static bool AcceptAllCertificates(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                => true;

            public void Dispose()
            {
                try
                {
                    _writer?.Dispose();
                    _reader?.Dispose();
                    _stream?.Dispose();
                    _client?.Dispose();
                }
                catch
                {
                    // 释放失败忽略
                }
            }
        }
    }

    /// <summary>
    /// 邮件消息（与 System.Net.Mail.MailMessage 解耦的轻量载体）
    /// </summary>
    public class MailMessage
    {
        public string From { get; set; }

        public string FromName { get; set; }

        public List<string> To { get; } = [];

        public List<string> Cc { get; } = [];

        public string Subject { get; set; }

        public string Body { get; set; }

        public bool IsHtml { get; set; }
    }

    /// <summary>
    /// 邮件模板渲染：标记语法
    /// <list type="bullet">
    /// <item><description><c>{{Key}}</c> —— 变量替换（大小写不敏感，缺失变量替换为空并记警告）</description></item>
    /// <item><description><c>{{Key|yyyy/MM/dd}}</c> —— 带格式化（日期/数字格式化字符串）</description></item>
    /// <item><description>普通 <c>{</c> <c>}</c> 字符无需转义，仅 <c>{{...}}</c> 视为标记</description></item>
    /// </list>
    /// </summary>
    public static class MailTemplate
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>标记正则：{{ Key }} 或 {{ Key|format }}</summary>
        private static readonly Regex TokenRegex = new(@"\{\{\s*([A-Za-z0-9_]+)\s*(?:\|([^}]*))?\}\}", RegexOptions.Compiled);

        /// <summary>内置公共变量（调用方变量同名时以调用方为准）</summary>
        public static Dictionary<string, object> BuildBaseVariables()
        {
            DateTime now = DateTime.Now;
            return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["Date"] = now.Date,
                ["DateTime"] = now,
                ["Time"] = now.ToString("HH:mm"),
                ["Year"] = now.Year,
                ["AppName"] = LanguageService.Get("App_MainTitle"),
            };
        }

        /// <summary>
        /// 渲染模板
        /// </summary>
        public static string Render(string template, IDictionary<string, object> variables)
        {
            if (string.IsNullOrEmpty(template))
            {
                return "";
            }
            return TokenRegex.Replace(template, match =>
            {
                string key = match.Groups[1].Value;
                string format = match.Groups[2].Success ? match.Groups[2].Value : null;
                if (variables == null || !variables.TryGetValue(key, out object value) || value == null)
                {
                    _logger.Warn($"邮件模板变量缺失: {{{{{key}}}}}");
                    return "";
                }
                if (string.IsNullOrEmpty(format))
                {
                    return value is DateTime dt ? dt.ToString("yyyy/M/d") : value.ToString();
                }
                return value switch
                {
                    DateTime date => date.ToString(format),
                    IFormattable formattable => formattable.ToString(format, null),
                    _ => value.ToString()
                };
            });
        }
    }
}
