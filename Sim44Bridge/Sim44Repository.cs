using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Globalization;
using System.Text;

namespace Sim44Bridge
{
    /// <summary>
    /// SIM44（MES「重工工令」表）访问层。只在 32 位进程里跑：
    /// Oracle 8.1.7 客户端（C:\oracle\ora81）是 32 位的，64 位进程根本加载不了。
    /// </summary>
    internal sealed class Sim44Repository : IDisposable
    {
        /// <summary>默认连接别名（与 BDE 别名 BSFC 指向的 SERVER NAME 一致）</summary>
        internal const string DefaultDataSource = "NT_BSFC.WORLD";

        /// <summary>表名（可用配置覆盖成 APITECH.SIM44 之类带 schema 的写法）</summary>
        private readonly string _table;

        private OracleConnection _conn;
        private string _user = "";

        internal Sim44Repository(string dataSource, string table)
        {
            DataSource = string.IsNullOrWhiteSpace(dataSource) ? DefaultDataSource : dataSource.Trim();
            _table = string.IsNullOrWhiteSpace(table) ? "SIM44" : table.Trim();
        }

        internal string DataSource { get; }

        internal string CurrentUser => _user;

        internal bool IsOpen => _conn != null && _conn.State == ConnectionState.Open;

        /// <summary>
        /// 探测本机能不能加载 Oracle 客户端：用一个不存在的 TNS 名去连，
        /// 拿到 ORA-12154/12541 说明客户端正常，拿到别的（找不到 DLL / 位数不对）说明环境有问题。
        /// </summary>
        internal static (bool ok, string code, string message) SelfTest()
        {
            if (Environment.Is64BitProcess)
            {
                return (false, "bitness", "桥接进程必须是 32 位（Oracle 8i 客户端只有 32 位）");
            }
            try
            {
                string cs = "Data Source=SIM44_BRIDGE_SELFTEST_UNKNOWN;User Id=x;Password=y;";
                using (OracleConnection conn = new OracleConnection(cs))
                {
                    conn.Open();
                }
                return (true, "ok", "客户端可用（意外连上了自检源，忽略）");
            }
            catch (OracleException ex)
            {
                // 12154 = 解析不到服务名，12541 = 无监听：都证明 OCI 本身已经加载成功
                return (true, "ok", "Oracle 客户端可用（" + ex.Message.Split('\n')[0].Trim() + "）");
            }
            catch (Exception ex)
            {
                return (false, "nocli", "Oracle 客户端不可用：" + ex.Message);
            }
        }

        /// <summary>建立连接（登录）。抛出的异常由调用方分类。</summary>
        internal string Open(string user, string password)
        {
            Close();
            OracleConnectionStringBuilder builder = new OracleConnectionStringBuilder
            {
                DataSource = DataSource,
                UserID = user ?? "",
                Password = password ?? "",
                // 老库不建议连接池：8i 客户端上连接池偶发挂死，且本进程本身是短连接
                Pooling = false,
            };
            _conn = new OracleConnection(builder.ConnectionString);
            _conn.Open();
            using (OracleCommand cmd = _conn.CreateCommand())
            {
                cmd.CommandText = "SELECT USER FROM DUAL";
                object value = cmd.ExecuteScalar();
                _user = value == null ? (user ?? "") : Convert.ToString(value);
            }
            return _user;
        }

        /// <summary>按条件查询 SIM44</summary>
        internal List<BridgeProtocol.Sim44Row> Query(BridgeProtocol.Sim44Filter filter)
        {
            EnsureOpen();
            StringBuilder sql = new StringBuilder();
            sql.Append("SELECT MO_NO, LINE_NO, LINE_NAME, START_DATE, PART_NO, MO_QTY, RMA_NO, RMA_SEQ, UPD_USER, UPD_DATE, UPD_TIME FROM ( SELECT MO_NO, LINE_NO, LINE_NAME, START_DATE, PART_NO, MO_QTY, RMA_NO, RMA_SEQ, UPD_USER, UPD_DATE, UPD_TIME FROM ");
            sql.Append(_table);
            sql.Append(" WHERE 1 = 1");

            using (OracleCommand cmd = _conn.CreateCommand())
            {
                if (!string.IsNullOrEmpty(filter.MoNo))
                {
                    sql.Append(" AND MO_NO LIKE :mo ");
                    cmd.Parameters.Add("mo", OracleType.VarChar).Value = filter.MoNo + "%";
                }
                if (!string.IsNullOrEmpty(filter.PartNo))
                {
                    sql.Append(" AND (PART_NO LIKE :pn OR PART_NO = :pn2) ");
                    cmd.Parameters.Add("pn", OracleType.VarChar).Value = "%" + filter.PartNo + "%";
                    cmd.Parameters.Add("pn2", OracleType.VarChar).Value = filter.PartNo;
                }
                if (!string.IsNullOrEmpty(filter.LineName))
                {
                    sql.Append(" AND (LINE_NAME LIKE :ln OR LINE_NO = :lno) ");
                    cmd.Parameters.Add("ln", OracleType.VarChar).Value = "%" + filter.LineName + "%";
                    cmd.Parameters.Add("lno", OracleType.VarChar).Value = filter.LineName;
                }
                if (!string.IsNullOrEmpty(filter.DateFrom))
                {
                    sql.Append(" AND START_DATE >= :d1 ");
                    cmd.Parameters.Add("d1", OracleType.VarChar).Value = filter.DateFrom;
                }
                if (!string.IsNullOrEmpty(filter.DateTo))
                {
                    sql.Append(" AND START_DATE <= :d2 ");
                    cmd.Parameters.Add("d2", OracleType.VarChar).Value = filter.DateTo;
                }
                sql.Append(" ORDER BY MO_NO, LINE_NO ) WHERE ROWNUM <= :lim");
                int limit = filter.Limit <= 0 ? 200 : Math.Min(filter.Limit, 2000);
                cmd.Parameters.Add("lim", OracleType.Int32).Value = limit;
                cmd.CommandText = sql.ToString();

                List<BridgeProtocol.Sim44Row> rows = new List<BridgeProtocol.Sim44Row>();
                using (OracleDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        rows.Add(new BridgeProtocol.Sim44Row
                        {
                            MoNo = Col(reader, 0),
                            LineNo = Col(reader, 1),
                            LineName = Col(reader, 2),
                            StartDate = Col(reader, 3),
                            PartNo = Col(reader, 4),
                            MoQty = Col(reader, 5),
                            RmaNo = Col(reader, 6),
                            RmaSeq = Col(reader, 7),
                            UpdUser = Col(reader, 8),
                            UpdDate = Col(reader, 9),
                            UpdTime = Col(reader, 10),
                        });
                    }
                }
                return rows;
            }
        }

        /// <summary>按工令查一笔（用于「修改」前的定位与重复判断）</summary>
        internal List<BridgeProtocol.Sim44Row> GetByMoNo(string moNo)
        {
            return Query(new BridgeProtocol.Sim44Filter { MoNo = moNo, Limit = 50 });
        }

        /// <summary>新增一笔；工令已存在时直接报错（与原 MES 画面口径一致）</summary>
        internal int Insert(BridgeProtocol.Sim44Row row)
        {
            EnsureOpen();
            ValidateRow(row);
            if (GetByMoNo(row.MoNo).Count > 0)
            {
                throw new BridgeException("dup", "工令 " + row.MoNo + " 已存在，不能重复新增");
            }
            const string sql = @"INSERT INTO {0}
                (MO_NO, LINE_NO, LINE_NAME, START_DATE, PART_NO, MO_QTY, RMA_NO, RMA_SEQ, UPD_DATE, UPD_USER, UPD_TIME)
                VALUES (:mo, :lno, :lname, :sdate, :pno, :qty, :rma, :rseq,
                        TO_CHAR(SYSDATE,'YYYYMMDD'), :uuser, TO_CHAR(SYSDATE,'HH:MI:SS'))";
            using (OracleCommand cmd = _conn.CreateCommand())
            {
                cmd.CommandText = string.Format(sql, _table);
                cmd.Parameters.Add("mo", OracleType.VarChar).Value = row.MoNo;
                cmd.Parameters.Add("lno", OracleType.VarChar).Value = Nullable(row.LineNo);
                cmd.Parameters.Add("lname", OracleType.VarChar).Value = Nullable(row.LineName);
                cmd.Parameters.Add("sdate", OracleType.VarChar).Value = Nullable(row.StartDate);
                cmd.Parameters.Add("pno", OracleType.VarChar).Value = Nullable(row.PartNo);
                cmd.Parameters.Add("qty", OracleType.Double).Value = ToNumber(row.MoQty, "数量");
                cmd.Parameters.Add("rma", OracleType.VarChar).Value = Nullable(row.RmaNo);
                cmd.Parameters.Add("rseq", OracleType.Double).Value = ToNumberOrNull(row.RmaSeq, "RMA_SEQ");
                cmd.Parameters.Add("uuser", OracleType.VarChar).Value = Nullable(row.UpdUser, 20);
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>按工令修改一笔</summary>
        internal int Update(BridgeProtocol.Sim44Row row)
        {
            EnsureOpen();
            ValidateRow(row);
            const string sql = @"UPDATE {0} SET
                    LINE_NO = :lno, LINE_NAME = :lname, START_DATE = :sdate, PART_NO = :pno,
                    MO_QTY = :qty, RMA_NO = :rma, RMA_SEQ = :rseq,
                    UPD_DATE = TO_CHAR(SYSDATE,'YYYYMMDD'), UPD_USER = :uuser,
                    UPD_TIME = TO_CHAR(SYSDATE,'HH:MI:SS')
                WHERE MO_NO = :mo";
            using (OracleCommand cmd = _conn.CreateCommand())
            {
                cmd.CommandText = string.Format(sql, _table);
                cmd.Parameters.Add("lno", OracleType.VarChar).Value = Nullable(row.LineNo);
                cmd.Parameters.Add("lname", OracleType.VarChar).Value = Nullable(row.LineName);
                cmd.Parameters.Add("sdate", OracleType.VarChar).Value = Nullable(row.StartDate);
                cmd.Parameters.Add("pno", OracleType.VarChar).Value = Nullable(row.PartNo);
                cmd.Parameters.Add("qty", OracleType.Double).Value = ToNumber(row.MoQty, "数量");
                cmd.Parameters.Add("rma", OracleType.VarChar).Value = Nullable(row.RmaNo);
                cmd.Parameters.Add("rseq", OracleType.Double).Value = ToNumberOrNull(row.RmaSeq, "RMA_SEQ");
                cmd.Parameters.Add("uuser", OracleType.VarChar).Value = Nullable(row.UpdUser, 20);
                cmd.Parameters.Add("mo", OracleType.VarChar).Value = row.MoNo;
                int affected = cmd.ExecuteNonQuery();
                if (affected <= 0)
                {
                    throw new BridgeException("notfound", "找不到工令 " + row.MoNo + "，无法修改");
                }
                return affected;
            }
        }

        /// <summary>机种名/料号查 BMM04（MES 的料号主档，PART_TYPE='MP' 才有下拉）</summary>
        internal List<Dictionary<string, object>> LookupPart(string text, int limit)
        {
            EnsureOpen();
            List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();
            using (OracleCommand cmd = _conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT PART_NO, PART_DESCRP, PART_SPEC FROM (
                        SELECT PART_NO, PART_DESCRP, PART_SPEC FROM BMM04
                         WHERE PART_TYPE = 'MP'
                           AND (PART_NO LIKE :t OR PART_DESCRP LIKE :t2)
                         ORDER BY PART_NO) WHERE ROWNUM <= :lim";
                cmd.Parameters.Add("t", OracleType.VarChar).Value = "%" + (text ?? "") + "%";
                cmd.Parameters.Add("t2", OracleType.VarChar).Value = "%" + (text ?? "") + "%";
                cmd.Parameters.Add("lim", OracleType.Int32).Value = limit <= 0 ? 50 : Math.Min(limit, 500);
                using (OracleDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new Dictionary<string, object>
                        {
                            ["partNo"] = Col(reader, 0),
                            ["model"] = Col(reader, 1),
                            ["spec"] = Col(reader, 2),
                        });
                    }
                }
            }
            return result;
        }

        /// <summary>线别查 SIM11（线名 ↔ 线代号）</summary>
        internal List<Dictionary<string, object>> LookupLine(string text, int limit)
        {
            EnsureOpen();
            List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();
            using (OracleCommand cmd = _conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT LINE_NO, LINE_NAME FROM (
                        SELECT LINE_NO, LINE_NAME FROM SIM11
                         WHERE LINE_NAME LIKE :t OR LINE_NO = :t2
                         ORDER BY LINE_NO) WHERE ROWNUM <= :lim";
                cmd.Parameters.Add("t", OracleType.VarChar).Value = "%" + (text ?? "") + "%";
                cmd.Parameters.Add("t2", OracleType.VarChar).Value = text ?? "";
                cmd.Parameters.Add("lim", OracleType.Int32).Value = limit <= 0 ? 50 : Math.Min(limit, 500);
                using (OracleDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new Dictionary<string, object>
                        {
                            ["lineNo"] = Col(reader, 0),
                            ["lineName"] = Col(reader, 1),
                        });
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 改自己的数据库密码：先用旧密码验证一次，再 ALTER USER，最后用新密码回连确认。
        /// </summary>
        internal void ChangePassword(string oldPassword, string newPassword)
        {
            if (string.IsNullOrEmpty(_user))
            {
                throw new BridgeException("nologin", "尚未登录");
            }
            if (string.IsNullOrEmpty(newPassword))
            {
                throw new BridgeException("empty", "新密码不能为空");
            }
            // 1. 旧密码必须对（用临时连接验证，不动当前连接）
            using (OracleConnection check = new OracleConnection(BuildConnectionString(oldPassword)))
            {
                check.Open();
            }
            // 2. 改密码（Oracle 里双引号包起来才不会把特殊字符当语法）
            string quotedUser = "\"" + _user.Replace("\"", "\"\"") + "\"";
            string quotedPass = "\"" + newPassword.Replace("\"", "\"\"") + "\"";
            using (OracleConnection admin = new OracleConnection(BuildConnectionString(oldPassword)))
            {
                admin.Open();
                using (OracleCommand cmd = admin.CreateCommand())
                {
                    cmd.CommandText = "ALTER USER " + quotedUser + " IDENTIFIED BY " + quotedPass;
                    cmd.ExecuteNonQuery();
                }
            }
            // 3. 用新密码回连，确认改成功
            Open(_user, newPassword);
        }

        private string BuildConnectionString(string password)
        {
            OracleConnectionStringBuilder builder = new OracleConnectionStringBuilder
            {
                DataSource = DataSource,
                UserID = _user,
                Password = password ?? "",
                Pooling = false,
            };
            return builder.ConnectionString;
        }

        private void EnsureOpen()
        {
            if (!IsOpen)
            {
                throw new BridgeException("nologin", "尚未登录数据库");
            }
        }

        private static void ValidateRow(BridgeProtocol.Sim44Row row)
        {
            if (string.IsNullOrWhiteSpace(row.MoNo))
            {
                throw new BridgeException("invalid", "工令（MO_NO）不能为空");
            }
            if (row.MoNo.Length > 30)
            {
                throw new BridgeException("invalid", "工令（MO_NO）长度不能超过 30");
            }
            if (string.IsNullOrWhiteSpace(row.MoQty))
            {
                throw new BridgeException("invalid", "数量（MO_QTY）不能为空");
            }
        }

        private static string Nullable(string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static string Nullable(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }

        private static double ToNumber(string text, string fieldName)
        {
            if (!double.TryParse((text ?? "").Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                throw new BridgeException("invalid", fieldName + " 必须是数字：" + text);
            }
            return value;
        }

        private static object ToNumberOrNull(string text, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return DBNull.Value;
            }
            return ToNumber(text, fieldName);
        }

        private static string Col(OracleDataReader reader, int index)
        {
            if (reader.IsDBNull(index))
            {
                return "";
            }
            object value = reader.GetValue(index);
            if (value is double || value is decimal || value is float)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.####", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value).Trim();
        }

        internal void Close()
        {
            try
            {
                _conn?.Close();
            }
            catch
            {
                // 关闭失败不影响后续
            }
            finally
            {
                _conn?.Dispose();
                _conn = null;
            }
        }

        public void Dispose()
        {
            Close();
        }
    }

    /// <summary>带错误码的业务异常（桥接进程把它翻译成 {ok:false, code:...}）</summary>
    internal sealed class BridgeException : Exception
    {
        internal BridgeException(string code, string message) : base(message)
        {
            Code = code;
        }

        internal string Code { get; }
    }
}
