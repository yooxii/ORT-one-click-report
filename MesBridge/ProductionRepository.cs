using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;

namespace MesBridge
{
    /// <summary>
    /// 生产排程（IVM540）查询：还原自 C_IVI540_BL「小工令查询 / 月排程」。
    /// 这张表就是 ORT 预领「月排程」的数据来源，按生产日期区间（可选 PCB）取。只读。
    /// </summary>
    internal sealed class ProductionRepository
    {
        private readonly Sim44Repository _owner;
        private string _prodDateType;

        internal ProductionRepository(Sim44Repository owner)
        {
            _owner = owner;
        }

        /// <summary>
        /// 月排程查询。日期按 yyyyMMdd 传入；PROD_DATE 的实际类型（DATE / 字符串）
        /// 先查一次数据字典再决定怎么写条件，避免隐式转换把查询弄坏。
        /// </summary>
        internal List<Dictionary<string, object>> Schedule(string dateFrom, string dateTo, string pcb, int limit)
        {
            string table = _owner.ResolveTable("IVM540");
            bool isDate = ProdDateIsDate(table);
            List<string> where = new List<string>();
            List<KeyValuePair<string, object>> args = new List<KeyValuePair<string, object>>();

            if (!string.IsNullOrEmpty(dateFrom))
            {
                where.Add(isDate ? "PROD_DATE >= TO_DATE(:d1,'YYYYMMDD')" : "PROD_DATE >= :d1");
                args.Add(new KeyValuePair<string, object>("d1", dateFrom));
            }
            if (!string.IsNullOrEmpty(dateTo))
            {
                where.Add(isDate ? "PROD_DATE <= TO_DATE(:d2,'YYYYMMDD')" : "PROD_DATE <= :d2");
                args.Add(new KeyValuePair<string, object>("d2", dateTo));
            }
            if (!string.IsNullOrWhiteSpace(pcb))
            {
                where.Add("PCB LIKE :pcb");
                args.Add(new KeyValuePair<string, object>("pcb", "%" + pcb.Trim() + "%"));
            }
            string condition = where.Count == 0 ? "" : " AND " + string.Join(" AND ", where.ToArray());

            // 列名不写死：不同版本 IVM540 的列会变，SELECT * 交给界面按实际列渲染
            string sql = "SELECT * FROM ( SELECT * FROM " + table + " WHERE 1 = 1" + condition
                         + " ORDER BY PROD_DATE, PCB ) WHERE ROWNUM <= :lim";
            using (OracleCommand cmd = _owner.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                _owner.SetLastSql(sql);
                foreach (KeyValuePair<string, object> arg in args)
                {
                    cmd.Parameters.Add(arg.Key, OracleType.VarChar).Value = arg.Value;
                }
                cmd.Parameters.Add("lim", OracleType.Int32).Value = limit <= 0 ? 500 : Math.Min(limit, 3000);
                return Read(cmd);
            }
        }

        /// <summary>准备领用单（IVM530_MP）：还原自 C_IVR621_BL，按单据日期区间/单据号区间过滤。只读。</summary>
        internal List<Dictionary<string, object>> PrepareIssue(string dateFrom, string dateTo, string user,
            string vchnoFrom, string vchnoTo, string partFrom, string partTo, int limit)
        {
            string table = _owner.ResolveTable("IVM530_MP");
            List<string> where = new List<string>();
            List<KeyValuePair<string, object>> args = new List<KeyValuePair<string, object>>();
            if (!string.IsNullOrEmpty(dateFrom))
            {
                where.Add("VCHNO_DATE >= :d1");
                args.Add(new KeyValuePair<string, object>("d1", dateFrom));
            }
            if (!string.IsNullOrEmpty(dateTo))
            {
                where.Add("VCHNO_DATE <= :d2");
                args.Add(new KeyValuePair<string, object>("d2", dateTo));
            }
            if (!string.IsNullOrWhiteSpace(user))
            {
                where.Add("UPD_USER LIKE :u");
                args.Add(new KeyValuePair<string, object>("u", "%" + user.Trim() + "%"));
            }
            if (!string.IsNullOrWhiteSpace(vchnoFrom))
            {
                where.Add("VCHNO >= :v1");
                args.Add(new KeyValuePair<string, object>("v1", vchnoFrom.Trim()));
            }
            if (!string.IsNullOrWhiteSpace(vchnoTo))
            {
                where.Add("VCHNO <= :v2");
                args.Add(new KeyValuePair<string, object>("v2", vchnoTo.Trim()));
            }
            if (!string.IsNullOrWhiteSpace(partFrom))
            {
                where.Add("PART_NO >= :p1");
                args.Add(new KeyValuePair<string, object>("p1", partFrom.Trim()));
            }
            if (!string.IsNullOrWhiteSpace(partTo))
            {
                where.Add("PART_NO <= :p2");
                args.Add(new KeyValuePair<string, object>("p2", partTo.Trim()));
            }
            string condition = where.Count == 0 ? "" : " WHERE " + string.Join(" AND ", where.ToArray());
            string sql = "SELECT * FROM ( SELECT * FROM " + table + condition
                         + " ORDER BY VCHNO_DATE DESC, VCHNO, VCH_SEQ ) WHERE ROWNUM <= :lim";
            using (OracleCommand cmd = _owner.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                _owner.SetLastSql(sql);
                foreach (KeyValuePair<string, object> arg in args)
                {
                    cmd.Parameters.Add(arg.Key, OracleType.VarChar).Value = arg.Value;
                }
                cmd.Parameters.Add("lim", OracleType.Int32).Value = limit <= 0 ? 500 : Math.Min(limit, 3000);
                return Read(cmd);
            }
        }

        /// <summary>PROD_DATE 是不是 DATE 类型（查一次就缓存）</summary>
        private bool ProdDateIsDate(string table)
        {
            if (_prodDateType != null)
            {
                return _prodDateType == "DATE";
            }
            string owner = "";
            string name = table;
            int dot = table.IndexOf('.');
            if (dot > 0)
            {
                owner = table.Substring(0, dot);
                name = table.Substring(dot + 1);
            }
            _prodDateType = "VARCHAR2";
            try
            {
                using (OracleCommand cmd = _owner.Connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT DATA_TYPE FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = :t AND COLUMN_NAME = 'PROD_DATE'"
                                      + (owner.Length > 0 ? " AND OWNER = :o" : " AND OWNER = USER");
                    cmd.Parameters.Add("t", OracleType.VarChar).Value = name.ToUpperInvariant();
                    if (owner.Length > 0)
                    {
                        cmd.Parameters.Add("o", OracleType.VarChar).Value = owner.ToUpperInvariant();
                    }
                    object value = cmd.ExecuteScalar();
                    if (value != null)
                    {
                        _prodDateType = Convert.ToString(value).Trim().ToUpperInvariant();
                    }
                }
            }
            catch
            {
                // 查不到就按字符串处理（MES 里日期大多是 yyyyMMdd 字符串）
            }
            return _prodDateType == "DATE";
        }

        private static List<Dictionary<string, object>> Read(OracleCommand cmd)
        {
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            using (OracleDataReader reader = cmd.ExecuteReader())
            {
                int count = reader.FieldCount;
                while (reader.Read())
                {
                    Dictionary<string, object> row = new Dictionary<string, object>();
                    for (int i = 0; i < count; i++)
                    {
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? "" : Convert.ToString(reader.GetValue(i))?.Trim();
                    }
                    rows.Add(row);
                }
            }
            return rows;
        }
    }
}
