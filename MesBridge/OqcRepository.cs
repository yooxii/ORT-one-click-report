using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;

namespace MesBridge
{
    /// <summary>
    /// OQC 测试记录查询（还原自 C_SII51.EXE 内嵌 SQL）。
    /// 主查询＝托盘级汇总（SIM92 主档 + SIM96 检验结果 + SIM94 明细，按 :PART_NO / 日期 / :PALLET_NO /
    /// :MO_NO / :BAR_CODE / :RESULT / :STATUS 过滤）；明细＝SIM94 的条码级数据。
    /// 只读，不改任何数据。
    /// </summary>
    internal sealed class OqcRepository
    {
        private readonly Sim44Repository _owner;

        internal OqcRepository(Sim44Repository owner)
        {
            _owner = owner;
        }

        /// <summary>
        /// 主查询：与原程式 TFRMMASTER 的第一条 SQL 一致。
        /// 日期是 yyyyMMdd 字符串，起止都为空时用 '00000000'~'99999999' 表示不限。
        /// </summary>
        internal List<Dictionary<string, object>> Query(string partNo, string sDate, string eDate,
            string palletNo, string moNo, string barCode, string result, string status, int limit)
        {
            const string sql = @"SELECT * FROM (
                SELECT distinct A.PALLET_NO,DELIVER_DATE,A.LINE_NAME,A.WARE_CODE,
                       A.PART_NO,PART_VER,a.STATUS,DELIVER_QTY,SAMPLE_QTY,RESULT,INSPECTOR,INSPECT_DATE,
                       A.TRANS_YN,REMARK
                  FROM APITECH.SIM92 A, APITECH.SIM96 B, apitech.sim94 c
                 WHERE A.PALLET_NO = B.PALLET_NO(+)
                   AND a.pallet_no = c.pallet_no(+)
                   AND A.PART_NO LIKE :PART_NO
                   AND 1 = 1
                   AND NVL(INSPECT_DATE,'00000000') BETWEEN :s_date AND :e_date
                   AND A.PALLET_NO LIKE :PALLET_NO
                   AND NVL(c.mo_no,' ') LIKE :mo_no
                   AND NVL(c.bar_code,' ') LIKE :bar_code
                   AND NVL(B.RESULT,' ') LIKE :result
                   AND A.STATUS LIKE :STATUS
                 ORDER BY A.PALLET_NO) WHERE ROWNUM <= :lim";
            using (OracleCommand cmd = _owner.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                _owner.SetLastSql(sql);
                cmd.Parameters.Add("PART_NO", OracleType.VarChar).Value = Like(partNo);
                cmd.Parameters.Add("s_date", OracleType.VarChar).Value = string.IsNullOrEmpty(sDate) ? "00000000" : sDate;
                cmd.Parameters.Add("e_date", OracleType.VarChar).Value = string.IsNullOrEmpty(eDate) ? "99999999" : eDate;
                cmd.Parameters.Add("PALLET_NO", OracleType.VarChar).Value = Like(palletNo);
                cmd.Parameters.Add("mo_no", OracleType.VarChar).Value = Like(moNo);
                cmd.Parameters.Add("bar_code", OracleType.VarChar).Value = Like(barCode);
                cmd.Parameters.Add("result", OracleType.VarChar).Value = Like(result);
                cmd.Parameters.Add("STATUS", OracleType.VarChar).Value = Like(status);
                cmd.Parameters.Add("lim", OracleType.Int32).Value = Clamp(limit);
                return Read(cmd);
            }
        }

        /// <summary>某个托盘下的条码明细（SIM94）</summary>
        internal List<Dictionary<string, object>> Detail(string partNo, string palletNo, string moNo, string barCode,
            int limit)
        {
            const string sql = @"SELECT * FROM (
                SELECT PART_NO,PALLET_NO,MO_NO,SEQ_NO,BAR_CODE,IS_INDEX,BAR_CODE1,BAR_CODE2,BAR_CODE3,BAR_CODE4,
                       BAR_CODE5,STATUS,RC_NO,RW_YN,BOX_NO,RW_REMARK
                  FROM APITECH.SIM94
                 WHERE PART_NO LIKE :PART_NO
                   AND PALLET_NO LIKE :PALLET_NO
                   AND MO_NO LIKE :MO_NO
                   AND BAR_CODE LIKE :BAR_CODE
                 ORDER BY pallet_no,seq_no) WHERE ROWNUM <= :lim";
            using (OracleCommand cmd = _owner.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                _owner.SetLastSql(sql);
                cmd.Parameters.Add("PART_NO", OracleType.VarChar).Value = Like(partNo);
                cmd.Parameters.Add("PALLET_NO", OracleType.VarChar).Value = Like(palletNo);
                cmd.Parameters.Add("MO_NO", OracleType.VarChar).Value = Like(moNo);
                cmd.Parameters.Add("BAR_CODE", OracleType.VarChar).Value = Like(barCode);
                cmd.Parameters.Add("lim", OracleType.Int32).Value = Clamp(limit);
                return Read(cmd);
            }
        }

        private static string Like(string text)
        {
            return "%" + (text ?? "").Trim() + "%";
        }

        private static int Clamp(int limit)
        {
            int value = limit <= 0 ? 300 : limit;
            return Math.Min(value, 3000);
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
