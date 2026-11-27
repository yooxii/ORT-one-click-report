using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;

namespace MesBridge
{
    /// <summary>
    /// 成品库存查询：还原自 C_IVI33（TIVI331FM 的主查询）。
    /// 原程式连的是 DACCT 库、所有表都通过 @bsfc_db 链接到 BSFC；
    /// 而我们的连接本来就在 BSFC 上，所以把 @bsfc_db 后缀去掉、原样执行即可，
    /// 不需要再开第二个连接。逻辑与原程式一致（含 180 天逾期判定、报废/领用核销排除）。只读。
    /// </summary>
    internal sealed class InventoryRepository
    {
        private readonly Sim44Repository _owner;

        internal InventoryRepository(Sim44Repository owner)
        {
            _owner = owner;
        }

        /// <summary>
        /// 按料号查成品在库栈板。日期是包装入库日期区间（DELIVER_DATE），料号精确匹配（与原程式一致），
        /// 栈板号支持模糊。
        /// </summary>
        internal List<Dictionary<string, object>> Stock(string partNo, string recvDateFrom, string recvDateTo,
            string palletNo, int limit)
        {
            const string inner = @"SELECT A.*,
                       TRUNC(SYSDATE - TO_DATE(MO_IN_DATE, 'yyyy-mm-dd')) MO_IN_DAY,
                       DECODE(SIGN(TRUNC(SYSDATE - TO_DATE(MO_IN_DATE, 'yyyy-mm-dd')) - 180),1,'逾期',0,'非逾期',-1,'非逾期') OVER_STATUS,
                       NVL(D.UNFORBID_YN, 'N') UNFORBID_YN
                  FROM (
                        SELECT A.PART_NO, A.PALLET_NO, A.WARE_CODE, COUNT(A.PALLET_NO) QTY, A.PART_VER, F.DATE_CODE,
                               NVL((SELECT DISTINCT MO_IN_DATE FROM IVM10 WHERE API_IN_NO = A.IN_NO AND ROWNUM = 1), B.INSPECT_DATE) MO_IN_DATE
                          FROM (SELECT PART_NO, PALLET_NO, WARE_CODE, PART_VER, LINE_NAME, REMARK, IN_NO FROM SIM92
                                 WHERE PART_NO = :PART_NO
                                   AND NVL(LAST_PALLET,' ') <> '結案掃描'
                                   AND DELIVER_DATE BETWEEN :RECV_DATE1 AND :RECV_DATE2) A,
                               (SELECT PALLET_NO, INSPECT_DATE FROM SIM96
                                 WHERE RESULT IN ('ACC','UAI')
                                   AND PART_NO = :PART_NO2) B,
                               (SELECT PALLET_NO, DATE_CODE FROM SIM94 C
                                 WHERE PART_NO = :PART_NO3
                                   AND PALLET_NO LIKE :PALLET_NO
                                   AND NOT EXISTS (SELECT PELLET_NO FROM SIM33 D WHERE PELLET_NO = C.PALLET_NO AND BAR_CODE_START = C.BAR_CODE)
                                   AND NOT EXISTS (SELECT PELLET_NO FROM IVM635_DETAIL D WHERE PELLET_NO = C.PALLET_NO AND BAR_CODE = C.BAR_CODE)
                                   AND NOT EXISTS (SELECT PALLET_NO FROM IVM634_DETAIL D WHERE PALLET_NO = C.PALLET_NO AND BAR_CODE = C.BAR_CODE)
                                   AND GET_SIM33_BAR_YN(3, BAR_CODE, C.PART_NO) = 'N'
                                   AND GET_SIM33_BAR_YN(4, BAR_CODE, NULL) = 'N'
                                   AND GET_SIM33_BAR_YN(5, BAR_CODE, NULL) = 'N'
                                   AND GET_SIM33_BAR_YN(6, BAR_CODE, C.PART_NO) = 'N'
                                   AND GET_SIM33_BAR_YN(7, BAR_CODE, NULL) = 'N'
                                   AND GET_SIM33_BAR_YN(8, BAR_CODE, C.PART_NO) = 'N') F
                         WHERE A.PALLET_NO = B.PALLET_NO
                           AND A.PALLET_NO = F.PALLET_NO
                        GROUP BY PART_NO, A.PALLET_NO, WARE_CODE, A.PART_VER, F.DATE_CODE, IN_NO, INSPECT_DATE
                  ) A, SIM155_DETAIL D
                 WHERE A.PART_NO = D.PART_NO(+)
                   AND A.PALLET_NO = D.PALLET_NO(+)
                 ORDER BY MO_IN_DAY DESC";

            string sql = "SELECT * FROM (" + inner + ") WHERE ROWNUM <= :lim";
            using (OracleCommand cmd = _owner.Connection.CreateCommand())
            {
                cmd.CommandText = sql;
                _owner.SetLastSql(sql);
                string part = (partNo ?? "").Trim();
                cmd.Parameters.Add("PART_NO", OracleType.VarChar).Value = part;
                cmd.Parameters.Add("PART_NO2", OracleType.VarChar).Value = part;
                cmd.Parameters.Add("PART_NO3", OracleType.VarChar).Value = part;
                cmd.Parameters.Add("RECV_DATE1", OracleType.VarChar).Value =
                    string.IsNullOrEmpty(recvDateFrom) ? "00000000" : recvDateFrom;
                cmd.Parameters.Add("RECV_DATE2", OracleType.VarChar).Value =
                    string.IsNullOrEmpty(recvDateTo) ? "99999999" : recvDateTo;
                cmd.Parameters.Add("PALLET_NO", OracleType.VarChar).Value = "%" + (palletNo ?? "").Trim() + "%";
                cmd.Parameters.Add("lim", OracleType.Int32).Value = limit <= 0 ? 500 : Math.Min(limit, 3000);
                return Read(cmd);
            }
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
