using System;
using System.Collections.Generic;

namespace MesBridge
{
    /// <summary>
    /// 桥接协议：主程序（64 位/AnyCPU）与 32 位桥接进程之间一行一条 JSON。
    /// 请求字段全部用 <see cref="Dictionary{TKey,TValue}"/> 读取，避免两边大小写/命名差异带来的坑。
    /// </summary>
    internal static class BridgeProtocol
    {
        /// <summary>SIM44 一行数据（字段语义与 MES「重工工令輸入」画面一致）</summary>
        internal sealed class Sim44Row
        {
            public string MoNo = "";
            public string LineNo = "";
            public string LineName = "";
            public string StartDate = "";
            public string PartNo = "";
            public string MoQty = "";
            public string RmaNo = "";
            public string RmaSeq = "";
            public string UpdUser = "";
            public string UpdDate = "";
            public string UpdTime = "";

            public Dictionary<string, object> ToDictionary()
            {
                return new Dictionary<string, object>
                {
                    ["moNo"] = MoNo,
                    ["lineNo"] = LineNo,
                    ["lineName"] = LineName,
                    ["startDate"] = StartDate,
                    ["partNo"] = PartNo,
                    ["moQty"] = MoQty,
                    ["rmaNo"] = RmaNo,
                    ["rmaSeq"] = RmaSeq,
                    ["updUser"] = UpdUser,
                    ["updDate"] = UpdDate,
                    ["updTime"] = UpdTime,
                };
            }
        }

        /// <summary>查询条件</summary>
        internal sealed class Sim44Filter
        {
            public string MoNo = "";
            public string PartNo = "";
            public string LineName = "";
            public string DateFrom = "";
            public string DateTo = "";
            public int Limit = 200;
        }

        /// <summary>从请求字典里取字符串（缺失/null 都返回空串）</summary>
        internal static string Str(IDictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object value) || value == null)
            {
                return "";
            }
            return Convert.ToString(value).Trim();
        }

        /// <summary>从请求字典里取整数，缺失或非法时用默认值</summary>
        internal static int Int(IDictionary<string, object> dict, string key, int fallback)
        {
            string text = Str(dict, key);
            return int.TryParse(text, out int value) ? value : fallback;
        }

        /// <summary>子对象（row / key）</summary>
        internal static IDictionary<string, object> Obj(IDictionary<string, object> dict, string key)
        {
            if (dict != null && dict.TryGetValue(key, out object value) && value is IDictionary<string, object> sub)
            {
                return sub;
            }
            return null;
        }

        /// <summary>按子对象构造 SIM44 行</summary>
        internal static Sim44Row RowFrom(IDictionary<string, object> dict)
        {
            IDictionary<string, object> row = Obj(dict, "row") ?? dict;
            return new Sim44Row
            {
                MoNo = Str(row, "moNo"),
                LineNo = Str(row, "lineNo"),
                LineName = Str(row, "lineName"),
                StartDate = Str(row, "startDate"),
                PartNo = Str(row, "partNo"),
                MoQty = Str(row, "moQty"),
                RmaNo = Str(row, "rmaNo"),
                RmaSeq = Str(row, "rmaSeq"),
                UpdUser = Str(row, "updUser"),
            };
        }
    }
}
