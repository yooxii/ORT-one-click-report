using System;
using System.Linq;
using System.Reflection;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 行快照与比对（领退表 / 计划表"未提交修改"检测用）：
    /// 快照 = MemberwiseClone 浅拷贝（两张表的字段只有字符串 / 数字 / 日期，浅拷贝即等价深拷贝）；
    /// 比对 = 逐属性值比较，只比参与存储的属性（跳过 [JsonIgnore] 的显示用属性，如 HasReportLink）。
    /// 替代原先每次都比 JSON 序列化字符串的做法——那是编辑时卡顿的主因。
    /// </summary>
    public static class RowSnapshot
    {
        /// <summary>
        /// 参与快照比对的属性：可读写、且没有 [JsonIgnore]（与数据库存储/原 JSON 对比口径一致）。
        /// </summary>
        public static PropertyInfo[] TrackedProperties(Type type) => type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead && p.CanWrite)
            .Where(p => p.GetCustomAttributes(typeof(Newtonsoft.Json.JsonIgnoreAttribute), true).Length == 0)
            .ToArray();

        /// <summary>克隆一行（浅拷贝，字段全是值类型/字符串时与深拷贝等价）</summary>
        public static T Clone<T>(T source) where T : class
        {
            if (source == null)
            {
                return null;
            }
            MethodInfo memberwiseClone = typeof(object)
                .GetMethod("MemberwiseClone", BindingFlags.Instance | BindingFlags.NonPublic);
            return (T)memberwiseClone.Invoke(source, null);
        }

        /// <summary>两行的参与存储属性是否完全一致</summary>
        public static bool AreSame(object a, object b, PropertyInfo[] properties)
        {
            if (ReferenceEquals(a, b))
            {
                return true;
            }
            if (a == null || b == null)
            {
                return false;
            }
            foreach (PropertyInfo property in properties)
            {
                if (!Equals(property.GetValue(a), property.GetValue(b)))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
