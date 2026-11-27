using NUnit.Framework;
using ORT一键报告.Services.Sim44;
using System;
using System.Linq;
using System.Reflection;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// SIM44 回线工令相关规则的测试：
    /// 1) 数据库密码「每 90 天变更」的到期计算；
    /// 2) 密码 DPAPI 加解密往返（不落盘，只测算法本身）；
    /// 3) 工令行与桥接程序之间的 JSON 往返。
    /// </summary>
    [TestFixture]
    public class Sim44CredentialTests
    {
        [Test]
        public void 到期计算_今天改密还剩满额天数()
        {
            Assert.That(Sim44CredentialStore.DaysUntilExpiry(DateTime.Today),
                Is.EqualTo(Sim44CredentialStore.PasswordValidDays));
        }

        [Test]
        public void 到期计算_按天数递减_超期给负数()
        {
            Assert.That(Sim44CredentialStore.DaysUntilExpiry(DateTime.Today.AddDays(-89)), Is.EqualTo(1));
            Assert.That(Sim44CredentialStore.DaysUntilExpiry(DateTime.Today.AddDays(-90)), Is.EqualTo(0));
            Assert.That(Sim44CredentialStore.DaysUntilExpiry(DateTime.Today.AddDays(-91)), Is.EqualTo(-1));
        }

        [Test]
        public void 到期计算_不知道上次改密时间时不给结论()
        {
            Assert.That(Sim44CredentialStore.DaysUntilExpiry(null), Is.Null);
        }

        [Test]
        public void 到期计算_忽略时分秒只按日期算()
        {
            DateTime changed = DateTime.Today.AddDays(-10).AddHours(23).AddMinutes(59);
            Assert.That(Sim44CredentialStore.DaysUntilExpiry(changed),
                Is.EqualTo(Sim44CredentialStore.PasswordValidDays - 10));
        }

        [Test]
        public void 密码加密_能解回原文_且密文里看不到明文()
        {
            MethodInfo encrypt = typeof(Sim44CredentialStore).GetMethod("Encrypt",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.That(encrypt, Is.Not.Null, "找不到 DPAPI 加密方法");

            const string plain = "P@ssw0rd-测试-123";
            string cipher = (string)encrypt.Invoke(null, [plain]);
            Assert.That(cipher, Is.Not.Null.And.Not.Empty);
            Assert.That(cipher, Does.Not.Contain("P@ssw0rd"));

            string decrypted = Sim44CredentialStore.DecryptPassword(cipher);
            Assert.That(decrypted, Is.EqualTo(plain));
        }

        [Test]
        public void 密码解密_空值或坏数据返回空而不是抛异常()
        {
            Assert.That(Sim44CredentialStore.DecryptPassword(null), Is.Null);
            Assert.That(Sim44CredentialStore.DecryptPassword(""), Is.Null);
            Assert.That(Sim44CredentialStore.DecryptPassword("这不是Base64"), Is.Null);
        }

        [Test]
        public void 工令行_JSON往返不丢字段()
        {
            Sim44Row row = new()
            {
                MoNo = "RTAH2026110001",
                LineNo = "A01",
                LineName = "A線",
                StartDate = "2026/11/27",
                PartNo = "123-4567-89",
                MoQty = "120",
                RmaNo = "RMA0001",
                RmaSeq = "1",
            };
            Sim44Row back = Sim44Row.FromJson(row.ToJson());
            Assert.That(back.MoNo, Is.EqualTo(row.MoNo));
            Assert.That(back.LineNo, Is.EqualTo(row.LineNo));
            Assert.That(back.LineName, Is.EqualTo(row.LineName));
            Assert.That(back.StartDate, Is.EqualTo(row.StartDate));
            Assert.That(back.PartNo, Is.EqualTo(row.PartNo));
            Assert.That(back.MoQty, Is.EqualTo(row.MoQty));
            Assert.That(back.RmaNo, Is.EqualTo(row.RmaNo));
            Assert.That(back.RmaSeq, Is.EqualTo(row.RmaSeq));
        }

        [Test]
        public void 工令行_缺字段时按空串处理()
        {
            Sim44Row back = Sim44Row.FromJson(new Newtonsoft.Json.Linq.JObject());
            Assert.That(back.MoNo, Is.Empty);
            Assert.That(back.MoQty, Is.Empty);
            Assert.That(back.StartDate, Is.Empty);
        }
    }
}
