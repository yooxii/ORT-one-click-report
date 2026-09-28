using NUnit.Framework;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 序列号核对（SnVerification）测试：文本拆分、清单比对（多/少/重复/空）、
    /// 文件提取（txt 按行、csv 按列识别并跳过表头、非法内容拒绝）与附件+文本并集。
    /// </summary>
    [TestFixture]
    public class SnVerificationTests
    {
        private readonly List<string> _tempFiles = [];

        [TearDown]
        public void Cleanup()
        {
            foreach (string file in _tempFiles.Where(File.Exists))
            {
                try { File.Delete(file); } catch { /* 测试清理失败不阻塞 */ }
            }
            _tempFiles.Clear();
        }

        private string WriteTemp(string content, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), $"ort_sn_test_{Guid.NewGuid():N}{extension}");
            File.WriteAllText(path, content);
            _tempFiles.Add(path);
            return path;
        }

        /* ###############################  文本拆分  ################################ */

        [Test]
        public void 文本拆分_支持换行逗号分号制表符顿号()
        {
            List<string> list = SnVerification.ParseText("A1\nA2；a3,A4、A5\tA6\r\n");

            Assert.That(list, Is.EqualTo(new[] { "A1", "A2", "a3", "A4", "A5", "A6" }));
        }

        [Test]
        public void 文本拆分_去空白行并保留重复()
        {
            List<string> list = SnVerification.ParseText("  X1 \r\n\r\n  X1  \n   ");

            Assert.That(list, Is.EqualTo(new[] { "X1", "X1" }));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("   \n  \t ")]
        public void 文本拆分_空白输入返回空列表(string text)
            => Assert.That(SnVerification.ParseText(text), Is.Empty);

        /* ###############################  清单比对  ################################ */

        [Test]
        public void 比对_完全一致_顺序与大小写无关()
        {
            var result = SnVerification.Compare(["ACBEL-001", "ACBEL-002"], ["acbel-002", "ACBEL-001"]);

            Assert.That(result.Ok, Is.True);
            Assert.That(result.RequisitionCount, Is.EqualTo(2));
            Assert.That(result.ScrapCount, Is.EqualTo(2));
            Assert.That(result.NotInRequisition, Is.Empty);
            Assert.That(result.MissingInScrap, Is.Empty);
            Assert.That(result.DuplicatesInScrap, Is.Empty);
        }

        [Test]
        public void 比对_报废清单多一个_不通过并列出()
        {
            var result = SnVerification.Compare(["ACBEL-001"], ["ACBEL-001", "ACBEL-009"]);

            Assert.That(result.Ok, Is.False);
            Assert.That(result.NotInRequisition, Is.EqualTo(new[] { "ACBEL-009" }));
            Assert.That(result.MissingInScrap, Is.Empty);
        }

        [Test]
        public void 比对_漏报废一个_不通过并列出()
        {
            var result = SnVerification.Compare(["ACBEL-001", "ACBEL-002"], ["ACBEL-001"]);

            Assert.That(result.Ok, Is.False);
            Assert.That(result.MissingInScrap, Is.EqualTo(new[] { "ACBEL-002" }));
            Assert.That(result.NotInRequisition, Is.Empty);
        }

        [Test]
        public void 比对_报废清单内重复_不通过并给出次数()
        {
            var result = SnVerification.Compare(["ACBEL-001"], ["ACBEL-001", "acbel-001"]);

            Assert.That(result.Ok, Is.False);
            Assert.That(result.DuplicatesInScrap.Any(d => d.Sn == "ACBEL-001" && d.Count == 2), Is.True);
        }

        [Test]
        public void 比对_任一清单为空_不通过并标记()
        {
            var emptyReq = SnVerification.Compare([], ["ACBEL-001"]);
            Assert.That(emptyReq.Ok, Is.False);
            Assert.That(emptyReq.RequisitionListEmpty, Is.True);

            var emptyScrap = SnVerification.Compare(["ACBEL-001"], []);
            Assert.That(emptyScrap.Ok, Is.False);
            Assert.That(emptyScrap.ScrapListEmpty, Is.True);
        }

        /* ###############################  文件提取  ################################ */

        [Test]
        public void 提取_文件不存在_抛出明确异常()
        {
            string missing = Path.Combine(Path.GetTempPath(), $"ort_missing_{Guid.NewGuid():N}.txt");
            Assert.Throws<FileNotFoundException>(() => SnVerification.ExtractFromFile(missing));
        }

        [Test]
        public void 提取_txt_按行取值()
        {
            string path = WriteTemp("ACBEL-0001\nACBEL-0002\nACBEL-0003\n", ".txt");

            List<string> list = SnVerification.ExtractFromFile(path);

            Assert.That(list, Is.EqualTo(new[] { "ACBEL-0001", "ACBEL-0002", "ACBEL-0003" }));
        }

        [Test]
        public void 提取_txt_重复序列号_拒绝()
        {
            string path = WriteTemp("ACBEL-0001\nACBEL-0001\n", ".txt");

            Assert.Throws<InvalidDataException>(() => SnVerification.ExtractFromFile(path));
        }

        [Test]
        public void 提取_csv_识别序列号列并跳过表头()
        {
            string path = WriteTemp("SN,Note\nACBEL-0001,第一台\nACBEL-0002,第二台\n", ".csv");

            List<string> list = SnVerification.ExtractFromFile(path);

            Assert.That(list, Is.EqualTo(new[] { "ACBEL-0001", "ACBEL-0002" }));
        }

        [Test]
        public void 提取_csv_无合格列_拒绝()
        {
            string path = WriteTemp("Name,Value\nfoo,1\nfoo,2\n", ".csv");

            Assert.Throws<InvalidDataException>(() => SnVerification.ExtractFromFile(path));
        }

        /* ###############################  领用清单（附件 + 文本）  ################################ */

        [Test]
        public void 领用清单_附件与文本并集去重()
        {
            string path = WriteTemp("ACBEL-0001\nACBEL-0002\n", ".txt");

            List<string> list = SnVerification.ExtractFromRequisition("ACBEL-0002\nACBEL-0003", path);

            Assert.That(list, Is.EquivalentTo(new[] { "ACBEL-0001", "ACBEL-0002", "ACBEL-0003" }));
            Assert.That(list.Count, Is.EqualTo(3));
        }

        [Test]
        public void 领用清单_附件不存在_退回文本()
        {
            string missing = Path.Combine(Path.GetTempPath(), $"ort_missing_{Guid.NewGuid():N}.txt");

            List<string> list = SnVerification.ExtractFromRequisition("ACBEL-0001\nACBEL-0002", missing);

            Assert.That(list, Is.EqualTo(new[] { "ACBEL-0001", "ACBEL-0002" }));
        }

        [Test]
        public void 领用清单_两者都取不到_返回空()
            => Assert.That(SnVerification.ExtractFromRequisition(null, null), Is.Empty);
    }
}
