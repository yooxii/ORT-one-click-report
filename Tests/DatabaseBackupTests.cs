using NUnit.Framework;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 数据库备份的纯规则测试：备份文件名的生成/解析、逐行差异比对（新增/修改/删除）、
    /// 取值编码与 SQL 字面量转义、增量还原链的组装与缺链判定。
    /// </summary>
    [TestFixture]
    public class DatabaseBackupTests
    {
        private static readonly DateTime Sample = new(2026, 10, 9, 15, 30, 0);

        /* ###############################  文件名  ################################ */

        [Test]
        public void 全量文件名_可解析回时间与类型()
        {
            string name = DatabaseBackupDiffer.FullFileName(Sample);
            Assert.That(name, Is.EqualTo("ort_plans_full_20261009_153000.db"));
            Assert.That(DatabaseBackupDiffer.TryParseFileName(name, out DatabaseBackupKind kind, out DateTime time), Is.True);
            Assert.That(kind, Is.EqualTo(DatabaseBackupKind.Full));
            Assert.That(time, Is.EqualTo(Sample));
        }

        [Test]
        public void 增量文件名_可解析回时间与类型()
        {
            string name = DatabaseBackupDiffer.IncrementalFileName(Sample);
            Assert.That(name, Is.EqualTo("ort_plans_inc_20261009_153000.json"));
            Assert.That(DatabaseBackupDiffer.TryParseFileName(name, out DatabaseBackupKind kind, out DateTime time), Is.True);
            Assert.That(kind, Is.EqualTo(DatabaseBackupKind.Incremental));
            Assert.That(time, Is.EqualTo(Sample));
        }

        [TestCase(null)]
        [TestCase("")]
        [TestCase("ort_plans.db")]
        [TestCase("ort_plans_full_20261009.db")]        // 缺时分秒
        [TestCase("ort_plans_inc_20261009_1530.json")]  // 秒位缺失
        [TestCase("backup_20261009_153000.db")]
        public void 认不出的文件名_返回false(string name)
            => Assert.That(DatabaseBackupDiffer.TryParseFileName(name, out _, out _), Is.False);

        /* ###############################  逐行差异  ################################ */

        private static DatabaseBackupTableSnapshot Table(params (long Key, object[] Values)[] rows)
        {
            DatabaseBackupTableSnapshot snapshot = new()
            {
                Table = "plans",
                KeyColumn = "Id",
                Columns = ["Id", "JobNo", "ModelName"]
            };
            foreach ((long key, object[] values) in rows)
            {
                snapshot.Rows[key] = values;
            }
            return snapshot;
        }

        [Test]
        public void 差异_新增修改删除都能识别()
        {
            DatabaseBackupTableSnapshot before = Table(
                (1, [1L, "RT2609001", "A"]),
                (2, [2L, "RT2609002", "B"]),
                (3, [3L, "RT2609003", "C"]));
            DatabaseBackupTableSnapshot after = Table(
                (1, [1L, "RT2609001", "A"]),      // 没变
                (2, [2L, "RT2609011", "B"]),      // 编号改了（增大一位）
                (4, [4L, "RT2609004", "D"]));     // 新增

            DatabaseBackupTablePatch patch = DatabaseBackupDiffer.DiffTable(before, after);

            Assert.That(patch.ChangedCount, Is.EqualTo(3));
            Assert.That(patch.Rows.Count, Is.EqualTo(2));
            Assert.That(patch.DeletedKeys, Is.EqualTo(new List<long> { 3 }));
            Assert.That(patch.Rows.Exists(r => r.Key == 2), Is.True);
            Assert.That(patch.Rows.Exists(r => r.Key == 4), Is.True);
            Assert.That(patch.KeyColumn, Is.EqualTo("Id"));
        }

        [Test]
        public void 差异_没有变化时为空补丁()
        {
            DatabaseBackupTableSnapshot before = Table((1, [1L, "RT2609001", "A"]));
            DatabaseBackupTableSnapshot after = Table((1, [1L, "RT2609001", "A"]));
            Assert.That(DatabaseBackupDiffer.DiffTable(before, after).ChangedCount, Is.EqualTo(0));
        }

        [Test]
        public void 差异_没有基准备份时整表都算新增()
        {
            DatabaseBackupTablePatch patch = DatabaseBackupDiffer.DiffTable(null, Table((1, [1L, "RT2609001", "A"])));
            Assert.That(patch.ChangedCount, Is.EqualTo(1));
            Assert.That(patch.DeletedKeys, Is.Empty);
        }

        [Test]
        public void 行签名_不同类型与空值不会混淆()
        {
            Assert.That(DatabaseBackupDiffer.RowSignature([1L, null]), Is.Not.EqualTo(DatabaseBackupDiffer.RowSignature([1L, ""])));
            Assert.That(DatabaseBackupDiffer.RowSignature([1L]), Is.Not.EqualTo(DatabaseBackupDiffer.RowSignature(["1"])));
            Assert.That(DatabaseBackupDiffer.RowSignature([1.0d]), Is.Not.EqualTo(DatabaseBackupDiffer.RowSignature([1L])));
            Assert.That(DatabaseBackupDiffer.RowSignature([1L, "A"]), Is.EqualTo(DatabaseBackupDiffer.RowSignature([1L, "A"])));
        }

        /* ###############################  取值编码  ################################ */

        [Test]
        public void 取值的SQL字面量_文本转义与空值()
        {
            Assert.That(DatabaseBackupDiffer.ValueLiteral(null), Is.EqualTo("NULL"));
            Assert.That(DatabaseBackupDiffer.ValueLiteral(12L), Is.EqualTo("12"));
            Assert.That(DatabaseBackupDiffer.ValueLiteral("O'Brien"), Is.EqualTo("'O''Brien'"));
            Assert.That(DatabaseBackupDiffer.ValueLiteral(""), Is.EqualTo("''"));
        }

        [Test]
        public void 二进制取值_编码后能还原成BLOB字面量()
        {
            byte[] blob = [1, 2, 255];
            object stored = DatabaseBackupDiffer.ValueForStorage(blob);
            Assert.That(stored, Is.TypeOf<string>());
            Assert.That(DatabaseBackupDiffer.ValueLiteral(stored), Is.EqualTo("X'0102FF'"));
        }

        /* ###############################  还原链  ################################ */

        private static DatabaseBackupEntry Entry(string fileName, DatabaseBackupKind kind, string baseFileName = null)
            => new()
            {
                FileName = fileName,
                FilePath = @"C:\Backups\" + fileName,
                Kind = kind,
                BaseFileName = baseFileName,
                CreatedAt = Sample
            };

        [Test]
        public void 还原链_从全量一路回放到选中的增量()
        {
            DatabaseBackupEntry full = Entry("ort_plans_full_20261001_090000.db", DatabaseBackupKind.Full);
            DatabaseBackupEntry inc1 = Entry("ort_plans_inc_20261002_090000.json", DatabaseBackupKind.Incremental, full.FileName);
            DatabaseBackupEntry inc2 = Entry("ort_plans_inc_20261003_090000.json", DatabaseBackupKind.Incremental, inc1.FileName);
            Dictionary<string, DatabaseBackupEntry> byName = new(StringComparer.OrdinalIgnoreCase)
            {
                [full.FileName] = full,
                [inc1.FileName] = inc1,
                [inc2.FileName] = inc2
            };

            List<DatabaseBackupEntry> chain = DatabaseBackupDiffer.BuildRestoreChain(byName, inc2, out string problem);

            Assert.That(problem, Is.Null);
            Assert.That(chain, Has.Count.EqualTo(3));
            Assert.That(chain[0].FileName, Is.EqualTo(full.FileName));
            Assert.That(chain[1].FileName, Is.EqualTo(inc1.FileName));
            Assert.That(chain[2].FileName, Is.EqualTo(inc2.FileName));
        }

        [Test]
        public void 还原链_全量备份自己就是一条链()
        {
            DatabaseBackupEntry full = Entry("ort_plans_full_20261001_090000.db", DatabaseBackupKind.Full);
            Dictionary<string, DatabaseBackupEntry> byName = new(StringComparer.OrdinalIgnoreCase) { [full.FileName] = full };
            List<DatabaseBackupEntry> chain = DatabaseBackupDiffer.BuildRestoreChain(byName, full, out string problem);
            Assert.That(problem, Is.Null);
            Assert.That(chain, Has.Count.EqualTo(1));
        }

        [Test]
        public void 还原链_基准备份缺失时判定不可还原()
        {
            DatabaseBackupEntry inc = Entry("ort_plans_inc_20261003_090000.json", DatabaseBackupKind.Incremental,
                "ort_plans_full_20261001_090000.db");
            Dictionary<string, DatabaseBackupEntry> byName = new(StringComparer.OrdinalIgnoreCase) { [inc.FileName] = inc };

            Assert.That(DatabaseBackupDiffer.BuildRestoreChain(byName, inc, out string problem), Is.Null);
            Assert.That(problem, Does.Contain("ort_plans_full_20261001_090000.db"));
        }

        [Test]
        public void 还原链_增量没有记录基准时判定不可还原()
        {
            DatabaseBackupEntry inc = Entry("ort_plans_inc_20261003_090000.json", DatabaseBackupKind.Incremental);
            Dictionary<string, DatabaseBackupEntry> byName = new(StringComparer.OrdinalIgnoreCase) { [inc.FileName] = inc };
            Assert.That(DatabaseBackupDiffer.BuildRestoreChain(byName, inc, out string problem), Is.Null);
            Assert.That(problem, Is.Not.Null.And.Not.Empty);
        }

        [Test]
        public void 文件大小文本_按KB与MB显示()
        {
            Assert.That(DatabaseBackupEntry.FormatSize(512), Is.EqualTo("512 B"));
            Assert.That(DatabaseBackupEntry.FormatSize(2048), Is.EqualTo("2.0 KB"));
            Assert.That(DatabaseBackupEntry.FormatSize(3 * 1024 * 1024), Is.EqualTo("3.0 MB"));
        }
    }
}
