using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Services;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 状态取值常量（报告状态 / 单体去向）与本地化资源的冒烟测试：
    /// 保证取值集合、锁定规则与四种语言资源中的关键键都能正常解析（不会显示为键名）。
    /// </summary>
    [TestFixture]
    public class KindAndResourceTests
    {
        /* ###############################  报告状态  ################################ */

        [Test]
        public void 报告状态_只允许三种取值()
            => Assert.That(ReportStatusKind.All, Is.EqualTo(new[] { "已完成", "进行中", "无要求" }));

        [Test]
        public void 报告状态_仅无要求被用户锁定()
        {
            Assert.That(ReportStatusKind.IsUserLocked(ReportStatusKind.NotRequired), Is.True);
            Assert.That(ReportStatusKind.IsUserLocked(ReportStatusKind.Complete), Is.False);
            Assert.That(ReportStatusKind.IsUserLocked(ReportStatusKind.InProgress), Is.False);
            Assert.That(ReportStatusKind.IsUserLocked(null), Is.False);
            Assert.That(ReportStatusKind.IsUserLocked(""), Is.False);
        }

        [Test]
        public void 报告状态_合法性校验()
        {
            Assert.That(ReportStatusKind.IsValid(null), Is.True);        // 未设置为合法
            Assert.That(ReportStatusKind.IsValid(""), Is.True);
            Assert.That(ReportStatusKind.IsValid(ReportStatusKind.Complete), Is.True);
            Assert.That(ReportStatusKind.IsValid(ReportStatusKind.InProgress), Is.True);
            Assert.That(ReportStatusKind.IsValid(ReportStatusKind.NotRequired), Is.True);
            Assert.That(ReportStatusKind.IsValid("随便写"), Is.False);
        }

        /* ###############################  单体去向  ################################ */

        [Test]
        public void 单体去向_两个取值与顺序()
        {
            Assert.That(RequisitionDispositionKind.StockIn, Is.EqualTo("入库"));
            Assert.That(RequisitionDispositionKind.Scrap, Is.EqualTo("报废"));
            Assert.That(RequisitionDispositionKind.All,
                Is.EqualTo(new[] { RequisitionDispositionKind.StockIn, RequisitionDispositionKind.Scrap }));
        }

        /* ###############################  本地化资源  ################################ */

        [Test]
        public void 关键本地化键_都能解析出文本()
        {
            string[] keys =
            [
                "Flow_Step_UnitReturn", "Flow_Ev_UnitReturnPending", "Flow_Branch_Planned",
                "Plans_UnitReturn", "Plans_Menu_UnitReturn", "Plans_Menu_GoToPlan", "Plans_Menu_GoToReq",
                "Msg_StockInNeedReturn", "Msg_StockInAfterScrap", "Msg_ScrapAfterStockIn",
                "Plans_Disposition", "ReqEdit_Disposition", "Msg_SelectDisposition",
                "Common_SaveAndContinue", "Common_SaveAndContinueHint",
                "ReqEdit_LastRevHintFormat", "ReqEdit_ModelHint", "ReqEdit_WorkOrderHint",
                "ReqEdit_QtySyncHint", "ReqEdit_PlanRemarkHint",
                "Plans_SearchHistory", "Plans_SearchHistoryHint", "Plans_SearchHistoryEmpty", "Plans_SearchClearHint",
                "Backup_Title", "Backup_Hint", "Backup_Folder", "Backup_OpenFolder", "Backup_FullNow",
                "Backup_IncrementalNow", "Backup_Col_Time", "Backup_Col_Kind", "Backup_Col_Size",
                "Backup_Col_Changed", "Backup_Col_Base", "Backup_Col_File", "Backup_Col_State",
                "Backup_Kind_Full", "Backup_Kind_Incremental", "Backup_State_Ok", "Backup_Restore",
                "Backup_Delete", "Backup_RestoreTitle", "Backup_RestoreConfirm", "Backup_RestoreBusy",
                "Backup_RestoreDoneTitle", "Backup_RestoreDone", "Backup_DeleteConfirm", "Backup_NoSelection",
                "Backup_Busy", "Backup_LastFormat", "Backup_None", "Backup_DeletedFormat",
                "Backup_DeletedPartialFormat", "Backup_Settings_Title", "Backup_Settings_Hint",
                "Backup_Settings_Auto", "Backup_Settings_Folder", "Backup_Settings_FolderHint",
                "Backup_Settings_Default", "Backup_Settings_Open", "Backup_Settings_NoBackupFormat",
                "ReqReturn_QtyHint", "Plans_StagedPreview", "Plans_StagedPreviewCleared",
                "Msg_JobNoTakenInListFormat",
            ];
            foreach (string key in keys)
            {
                string text = LanguageService.Get(key);
                Assert.That(text, Is.Not.Null.And.Not.Empty, $"键 {key} 未解析出文本");
                Assert.That(text, Is.Not.EqualTo(key), $"键 {key} 缺失（原样返回了键名）");
            }
        }

        [Test]
        public void 不存在的键_原样返回键名()
            => Assert.That(LanguageService.Get("__NoSuchKey_ForTest__"), Is.EqualTo("__NoSuchKey_ForTest__"));
    }
}
