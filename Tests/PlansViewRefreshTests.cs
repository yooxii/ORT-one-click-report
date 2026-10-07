using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Windows.Data;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 表格视图刷新（PlansViewModel.SafeRefresh）测试：
    /// 单元格还在编辑（AddNew/EditItem 事务未结束）时 CollectionView 不允许 Refresh，
    /// 直接刷新会抛 InvalidOperationException；批量登记窗口关闭时会清筛选并刷新，
    /// 若此刻主表格正在编辑就会走到这条路径——必须不抛异常（延后重试），否则整个程序会崩。
    /// </summary>
    [TestFixture]
    public class PlansViewRefreshTests
    {
        [Test]
        public void 安全刷新_编辑事务未结束时_不抛异常()
        {
            ObservableCollection<Requisition> rows = [new Requisition { Id = 1, RequisitionNo = "WL-001" }];
            ListCollectionView view = new(rows);
            view.EditItem(rows[0]);

            // 前提：这种状态下直接刷新确实会被拒绝（正是用户遇到的报错）
            Assert.Throws<InvalidOperationException>(() => view.Refresh());

            Assert.DoesNotThrow(() => PlansViewModel.SafeRefresh(view));
        }

        [Test]
        public void 安全刷新_正常状态_照常刷新()
        {
            ObservableCollection<Requisition> rows = [new Requisition { Id = 1, RequisitionNo = "WL-001" }];
            ListCollectionView view = new(rows);

            Assert.DoesNotThrow(() => PlansViewModel.SafeRefresh(view));
        }

        [Test]
        public void 安全刷新_空视图_不抛异常()
        {
            Assert.DoesNotThrow(() => PlansViewModel.SafeRefresh(null));
        }
    }
}
