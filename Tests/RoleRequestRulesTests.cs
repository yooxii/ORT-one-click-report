using NUnit.Framework;
using ORT一键报告.Models;
using ORT一键报告.Utils;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 身份组申请的规则测试：
    /// 1) 只能申请技术员 / 审核员，管理员不能申请；
    /// 2) 技术员申请审核员或管理员都能批，审核员申请只有管理员能批。
    /// </summary>
    [TestFixture]
    public class RoleRequestRulesTests
    {
        [Test]
        public void 可申请的身份组_只有技术员与审核员()
        {
            Assert.That(RoleRequestRules.IsRequestable(UserRole.Technician), Is.True);
            Assert.That(RoleRequestRules.IsRequestable(UserRole.Reviewer), Is.True);
            Assert.That(RoleRequestRules.IsRequestable(UserRole.Administrator), Is.False);
            Assert.That(RoleRequestRules.IsRequestable(UserRole.GeneralUser), Is.False);
            Assert.That(RoleRequestRules.IsRequestable(UserRole.Guest), Is.False);
        }

        [Test]
        public void 审核员身份的申请_只有管理员能通过()
        {
            Assert.That(RoleRequestRules.CanApprove(UserRole.Reviewer, approverIsAdmin: true, approverIsReviewer: true), Is.True);
            Assert.That(RoleRequestRules.CanApprove(UserRole.Reviewer, approverIsAdmin: false, approverIsReviewer: true), Is.False);
            Assert.That(RoleRequestRules.CanApprove(UserRole.Reviewer, approverIsAdmin: false, approverIsReviewer: false), Is.False);
        }

        [Test]
        public void 技术员身份的申请_审核员或管理员都能通过()
        {
            Assert.That(RoleRequestRules.CanApprove(UserRole.Technician, approverIsAdmin: false, approverIsReviewer: true), Is.True);
            Assert.That(RoleRequestRules.CanApprove(UserRole.Technician, approverIsAdmin: true, approverIsReviewer: false), Is.True);
            Assert.That(RoleRequestRules.CanApprove(UserRole.Technician, approverIsAdmin: false, approverIsReviewer: false), Is.False);
        }

        [Test]
        public void 管理员身份的申请_视同不可申请_只有管理员能批()
        {
            Assert.That(RoleRequestRules.CanApprove(UserRole.Administrator, approverIsAdmin: true, approverIsReviewer: false), Is.True);
            Assert.That(RoleRequestRules.CanApprove(UserRole.Administrator, approverIsAdmin: false, approverIsReviewer: true), Is.False);
        }

        [Test]
        public void 身份组显示名()
        {
            Assert.That(RoleRequestRules.DisplayName(UserRole.Technician), Is.EqualTo("技术员"));
            Assert.That(RoleRequestRules.DisplayName(UserRole.Reviewer), Is.EqualTo("审核员"));
            Assert.That(RoleRequestRules.DisplayName(UserRole.Administrator), Is.EqualTo("管理员"));
        }

        [Test]
        public void 申请内容_能解析出目标身份()
        {
            RoleRequestPayload payload = new() { RequestedRole = nameof(UserRole.Reviewer), Note = "需要审核领退表" };
            Assert.That(payload.ParsedRole, Is.EqualTo(UserRole.Reviewer));

            RoleRequestPayload broken = new() { RequestedRole = "NotARole" };
            Assert.That(broken.ParsedRole, Is.Null);
        }
    }
}
