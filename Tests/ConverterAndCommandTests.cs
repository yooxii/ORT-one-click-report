using NUnit.Framework;
using ORT一键报告.Converters;
using System;
using System.Globalization;

namespace ORT一键报告.Tests
{
    /// <summary>
    /// 值转换器（InverseBoolConverter）与命令（RelayCommand）测试：
    /// UI 绑定层的纯逻辑部分，不依赖窗口即可验证。
    /// </summary>
    [TestFixture]
    public class ConverterAndCommandTests
    {
        private readonly InverseBoolConverter _converter = new();
        private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;

        /* ###############################  布尔取反转换器  ################################ */

        [TestCase(true, false)]
        [TestCase(false, true)]
        public void 取反转换_布尔值(bool input, bool expected)
            => Assert.That(_converter.Convert(input, typeof(bool), null, Culture), Is.EqualTo(expected));

        [Test]
        public void 取反转换_非布尔或空值视为false()
        {
            Assert.That(_converter.Convert(null, typeof(bool), null, Culture), Is.EqualTo(true));
            Assert.That(_converter.Convert("x", typeof(bool), null, Culture), Is.EqualTo(true));
        }

        [TestCase(true, false)]
        [TestCase(false, true)]
        public void 取反转换_反向(bool input, bool expected)
            => Assert.That(_converter.ConvertBack(input, typeof(bool), null, Culture), Is.EqualTo(expected));

        [Test]
        public void 取反转换_反向非布尔值返回false()
            => Assert.That(_converter.ConvertBack(null, typeof(bool), null, Culture), Is.EqualTo(false));

        /* ###############################  命令  ################################ */

        [Test]
        public void 命令_无CanExecute时始终可执行()
        {
            int executed = 0;
            RelayCommand command = new(() => executed++);

            Assert.That(command.CanExecute(null), Is.True);
            command.Execute(null);
            Assert.That(executed, Is.EqualTo(1));
        }

        [Test]
        public void 命令_CanExecute随判定函数变化()
        {
            bool allow = false;
            RelayCommand command = new(() => { }, () => allow);

            Assert.That(command.CanExecute(null), Is.False);
            allow = true;
            Assert.That(command.CanExecute(null), Is.True);
        }

        [Test]
        public void 命令_空执行体_构造时即报错()
        {
            Assert.Throws<ArgumentNullException>(() => _ = new RelayCommand((Action)null));
        }
    }
}
