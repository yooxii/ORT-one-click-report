using System;
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// Toast 提示窗口：无边框透明、不抢焦点、置顶，显示数秒后淡出自动关闭。
    /// 位置由 <see cref="PositionNear"/> 按设置（默认聚焦窗口右上角）计算。
    /// </summary>
    public partial class WindowToast : Window
    {
        /// <summary>显示时长（毫秒）</summary>
        private const int DisplayMs = 3000;

        private readonly DispatcherTimer _timer;

        public WindowToast(string message, ToastType type)
        {
            InitializeComponent();

            txt_msg.Text = message;

            // 类型 → 语义画刷键（随主题切换）
            string brushKey = type switch
            {
                ToastType.Success => "StatusOkBrush",
                ToastType.Warning => "StatusWarnBrush",
                ToastType.Error => "StatusErrorBrush",
                _ => "PrimaryBrush",
            };
            dot.SetResourceReference(System.Windows.Shapes.Shape.FillProperty, brushKey);
            rootBorder.SetResourceReference(System.Windows.Controls.Border.BorderBrushProperty, brushKey);

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(DisplayMs) };
            _timer.Tick += (s, e) => { _timer.Stop(); FadeOutAndClose(); };
        }

        /// <summary>
        /// 按位置设置把 Toast 放到目标窗口（默认聚焦窗口）的对应角落
        /// </summary>
        public void PositionNear(Window owner, string position)
        {
            UpdateLayout();
            double w = ActualWidth;
            double h = ActualHeight;
            const double margin = 14;

            if (owner == null)
            {
                // 无Owner时放主屏幕右上角
                var wa = SystemParameters.WorkArea;
                Left = wa.Right - w - margin;
                Top = wa.Top + margin;
                return;
            }

            double ol = owner.Left, ot = owner.Top, ow = owner.ActualWidth, oh = owner.ActualHeight;
            // 顶部预留标题栏高度，避免遮挡系统标题栏
            const double caption = 36;

            switch (position)
            {
                case "TopLeft":
                    Left = ol + margin;
                    Top = ot + caption + margin;
                    break;
                case "BottomRight":
                    Left = ol + ow - w - margin;
                    Top = ot + oh - h - margin;
                    break;
                case "BottomLeft":
                    Left = ol + margin;
                    Top = ot + oh - h - margin;
                    break;
                case "TopRight":
                default:
                    Left = ol + ow - w - margin;
                    Top = ot + caption + margin;
                    break;
            }
        }

        public void BeginDisplay()
        {
            Show();
            _timer.Start();
        }

        private void FadeOutAndClose()
        {
            var anim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
            };
            anim.Completed += (s, e) => Close();
            BeginAnimation(OpacityProperty, anim);
        }
    }

    /// <summary>Toast 类型（决定强调色）</summary>
    public enum ToastType
    {
        Info,
        Success,
        Warning,
        Error,
    }
}
