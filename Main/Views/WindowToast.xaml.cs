using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// Toast 提示窗口：无边框透明、不抢焦点、置顶，显示数秒后淡出自动关闭。
    /// 位置固定为「窗口内部居中」（没有可见窗口时放屏幕正中）。
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
        /// 显示并把 Toast 放到目标窗口（默认聚焦窗口）的**内部居中**位置，再开始计时淡出。
        /// 必须先 Show 再定位：SizeToContent 的 ActualWidth/Height 在显示前是 0，
        /// 早算会把它摆到窗口外（右侧），正是之前「跑到窗口外面右边」的原因。
        /// </summary>
        public void BeginDisplay(Window owner)
        {
            Show();
            UpdateLayout();
            PositionInCenter(owner);
            _timer.Start();
        }

        /// <summary>在 owner 窗口内部居中；owner 不可见/为空时放主屏幕正中</summary>
        private void PositionInCenter(Window owner)
        {
            double w = ActualWidth > 0 ? ActualWidth : DesiredSize.Width;
            double h = ActualHeight > 0 ? ActualHeight : DesiredSize.Height;
            if (double.IsNaN(w) || double.IsNaN(h) || w <= 0 || h <= 0)
            {
                w = 360;
                h = 60;
            }

            if (owner == null || !owner.IsVisible || owner.ActualWidth <= 0 || owner.ActualHeight <= 0)
            {
                Rect area = SystemParameters.WorkArea;
                Left = area.Left + (area.Width - w) / 2;
                Top = area.Top + (area.Height - h) / 2;
                return;
            }

            try
            {
                // owner 可能最大化 / 跨屏：用屏幕坐标取它的矩形，再按 DPI 换算成 DIP，
                // 这样不依赖 Window.Left/Top（最大化时它们不代表实际位置）
                PresentationSource source = PresentationSource.FromVisual(owner);
                Matrix toDip = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
                Point topLeft = toDip.Transform(owner.PointToScreen(new Point(0, 0)));
                Point bottomRight = toDip.Transform(owner.PointToScreen(new Point(owner.ActualWidth, owner.ActualHeight)));
                double width = Math.Max(1, bottomRight.X - topLeft.X);
                double height = Math.Max(1, bottomRight.Y - topLeft.Y);
                Left = topLeft.X + Math.Max(0, (width - w) / 2);
                Top = topLeft.Y + Math.Max(0, (height - h) / 2);
            }
            catch
            {
                // 取屏幕坐标失败时退回按窗口 Left/Top 估算，不影响提示本身
                Left = owner.Left + Math.Max(0, (owner.ActualWidth - w) / 2);
                Top = owner.Top + Math.Max(0, (owner.ActualHeight - h) / 2);
            }
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
