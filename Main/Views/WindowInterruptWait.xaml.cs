using ORT一键报告.Services;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowInterruptWait.xaml 的交互逻辑：高耗时任务互斥时弹出的「正在中断」等待窗。
    /// 非阻塞 Show，由 <see cref="HighCostTaskCoordinator"/> 在被中断任务真正结束后 Close。
    /// </summary>
    public partial class WindowInterruptWait : Window
    {
        public WindowInterruptWait(string interruptedTaskName)
        {
            InitializeComponent();
            string format = LanguageService.Get("Interrupt_WaitMessageFormat");
            txt_message.Text = string.IsNullOrWhiteSpace(format)
                ? $"正在中断「{interruptedTaskName}」，等当前报告夹读取完成后继续…"
                : string.Format(format, interruptedTaskName ?? "");
        }
    }
}
