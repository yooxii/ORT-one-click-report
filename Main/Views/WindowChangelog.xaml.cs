using NLog;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace ORT一键报告.Main.Views
{
    /// <summary>
    /// WindowChangelog.xaml 的交互逻辑：只读展示随程序发布的「更新日志.md」。
    /// 用 MdXaml 把 markdown 渲染为 WPF FlowDocument；只保留每个「## 版本号」小节及其更新条目
    /// （去掉文件顶部的标题与「&gt;」说明等元信息），并把版本反转为“最新在上”。
    /// </summary>
    public partial class WindowChangelog : Window
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public WindowChangelog()
        {
            InitializeComponent();
            Loaded += (s, e) => LoadChangelog();
        }

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 读取 exe 同级的「更新日志.md」（csproj 已登记随生成复制到输出目录），渲染到窗口
        /// </summary>
        private void LoadChangelog()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "更新日志.md");
            if (!File.Exists(path))
            {
                ShowMessage(string.Format(LanguageService.Get("Changelog_FileNotFoundFormat"), path));
                return;
            }
            try
            {
                string md = RebuildNewestFirst(File.ReadAllText(path, Encoding.UTF8));
                if (string.IsNullOrWhiteSpace(md))
                {
                    ShowMessage(LanguageService.Get("Changelog_Empty"));
                    return;
                }
                // 设置 Markdown 属性即触发 MdXaml 渲染（内部 Transform → Document）
                md_viewer.Markdown = md;
            }
            catch (Exception ex)
            {
                _logger.Warn($"加载更新日志失败: {ex.Message}");
                ShowMessage(string.Format(LanguageService.Get("Changelog_LoadFailedFormat"), ex.Message));
            }
        }

        /// <summary>
        /// 覆盖纸张显示一条提示信息（读取失败 / 无内容），并隐藏 markdown 视图
        /// </summary>
        private void ShowMessage(string message)
        {
            txt_message.Text = message;
            txt_message.Visibility = Visibility.Visible;
            md_viewer.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// 解析更新日志：丢弃文件顶部的 H1 标题与「&gt;」说明等元信息，只保留每个「## 版本号」小节
        /// （版本号 + 该版更新条目）；文件里最新版在最下方，这里反转为“最新在上”。
        /// </summary>
        internal static string RebuildNewestFirst(string raw)
        {
            if (string.IsNullOrEmpty(raw))
            {
                return string.Empty;
            }
            string[] lines = raw.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            List<List<string>> sections = [];
            List<string> current = null;
            foreach (string line in lines)
            {
                if (line.StartsWith("## ", StringComparison.Ordinal))
                {
                    // 一个版本小节开始（含版本号标题行）
                    current = [line];
                    sections.Add(current);
                }
                else if (line.StartsWith("# ", StringComparison.Ordinal))
                {
                    current = null; // H1 标题：忽略
                }
                else if (current != null)
                {
                    current.Add(line);
                }
                // 第一个「##」之前的内容（标题、说明、空行）一律忽略
            }
            if (sections.Count == 0)
            {
                return string.Empty;
            }
            sections.Reverse(); // 最新在上
            StringBuilder sb = new();
            foreach (List<string> sec in sections)
            {
                // 去掉小节尾部多余空行，小节之间统一空一行
                while (sec.Count > 0 && string.IsNullOrWhiteSpace(sec[sec.Count - 1]))
                {
                    sec.RemoveAt(sec.Count - 1);
                }
                _ = sb.AppendLine(string.Join("\n", sec));
                _ = sb.AppendLine();
            }
            return sb.ToString();
        }
    }
}
