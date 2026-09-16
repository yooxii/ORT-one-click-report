using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Plans.ViewModels;
using ORT一键报告.Reports.ViewModels;
using ORT一键报告.Reports.Views;
using ORT一键报告.Services;
using ORT一键报告.ViewModels;
using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using WPFLocalizeExtension.Engine;

namespace ORT一键报告
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static IServiceProvider ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            logger.Info("ORT一键报告程序启动");
            try
            {
                base.OnStartup(e);

                // Excel 读写统一使用 NPOI（Apache-2.0，无需任何许可调用/无写入署名）；
                // OLE 附件嵌入由 Utils/ExcelOleEmbedder 走 Excel COM 完成。

                // 初始化语言服务（读取上次保存的语言或使用系统语言）
                ORT一键报告.Services.LanguageService.Initialize();

                // 初始化 UI 主题（读取上次保存的方案，默认 Fluent）
                ORT一键报告.Services.ThemeService.Initialize();

                // 深色主题时统一深色标题栏/窗口边框（DWM 沉浸式深色模式）：
                // 每个窗口 Loaded 时自动应用（附加类处理器，覆盖所有 Window，无需逐窗口接线）。
                // Win10 活动标题栏缓存问题由 ApplyToWindow 内的 1px 宽度抖动强制重绘解决。
                EventManager.RegisterClassHandler(typeof(Window), FrameworkElement.LoadedEvent,
                    new RoutedEventHandler((s, args) =>
                    {
                        if (s is not Window win)
                        {
                            return;
                        }
                        // 开窗防闪：窗口先置为透明，等首帧渲染完成再恢复不透明。
                        // 否则重内容窗口（设置界面：目录树 + 大量表单 + 字体下拉）会先出现一个空白/半成品窗口，
                        // 再分几帧把内容补上，看起来就是"点开闪几次"。每个窗口只处理一次。
                        if (!win.AllowsTransparency && win.Opacity >= 1.0 && !(bool)win.GetValue(FlashGuardProperty))
                        {
                            win.SetValue(FlashGuardProperty, true);
                            win.Opacity = 0;
                            win.ContentRendered += WindowContentRenderedOnce;
                        }
                        ORT一键报告.Services.WindowThemeHelper.ApplyToWindow(win);
                        // 字体/字号/字重设置对所有窗口生效
                        if (ServiceProvider?.GetService(typeof(AppSettingsService)) is AppSettingsService settings)
                        {
                            settings.ApplyFont(win);
                        }
                    }));
                // 主题运行时切换：对所有已打开窗口重新应用
                ORT一键报告.Services.ThemeService.ThemeChanged +=
                    ORT一键报告.Services.WindowThemeHelper.ApplyToAllWindows;

                // 使用自定义本地化提供程序（直接读取 Resources.Strings 资源），
                // 解决 WPFLocalizeExtension 内置 Provider 对含中文程序集名解析失败、UI 显示 Key:xxx 的问题。
                WPFLocalizeExtension.Engine.LocalizeDictionary.Instance.DefaultProvider =
                    new ORT一键报告.Services.OrtLocalizationProvider();

                ServiceCollection services = new();
                // Services
                services.AddSingleton<IPathService, PathService>();
                services.AddSingleton<AppSettingsService>();
                services.AddSingleton<ReportService>();
                services.AddSingleton<DatabaseService>();
                services.AddSingleton<AuthService>();
                services.AddSingleton<IPermissionService, PermissionService>();
                services.AddSingleton<PlanExcelService>();
                services.AddSingleton<AdminService>();
                services.AddSingleton<ReviewService>();
                services.AddSingleton<MailService>();
                services.AddSingleton<MailNotifier>();
                services.AddSingleton<ReportGenerationService>();
                services.AddSingleton<ReportTemplateService>();
                services.AddSingleton<TestPlanService>();
                services.AddSingleton<PlanIndexService>();
                services.AddSingleton<PlanIndexScheduler>();

                // ViewModels
                services.AddTransient<MainViewModel>();
                services.AddTransient<MainReportViewModel>();
                services.AddTransient<BaseReportPageViewModel>();
                services.AddTransient<EMIReportViewModel>();
                services.AddTransient<EMISetupViewModel>();
                services.AddSingleton<SettingsViewModel>();
                services.AddTransient<MainSettingsViewModel>();
                services.AddTransient<ReturnLineViewModel>();
                services.AddTransient<ReturnLineSingleViewModel>();
                services.AddTransient<PlansViewModel>();

                ServiceProvider = services.BuildServiceProvider();
            }
            catch (Exception ex)
            {
                logger.Fatal(ex, "程序启动失败");
                throw;
            }
        }

        /// <summary>
        /// 开窗防闪标记（附加属性）：每个窗口只做一次透明处理，重复 Loaded 不会再次置空
        /// </summary>
        private static readonly DependencyProperty FlashGuardProperty =
            DependencyProperty.RegisterAttached("FlashGuard", typeof(bool), typeof(App), new PropertyMetadata(false));

        /// <summary>
        /// 首帧渲染完成后恢复不透明（配合开窗防闪处理；只订阅一次，触发后即解除）
        /// </summary>
        private static void WindowContentRenderedOnce(object sender, EventArgs e)
        {
            if (sender is not Window window)
            {
                return;
            }
            window.ContentRendered -= WindowContentRenderedOnce;
            window.Opacity = 1;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            logger.Info("程序退出");
            try
            {
                Utils.Report.ClearTempDir();
            }
            catch (Exception ex)
            {
                logger.Warn($"清理临时目录失败: {ex.Message}");
            }
            (ServiceProvider as IDisposable)?.Dispose();
            LogManager.Shutdown();
            base.OnExit(e);
        }
    }
}

