using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Reports.ViewModels;
using ORT一键报告.Reports.Views;
using ORT一键报告.Services;
using ORT一键报告.ViewModels;
using System;
using System.Windows;

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

                ServiceCollection services = new();
                // Services
                services.AddSingleton<IPathService, PathService>();
                services.AddSingleton<ReportService>();

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

                ServiceProvider = services.BuildServiceProvider();
            }
            catch (Exception ex)
            {
                logger.Fatal(ex, "程序启动失败");
                throw;
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            logger.Info("程序退出");
            LogManager.Shutdown();
            base.OnExit(e);
        }
    }
}
