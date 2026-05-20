using Microsoft.Extensions.DependencyInjection;
using NLog;
using System;
using System.Windows;
using ORT一键报告.ViewModels;

namespace ORT一键报告
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        protected override void OnStartup(StartupEventArgs e)
        {
            logger.Info("ORT一键报告程序启动");
            try
            {
                base.OnStartup(e);

                ServiceCollection services = new();
                services.AddTransient<EMIViewModel>();
                services.AddTransient<BaseReportPageViewModel>();
                services.AddSingleton<ReportHeaderViewModel>();
                services.AddSingleton<IEMIService, EMIService>();
                ServiceProvider serviceProvider = services.BuildServiceProvider();

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
