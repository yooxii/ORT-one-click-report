using System;
using NLog;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

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
            try { base.OnStartup(e); }
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
