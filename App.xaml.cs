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

                // EPPlus 非商业许可统一在程序入口设置（各服务不再重复设置）
                OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Lucas");

                // 初始化语言服务（读取上次保存的语言或使用系统语言）
                ORT一键报告.Services.LanguageService.Initialize();

                // 初始化 UI 主题（读取上次保存的方案，默认 Fluent）
                ORT一键报告.Services.ThemeService.Initialize();

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
                services.AddSingleton<ReportGenerationService>();

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

