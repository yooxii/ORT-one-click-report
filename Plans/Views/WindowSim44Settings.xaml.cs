using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Services.Sim44;
using System;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// MES 连接设置：数据源别名、工令表名、桥接程序路径。
    /// 之所以有这个窗口——登录成功却报「表不存在」时，需要能改表名（例如 SIM44 改 APITECH.SIM44），
    /// 以前只能手改 local_settings.json，用户根本找不到。
    /// 设置存本机（local_settings.json 的 sim44 段），保存后重新登录生效。
    /// </summary>
    public partial class WindowSim44Settings : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly Sim44Service _service;

        public WindowSim44Settings()
        {
            InitializeComponent();
            _service = App.ServiceProvider.GetRequiredService<Sim44Service>();
            Loaded += (_, _) => LoadSettings();
        }

        private void LoadSettings()
        {
            Sim44Settings settings = Sim44Service.ReadSettings();
            txt_dataSource.Text = settings.DataSource ?? "";
            txt_table.Text = settings.Table ?? "";
            txt_bridge.Text = settings.BridgePath ?? "";
            txt_probe.Text = string.IsNullOrEmpty(_service.CurrentUser)
                ? LanguageService.Get("Sim44_Settings_ProbeHint")
                : string.Format(LanguageService.Get("Sim44_Settings_ProbeUserFormat"), _service.CurrentUser);
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            string dataSource = (txt_dataSource.Text ?? "").Trim();
            string table = (txt_table.Text ?? "").Trim();
            string bridge = (txt_bridge.Text ?? "").Trim();
            if (string.IsNullOrEmpty(dataSource))
            {
                MessageBox.Show(this, LanguageService.Get("Sim44_Settings_MsgDataSourceRequired"),
                    LanguageService.Get("Cap_Info"));
                txt_dataSource.Focus();
                return;
            }
            if (string.IsNullOrEmpty(table))
            {
                MessageBox.Show(this, LanguageService.Get("Sim44_Settings_MsgTableRequired"), LanguageService.Get("Cap_Info"));
                txt_table.Focus();
                return;
            }
            try
            {
                Sim44Service.SaveSettings(s =>
                {
                    s.DataSource = dataSource;
                    s.Table = table;
                    s.BridgePath = string.IsNullOrEmpty(bridge) ? null : bridge;
                });
                _logger.Info($"SIM44 连接设置已保存：数据源={dataSource}，表名={table}"
                             + (string.IsNullOrEmpty(bridge) ? "" : "，桥接=" + bridge));
                ToastService.Show(LanguageService.Get("Sim44_Settings_Saved"), ToastType.Info);
                Close();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存 SIM44 连接设置失败");
                MessageBox.Show(this, ex.Message, LanguageService.Get("Cap_SaveFailed"));
            }
        }

        /// <summary>探测：用当前登录的账号查数据字典，看看三张表在哪、能不能看见</summary>
        private void Btn_Probe_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_service.CurrentUser))
            {
                txt_probe.Text = LanguageService.Get("Sim44_Settings_ProbeHint");
                return;
            }
            btn_probe.IsEnabled = false;
            try
            {
                System.Collections.Generic.List<string> lines = _service.Probe();
                txt_probe.Text = lines.Count == 0
                    ? LanguageService.Get("Sim44_Settings_ProbeEmpty")
                    : string.Join(Environment.NewLine, lines);
                _logger.Info("SIM44 数据字典探测：" + txt_probe.Text.Replace(Environment.NewLine, "；"));
            }
            catch (Exception ex)
            {
                txt_probe.Text = ex.Message;
                _logger.Warn("SIM44 数据字典探测失败：" + ex.Message);
            }
            finally
            {
                btn_probe.IsEnabled = true;
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
