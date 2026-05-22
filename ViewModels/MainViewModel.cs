using CommunityToolkit.Mvvm.ComponentModel;
using System.IO;
using System.Windows.Input;

namespace ORT一键报告.ViewModels
{
    public class MainViewModel : ObservableObject
    {
        private readonly IService Service;
        public string ATEPath { get; set; }


        private string _reportPath;
        public string ReportPath
        {
            get => _reportPath;
            set
            {
                if (SetProperty(ref _reportPath, value))
                {
                    MainWindow.RootPath = Path.GetDirectoryName(value);
                    selectReportPathCommand.RaiseCanExecuteChanged();
                }
            }
        }

        private string _title = "ORT一键报告";
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        public MainViewModel(IService service)
        {
            Service = service;
        }

        private RelayCommand selectReportPathCommand;
        public ICommand SelectReportPathCommand => selectReportPathCommand ??= new RelayCommand(SelectReportPath);


        private void SelectReportPath()
        {
            ReportPath = Service.OpenPathDialog("选择报告概览");
            string _title = Path.GetFileName(Path.GetDirectoryName(ReportPath));
            Title = _title.Split(' ')[0] + " " + _title.Split('_')[1] + " ORT一键报告";
        }

        private bool CanRootReportPath()
        {
            return !string.IsNullOrEmpty(ReportPath);
        }
    }
}
