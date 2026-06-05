using CommunityToolkit.Mvvm.ComponentModel;

namespace ORT一键报告.ViewModels
{
    internal class ReportTemplateViewModel : ObservableObject
    {

        private string _title;
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }
    }
}
