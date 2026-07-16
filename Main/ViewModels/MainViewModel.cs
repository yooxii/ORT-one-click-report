using CommunityToolkit.Mvvm.ComponentModel;

namespace ORT一键报告.ViewModels
{
    public class MainViewModel() : ObservableObject
    {
        private string _title = "ORT管理系统";
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }
    }
}
