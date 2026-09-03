using CommunityToolkit.Mvvm.ComponentModel;
using ORT一键报告.Services;

namespace ORT一键报告.ViewModels
{
    public class MainViewModel() : ObservableObject
    {
        private string _title = LanguageService.Get("App_MainTitle");
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        /// <summary>
        /// 订阅语言变更，实时刷新窗口标题
        /// </summary>
        public void SubscribeLanguageChange()
        {
            LanguageService.LanguageChanged += () => Title = LanguageService.Get("App_MainTitle");
        }
    }
}
