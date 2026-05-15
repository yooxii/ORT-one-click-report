using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ORT一键报告.ViewModels
{
    public class UserViewModel : ObservableObject
    {
        [ObservableProperty]
        private Models.User _user;

        [ObservableProperty]
        private string _errorMessage;

        [ObservableProperty]
        private bool _isLoading;


        public UserViewModel()
        {
            _user = new Models.User();
        }

    }
}
