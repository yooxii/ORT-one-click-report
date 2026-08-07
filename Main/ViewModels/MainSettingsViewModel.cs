using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ORT一键报告.ViewModels
{
    public partial class MainSettingsViewModel(IPathService service) : SettingsViewModel
    {
        private IPathService _service = service;

    }
}
