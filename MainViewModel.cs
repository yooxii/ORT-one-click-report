using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ORT一键报告
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value))
            {
                return false;
            }

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        // 可绑定属性
        private string _rootReportPath = "请选择";
        public string RootReportPath
        {
            get => _rootReportPath;
            set => SetProperty(ref _rootReportPath, value);
        }

        private DateTime _t_datepicker_start = DateTime.Now;
        public DateTime TDatepicker_start
        {
            get => _t_datepicker_start;
            set => SetProperty(ref _t_datepicker_start, value);
        }
        private DateTime _t_datepicker_end;
        public DateTime TDatepicker_end
        {
            get => _t_datepicker_end;
            set => SetProperty(ref _t_datepicker_end, value);
        }

        private DateTime _b_datepicker_start = DateTime.Now;
        public DateTime BDatepicker_start
        {
            get => _b_datepicker_start;
            set => SetProperty(ref _b_datepicker_start, value);
        }
        private DateTime _b_datepicker_end;
        public DateTime BDatepicker_end
        {
            get => _b_datepicker_end;
            set => SetProperty(ref _b_datepicker_end, value);
        }
    }
}
