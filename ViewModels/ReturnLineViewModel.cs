using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Windows.Input;

namespace ORT一键报告.ViewModels
{
    public partial class ReturnLineViewModel(Service service) : ObservableObject
    {
        private readonly IService _service = service;

        // 辅助方法：获取指定月份（或当前月）的第一个工作日
        private static DateTime GetFirstWorkingDayOfMonth(DateTime? date = null)
        {
            var targetDate = date ?? DateTime.Today;
            var firstDay = new DateTime(targetDate.Year, targetDate.Month, 1);

            // 循环跳过周末（周六和周日）
            while (firstDay.DayOfWeek == DayOfWeek.Saturday || firstDay.DayOfWeek == DayOfWeek.Sunday)
            {
                firstDay = firstDay.AddDays(1);
            }

            return firstDay;
        }

        private DateTime _startDate = GetFirstWorkingDayOfMonth();
        public DateTime StartDate
        {
            get => _startDate;
            set
            {
                if (SetProperty(ref _startDate, value))
                {
                    if (_endDate.HasValue && _startDate > _endDate) EndDate = _startDate;
                    else { if (!_endDate.HasValue) EndDate = _startDate; }
                }
            }
        }

        private DateTime? _endDate;
        public DateTime? EndDate
        {
            get => _endDate;
            set
            {
                if (SetProperty(ref _endDate, value))
                {
                    if (_endDate.HasValue && _endDate < _startDate) StartDate = _endDate;
                }
            }
        }

        private string _startRTAH;
        public string StartRTAH { get => _startRTAH; set => SetProperty(ref _startRTAH, value); }

        private string _endRTAH;
        public string EndRTAH { get => _endRTAH; set => SetProperty(ref _endRTAH, value); }

        private string _lTPath;
        public string LTPath { get => _lTPath; set => SetProperty(ref _lTPath, value); }


        public List<ReturnLineSingleViewModel> returnLineSingleViewModels { get; set; } = [];



        private RelayCommand selectLTPathCommand;
        public ICommand SelectLTPathCommand => selectLTPathCommand ??= new RelayCommand(SelectLTPath);

        private void SelectLTPath()
        {
            LTPath = _service.OpenPathDialog("选择领退路径");
        }
    }
}
