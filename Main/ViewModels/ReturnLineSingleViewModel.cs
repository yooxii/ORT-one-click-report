using CommunityToolkit.Mvvm.ComponentModel;

namespace ORT一键报告.ViewModels
{
    public partial class ReturnLineSingleViewModel : ObservableObject
    {
        private string _model = "XXXXXX-XXXX";
        public string Model
        {
            get => _model;
            set => SetProperty(ref _model, value);
        }

        private string _version;
        public string Version
        {
            get => _version;
            set => SetProperty(ref _version, value);
        }

        private string _dc;
        public string DC
        {
            get => _dc;
            set => SetProperty(ref _dc, value);
        }

        private int _count;
        public int Count
        {
            get => _count;
            set => SetProperty(ref _count, value);
        }

        private string _testitem;
        public string Testitem
        {
            get => _testitem;
            set => SetProperty(ref _testitem, value);
        }

        private string _worker;
        public string WorkerNo
        {
            get => _worker;
            set => SetProperty(ref _worker, value);
        }


        private string _date;
        public string Date
        {
            get => _date;
            set => SetProperty(ref _date, value);
        }

        private string _steps;
        public string Steps
        {
            get => _steps;
            set => SetProperty(ref _steps, value);
        }

        private string _returnBy;
        public string ReturnBy
        {
            get => _returnBy;
            set => SetProperty(ref _returnBy, value);
        }

        private string _phone;
        public string Phone
        {
            get => _phone;
            set => SetProperty(ref _phone, value);
        }

        private string _recive;
        public string Recive
        {
            get => _recive;
            set => SetProperty(ref _recive, value);
        }

        private string _lineNo;
        public string LineNo
        {
            get => _lineNo;
            set => SetProperty(ref _lineNo, value);
        }
    }
}
