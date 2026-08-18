using System;
using System.Windows.Input;

namespace ORT一键报告
{
    public class RelayCommand(Action execute, Func<bool> canExecute) : ICommand
    {
        private readonly Action _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Func<bool> _canExecute = canExecute;

        /// <summary>
        /// 挂接 CommandManager.RequerySuggested，使按钮状态随界面事件自动重新评估
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public RelayCommand(Action execute) : this(execute, null) { }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }

        public void Execute(object parameter)
        {
            _execute?.Invoke();
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }
}
