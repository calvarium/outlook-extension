using System;
using System.Windows.Input;

namespace outlook_extension.UI.Wpf.Infrastructure
{
    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool> _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            if (_canExecute == null)
            {
                return true;
            }

            if (parameter is T cast)
            {
                return _canExecute(cast);
            }

            return _canExecute(default);
        }

        public void Execute(object parameter)
        {
            if (parameter is T cast)
            {
                _execute(cast);
                return;
            }

            _execute(default);
        }

        public event EventHandler CanExecuteChanged;

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}
