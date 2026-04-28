using System;
using System.Windows.Input;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Commands
{

    internal sealed class RelayCommand(Action execute, Func<bool>? canExecute = null) : ICommand
    {
        private readonly Action _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Func<bool>? _canExecute = canExecute;

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) => _execute();

        public event EventHandler? CanExecuteChanged;

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    internal sealed class RelayCommand<T>(Action<T?> execute, Func<T?, bool>? canExecute = null) : ICommand
    {
        private readonly Action<T?> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Func<T?, bool>? _canExecute = canExecute;

        public bool CanExecute(object? parameter)
        {
            return _canExecute is null || (parameter is T t ? _canExecute(t) : _canExecute(default));
        }

        public void Execute(object? parameter)
        {
            if (parameter is T t) _execute(t);
            else _execute(default);
        }

        public event EventHandler? CanExecuteChanged;

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}
