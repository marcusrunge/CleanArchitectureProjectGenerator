using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Commands
{
    internal sealed class AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null, Action<Exception>? onException = null) : ICommand
    {
        private readonly Func<bool>? _canExecute = canExecute;
        private readonly Func<Task> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Action<Exception>? _onException = onException;

        private bool _isExecuting;

        public event EventHandler? CanExecuteChanged;

        public bool IsExecuting => _isExecuting;

        public bool CanExecute(object? parameter) => !_isExecuting && (_canExecute?.Invoke() ?? true);

        public async void Execute(object? parameter) => await ExecuteAsync().ConfigureAwait(false);

        public async Task ExecuteAsync()
        {
            if (!CanExecute(null))
                return;

            try
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();

                await _execute().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (_onException is not null)
                {
                    _onException(ex);
                    return;
                }

                // Wenn kein Handler vorhanden ist, Fehler bewusst weiterwerfen,
                // damit sie beim Debuggen sichtbar bleiben.
                throw;
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    internal sealed class AsyncRelayCommand<T>(Func<T?, Task> execute, Func<T?, bool>? canExecute = null, Action<Exception>? onException = null) : ICommand
    {
        private readonly Func<T?, bool>? _canExecute = canExecute;
        private readonly Func<T?, Task> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Action<Exception>? _onException = onException;

        private bool _isExecuting;

        public event EventHandler? CanExecuteChanged;

        public bool IsExecuting => _isExecuting;

        public bool CanExecute(object? parameter)
        {
            if (_isExecuting) return false;

            if (_canExecute is null) return true;

            return parameter is T t ? _canExecute(t) : _canExecute(default);
        }

        public async void Execute(object? parameter) => await ExecuteAsync(parameter).ConfigureAwait(false);

        public async Task ExecuteAsync(object? parameter)
        {
            if (!CanExecute(parameter))
                return;

            try
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();

                if (parameter is T t)
                    await _execute(t).ConfigureAwait(false);
                else
                    await _execute(default).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (_onException is not null)
                {
                    _onException(ex);
                    return;
                }

                throw;
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    internal sealed class RelayCommands(Action execute, Func<bool>? canExecute = null) : ICommand
    {
        private readonly Func<bool>? _canExecute = canExecute;
        private readonly Action _execute = execute ?? throw new ArgumentNullException(nameof(execute));

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) => _execute();

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }

    internal sealed class RelayCommand<T>(Action<T?> execute, Func<T?, bool>? canExecute = null) : ICommand
    {
        private readonly Func<T?, bool>? _canExecute = canExecute;
        private readonly Action<T?> _execute = execute ?? throw new ArgumentNullException(nameof(execute));

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => _canExecute is null || (parameter is T t ? _canExecute(t) : _canExecute(default));

        public void Execute(object? parameter)
        {
            if (parameter is T t) _execute(t);
            else _execute(default);
        }

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}