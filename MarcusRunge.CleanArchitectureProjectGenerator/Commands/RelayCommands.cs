using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Commands
{
    /// <summary>
    /// Helper that ensures actions are executed on the captured UI SynchronizationContext.
    /// </summary>
    internal static class UiThread
    {
        /// <summary>
        /// Executes the given action on the UI thread if a UI context is available; otherwise executes inline.
        /// </summary>
        public static void Post(SynchronizationContext? uiContext, Action action)
        {
            if (action is null) return;

            // If we have no context (e.g., unit tests/console), run inline.
            if (uiContext is null)
            {
                action();
                return;
            }

            // If we are already on that context, run inline.
            if (ReferenceEquals(SynchronizationContext.Current, uiContext))
            {
                action();
                return;
            }

            // Otherwise marshal to UI context.
            uiContext.Post(_ => action(), null);
        }
    }

    /// <summary>
    /// An <see cref="ICommand"/> implementation that executes an asynchronous action and
    /// automatically disables itself while the action is running.
    /// </summary>
    /// <remarks>
    /// Robust version: raises CanExecuteChanged on the captured UI thread to avoid cross-thread exceptions in WPF.
    /// </remarks>
    internal sealed class AsyncRelayCommand(
        Func<Task> execute,
        Func<bool>? canExecute = null,
        Action<Exception>? onException = null) : ICommand
    {
        private readonly Func<bool>? _canExecute = canExecute;
        private readonly Func<Task> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Action<Exception>? _onException = onException;

        // Capture the UI context at construction time (should be created on UI thread in WPF).
        private readonly SynchronizationContext? _uiContext = SynchronizationContext.Current;

        // Indicates whether the command is currently running to avoid concurrent executions.
        private bool _isExecuting;

        /// <summary>
        /// Raised when something changes that may affect whether the command can execute.
        /// </summary>
        public event EventHandler? CanExecuteChanged;

        /// <summary>
        /// Gets whether the command is currently executing.
        /// </summary>
        public bool IsExecuting => _isExecuting;

        /// <summary>
        /// Determines whether the command can execute.
        /// </summary>
        public bool CanExecute(object? parameter)
        {
            if (_isExecuting)
                return false;

            return _canExecute?.Invoke() ?? true;
        }

        /// <summary>
        /// Executes the command via the <see cref="ICommand"/> interface (cannot be awaited).
        /// </summary>
        public async void Execute(object? parameter)
        {
            // Keep ConfigureAwait(false) if you want; UI notifications are marshaled explicitly.
            await ExecuteAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// Executes the command and returns a task that completes when the operation finishes.
        /// </summary>
        public async Task ExecuteAsync()
        {
            if (!CanExecute(null))
                return;

            try
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();

                // Do not capture UI context; robust UI marshaling is handled separately.
                await _execute().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (_onException is not null)
                {
                    // Marshal exception callback to UI thread as well (common: MessageBox/status updates).
                    UiThread.Post(_uiContext, () => _onException(ex));
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

        /// <summary>
        /// Triggers <see cref="CanExecuteChanged"/> so the UI re-evaluates <see cref="CanExecute"/>.
        /// </summary>
        public void RaiseCanExecuteChanged()
        {
            // Copy delegate to avoid race conditions if subscribers change concurrently.
            var handler = CanExecuteChanged;
            if (handler is null) return;

            UiThread.Post(_uiContext, () => handler.Invoke(this, EventArgs.Empty));
        }
    }

    /// <summary>
    /// An <see cref="ICommand"/> implementation that executes an asynchronous action with a typed parameter and
    /// automatically disables itself while running.
    /// </summary>
    /// <typeparam name="T">The expected parameter type.</typeparam>
    internal sealed class AsyncRelayCommand<T>(
        Func<T?, Task> execute,
        Func<T?, bool>? canExecute = null,
        Action<Exception>? onException = null) : ICommand
    {
        private readonly Func<T?, bool>? _canExecute = canExecute;
        private readonly Func<T?, Task> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Action<Exception>? _onException = onException;

        // Capture UI context at construction.
        private readonly SynchronizationContext? _uiContext = SynchronizationContext.Current;

        // Re-entrancy guard.
        private bool _isExecuting;

        public event EventHandler? CanExecuteChanged;

        public bool IsExecuting => _isExecuting;

        /// <summary>
        /// Determines whether the command can execute for the given parameter.
        /// </summary>
        public bool CanExecute(object? parameter)
        {
            if (_isExecuting) return false;
            if (_canExecute is null) return true;

            return parameter is T t ? _canExecute(t) : _canExecute(default);
        }

        /// <summary>
        /// Executes the command via <see cref="ICommand"/> (void signature, cannot be awaited).
        /// </summary>
        public async void Execute(object? parameter)
        {
            await ExecuteAsync(parameter).ConfigureAwait(false);
        }

        /// <summary>
        /// Executes the command with the provided parameter and returns a task that completes when done.
        /// </summary>
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
                    UiThread.Post(_uiContext, () => _onException(ex));
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

        public void RaiseCanExecuteChanged()
        {
            var handler = CanExecuteChanged;
            if (handler is null) return;

            UiThread.Post(_uiContext, () => handler.Invoke(this, EventArgs.Empty));
        }
    }

    /// <summary>
    /// A synchronous <see cref="ICommand"/> implementation that delegates to an <see cref="Action"/>.
    /// </summary>
    /// <remarks>
    /// Even sync commands can be triggered from non-UI threads in edge cases (timers, background services),
    /// so CanExecuteChanged is marshaled robustly as well.
    /// </remarks>
    internal sealed class RelayCommands(Action execute, Func<bool>? canExecute = null) : ICommand
    {
        private readonly Func<bool>? _canExecute = canExecute;
        private readonly Action _execute = execute ?? throw new ArgumentNullException(nameof(execute));

        private readonly SynchronizationContext? _uiContext = SynchronizationContext.Current;

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter)
        {
            return _canExecute?.Invoke() ?? true;
        }

        public void Execute(object? parameter)
        {
            _execute();
        }

        public void RaiseCanExecuteChanged()
        {
            var handler = CanExecuteChanged;
            if (handler is null) return;

            UiThread.Post(_uiContext, () => handler.Invoke(this, EventArgs.Empty));
        }
    }

    /// <summary>
    /// A synchronous <see cref="ICommand"/> implementation that delegates to an <see cref="Action{T}"/>.
    /// </summary>
    internal sealed class RelayCommand<T>(Action<T?> execute, Func<T?, bool>? canExecute = null) : ICommand
    {
        private readonly Func<T?, bool>? _canExecute = canExecute;
        private readonly Action<T?> _execute = execute ?? throw new ArgumentNullException(nameof(execute));

        private readonly SynchronizationContext? _uiContext = SynchronizationContext.Current;

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter)
        {
            if (_canExecute is null) return true;
            return parameter is T t ? _canExecute(t) : _canExecute(default);
        }

        public void Execute(object? parameter)
        {
            _execute(parameter is T t ? t : default);
        }

        public void RaiseCanExecuteChanged()
        {
            var handler = CanExecuteChanged;
            if (handler is null) return;

            UiThread.Post(_uiContext, () => handler.Invoke(this, EventArgs.Empty));
        }
    }
}