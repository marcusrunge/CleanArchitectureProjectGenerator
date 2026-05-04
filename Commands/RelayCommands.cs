using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Commands
{
    /// <summary>
    /// An <see cref="ICommand"/> implementation that executes an asynchronous action and
    /// automatically disables itself while the action is running.
    /// </summary>
    /// <remarks>
    /// Prevents re-entrancy (e.g., double-clicking a button) by tracking execution state and
    /// raising <see cref="ICommand.CanExecuteChanged"/> before/after execution.
    /// </remarks>
    internal sealed class AsyncRelayCommand(
        Func<Task> execute,
        Func<bool>? canExecute = null,
        Action<Exception>? onException = null) : ICommand
    {
        private readonly Func<bool>? _canExecute = canExecute;
        private readonly Func<Task> _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        private readonly Action<Exception>? _onException = onException;

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
            // If already running, disallow execution to prevent re-entrancy.
            if (_isExecuting)
                return false;

            // If no custom predicate is provided, allow execution by default.
            // Otherwise evaluate the custom predicate.
            return _canExecute?.Invoke() ?? true;
        }

        /// <summary>
        /// Executes the command via the <see cref="ICommand"/> interface (cannot be awaited).
        /// </summary>
        public async void Execute(object? parameter)
        {
            // ICommand requires a void-returning Execute; we bridge to ExecuteAsync.
            // Any exceptions are handled within ExecuteAsync or bubble if no handler is provided.
            await ExecuteAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// Executes the command and returns a task that completes when the operation finishes.
        /// </summary>
        public async Task ExecuteAsync()
        {
            // Fast exit if command is not allowed to run (e.g., already running or canExecute == false).
            if (!CanExecute(null))
                return;

            try
            {
                // Mark as running so CanExecute returns false and the UI can disable the command source.
                _isExecuting = true;

                // Notify the UI that CanExecute changed (commonly disables bound buttons).
                RaiseCanExecuteChanged();

                // Execute the provided async delegate.
                // ConfigureAwait(false) avoids capturing the UI context (safe here because we don't update UI directly).
                await _execute().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // If an exception handler was supplied, treat the exception as handled and do not rethrow.
                if (_onException is not null)
                {
                    _onException(ex);
                    return;
                }

                // Rethrow to preserve original stack trace.
                throw;
            }
            finally
            {
                // Always reset execution state, even if the operation fails.
                _isExecuting = false;

                // Notify UI again so it can re-enable the command source.
                RaiseCanExecuteChanged();
            }
        }

        /// <summary>
        /// Triggers <see cref="CanExecuteChanged"/> so the UI re-evaluates <see cref="CanExecute"/>.
        /// </summary>
        public void RaiseCanExecuteChanged()
        {
            // Null-conditional to safely invoke event only when there are subscribers.
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
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

        // Re-entrancy guard.
        private bool _isExecuting;

        public event EventHandler? CanExecuteChanged;

        public bool IsExecuting => _isExecuting;

        /// <summary>
        /// Determines whether the command can execute for the given parameter.
        /// </summary>
        public bool CanExecute(object? parameter)
        {
            // Prevent concurrent execution (e.g., repeated clicks).
            if (_isExecuting) return false;

            // If there is no predicate, we can execute.
            if (_canExecute is null) return true;

            // If the parameter is of type T, evaluate predicate with that value.
            // Otherwise pass default(T) (supports null/missing parameter scenarios).
            return parameter is T t ? _canExecute(t) : _canExecute(default);
        }

        /// <summary>
        /// Executes the command via <see cref="ICommand"/> (void signature, cannot be awaited).
        /// </summary>
        public async void Execute(object? parameter)
        {
            // ICommand.Execute cannot return Task, so we forward into an awaitable method.
            await ExecuteAsync(parameter).ConfigureAwait(false);
        }

        /// <summary>
        /// Executes the command with the provided parameter and returns a task that completes when done.
        /// </summary>
        public async Task ExecuteAsync(object? parameter)
        {
            // Respect CanExecute (including re-entrancy and custom predicate).
            if (!CanExecute(parameter))
                return;

            try
            {
                // Enter executing state and notify UI so it can disable the command source.
                _isExecuting = true;
                RaiseCanExecuteChanged();

                // If the parameter can be cast to T, pass it to the delegate.
                // Otherwise, pass default(T) to keep behavior consistent for missing/wrong parameters.
                if (parameter is T t)
                    await _execute(t).ConfigureAwait(false);
                else
                    await _execute(default).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Optional exception handling hook. If provided, swallow after callback.
                if (_onException is not null)
                {
                    _onException(ex);
                    return;
                }

                // Otherwise rethrow to let callers observe failures (e.g., global exception handler).
                throw;
            }
            finally
            {
                // Ensure we always leave executing state and notify UI to re-enable controls.
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>
    /// A synchronous <see cref="ICommand"/> implementation that delegates to an <see cref="Action"/>.
    /// </summary>
    internal sealed class RelayCommands(Action execute, Func<bool>? canExecute = null) : ICommand
    {
        private readonly Func<bool>? _canExecute = canExecute;
        private readonly Action _execute = execute ?? throw new ArgumentNullException(nameof(execute));

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter)
        {
            // If no predicate is provided, allow execution by default.
            // Otherwise evaluate the predicate.
            return _canExecute?.Invoke() ?? true;
        }

        public void Execute(object? parameter)
        {
            // Synchronous execution: invoke the provided action immediately.
            _execute();
        }

        public void RaiseCanExecuteChanged()
        {
            // Notify UI to re-check CanExecute (e.g., enable/disable button).
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>
    /// A synchronous <see cref="ICommand"/> implementation that delegates to an <see cref="Action{T}"/>.
    /// </summary>
    /// <typeparam name="T">The expected parameter type.</typeparam>
    internal sealed class RelayCommand<T>(Action<T?> execute, Func<T?, bool>? canExecute = null) : ICommand
    {
        private readonly Func<T?, bool>? _canExecute = canExecute;
        private readonly Action<T?> _execute = execute ?? throw new ArgumentNullException(nameof(execute));

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter)
        {
            // If there is no predicate, the command is always executable.
            if (_canExecute is null) return true;

            // Evaluate predicate using typed parameter when possible.
            // If parameter is missing or not castable to T, use default(T).
            return parameter is T t ? _canExecute(t) : _canExecute(default);
        }

        public void Execute(object? parameter)
        {
            // Convert object parameter to T if possible; otherwise fall back to default(T).
            // This makes the command robust against null or unexpected parameter types.
            _execute(parameter is T t ? t : default);
        }

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}