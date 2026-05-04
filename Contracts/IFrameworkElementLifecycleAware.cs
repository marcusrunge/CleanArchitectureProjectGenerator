using System.Threading;
using System.Threading.Tasks;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Contracts
{
    /// <summary>
    /// Defines an opt-in lifecycle contract for view models (or other components) that need to react
    /// to a framework element being loaded or unloaded.
    /// </summary>
    /// <remarks>
    /// This interface is typically used in UI scenarios (e.g., WPF tool windows) where a view/control
    /// is created and later removed from the visual tree. Implementers can:
    /// <list type="bullet">
    /// <item><description>Start async initialization work when the view is loaded (e.g., load data, subscribe to events).</description></item>
    /// <item><description>Perform async cleanup when the view is unloaded (e.g., unsubscribe, stop timers, flush state).</description></item>
    /// </list>
    /// <para>
    /// <see cref="ValueTask"/> is used to reduce allocations for implementations that often complete synchronously
    /// (e.g., when no actual async work is required).
    /// </para>
    /// </remarks>
    internal interface IFrameworkElementLifecycleAware
    {
        /// <summary>
        /// Called when the associated framework element is loaded and ready for interaction.
        /// </summary>
        /// <param name="cancellationToken">
        /// A token that can be used to cancel initialization work, for example when the host is shutting down
        /// or the view is closed before loading completes.
        /// </param>
        /// <returns>
        /// A <see cref="ValueTask"/> that completes when the load/initialization logic has finished.
        /// </returns>
        /// <remarks>
        /// Implementations should be resilient to cancellation and should avoid blocking the UI thread.
        /// If the method interacts with UI state, ensure it runs on the UI thread according to the host framework rules.
        /// </remarks>
        ValueTask OnLoadedAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Called when the associated framework element is unloaded and should release resources.
        /// </summary>
        /// <param name="cancellationToken">
        /// A token that can be used to cancel cleanup work (optional). Many cleanup routines are best-effort and quick,
        /// but cancellation can be helpful if teardown needs to stop ongoing operations.
        /// </param>
        /// <returns>
        /// A <see cref="ValueTask"/> that completes when the unload/cleanup logic has finished.
        /// </returns>
        /// <remarks>
        /// Use this method to:
        /// <list type="bullet">
        /// <item><description>Unsubscribe from events to prevent memory leaks.</description></item>
        /// <item><description>Stop background operations started during <see cref="OnLoadedAsync"/>.</description></item>
        /// <item><description>Dispose timers/CTS/IDisposable resources if owned by the implementer.</description></item>
        /// </list>
        /// </remarks>
        ValueTask OnUnloadedAsync(CancellationToken cancellationToken = default);
    }
}