using MarcusRunge.CleanArchitectureProjectGenerator.Contracts;
using MarcusRunge.CleanArchitectureProjectGenerator.Helpers;
using MarcusRunge.CleanArchitectureProjectGenerator.ViewModels;
using MarcusRunge.CleanArchitectureProjectGenerator.Views;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Threading;
using System.Threading.Tasks;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Commands
{
    /// <summary>
    /// Registers and handles the command that opens the Project Creator tool window.
    /// </summary>
    /// <remarks>
    /// This command is registered via Visual Studio's command system (VSCT). When invoked,
    /// it shows a tool window and initializes its ViewModel via MEF (SComponentModel).
    /// </remarks>
    internal sealed class ProjectCreatorToolWindowCommand
    {
        /// <summary>
        /// The numeric command ID as defined in the VSCT file.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// The command set GUID (menu group) as defined in the VSCT file.
        /// </summary>
        public static readonly Guid CommandSet = new("5f8e3682-65e2-4b21-b209-1eecd93d7bbf");

        /// <summary>
        /// The owning Visual Studio package. Used as service provider and for tool window operations.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectCreatorToolWindowCommand"/> class and
        /// registers the menu command with Visual Studio.
        /// </summary>
        /// <param name="package">The owning package (must not be <c>null</c>).</param>
        /// <param name="commandService">The menu command service (must not be <c>null</c>).</param>
        /// <remarks>
        /// The constructor wires the command ID to the <see cref="Execute"/> handler.
        /// Visual Studio requires this registration to happen on the UI thread.
        /// </remarks>
        private ProjectCreatorToolWindowCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            // Store the package for later service lookups and tool window operations.
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // Create the VS command identifier based on the VSCT-defined command set and command ID.
            var menuCommandID = new CommandID(CommandSet, CommandId);

            // Bind the command invocation to our Execute method.
            var menuItem = new MenuCommand(this.Execute, menuCommandID);

            // Register the command with Visual Studio so it appears in the menu/command system.
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the singleton instance of the command handler.
        /// </summary>
        public static ProjectCreatorToolWindowCommand? Instance { get; private set; }

        /// <summary>
        /// Initializes (creates and registers) the singleton instance of the command.
        /// </summary>
        /// <param name="package">The owning package (must not be <c>null</c>).</param>
        /// <returns>A task that completes when initialization has finished.</returns>
        /// <remarks>
        /// This method must switch to the main thread because command registration interacts with the VS UI.
        /// </remarks>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Command registration must occur on the UI thread in Visual Studio.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            // Retrieve the command service and create the singleton instance if available.
            if (await package.GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
                Instance = new ProjectCreatorToolWindowCommand(package, commandService);
        }

        /// <summary>
        /// Command handler invoked when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender (provided by the VS command system).</param>
        /// <param name="e">The event arguments.</param>
        /// <remarks>
        /// Uses a JoinableTask to run async code safely from a synchronous event handler.
        /// </remarks>
        private void Execute(object sender, EventArgs e)
        {
            // The VS command callback is synchronous (void). We start an async workflow using RunAsync
            // to avoid blocking the UI thread while still keeping execution joined to VS threading rules.
            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                // Most VS shell operations (tool windows, UI objects) require the main thread.
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

                // Show (or create) the tool window instance.
                // - 0: tool window instance ID
                // - true: create if it does not exist
                // - DisposalToken: respects package shutdown/cancellation
                var window = await package.ShowToolWindowAsync(
                    typeof(ProjectCreatorToolWindow), 0, true, package.DisposalToken);

                // The shell returns a ToolWindowPane. If Frame is null, the window could not be created.
                if (window?.Frame == null)
                    throw new NotSupportedException("Cannot create tool window");

                // Resolve the ViewModel via MEF from Visual Studio's component model.
                // This keeps composition/DI consistent with VS extension patterns.
                var vm = await ResolveViewModelFromMefAsync();

                // Ensure the ToolWindow content is our expected control type, and the frame is a VS window frame.
                if (window.Content is ProjectCreatorToolWindowControl control && window.Frame is IVsWindowFrame vsWindowFrame)
                {
                    // Provide the IVsWindowFrame to helper logic so the control can close the window if needed.
                    ToolWindowCloser.SetFrame(control, vsWindowFrame);

                    // Attach the ViewModel to the view.
                    control.DataContext = vm;

                    // If the ViewModel participates in lifecycle events, wire up Loaded/Unloaded semantics.
                    if (vm is IFrameworkElementLifecycleAware lifecycle)
                    {
                        // Create a CTS linked to package disposal so:
                        // - closing VS / unloading package cancels ongoing work
                        // - we can also cancel when the control unloads
                        var lifetimeCts = CancellationTokenSource.CreateLinkedTokenSource(package.DisposalToken);

                        try
                        {
                            // Perform "Loaded" initialization (e.g., start background tasks, load configuration).
                            // Using the linked token means: if the tool window is closed or VS shuts down,
                            // the ViewModel can stop promptly.
                            await lifecycle.OnLoadedAsync(lifetimeCts.Token);
                        }
                        catch (OperationCanceledException)
                        {
                            // If initialization was canceled (window closed or package disposed),
                            // dispose the CTS and exit silently.
                            lifetimeCts.Dispose();
                            return;
                        }
                        catch (Exception ex)
                        {
                            // Always dispose resources on error to prevent leaks.
                            lifetimeCts.Dispose();

                            // TODO: Log the exception to VS ActivityLog or your own logger.
                            // Keeping the throw preserves the failure for diagnosis.
                            throw;
                        }

                        // Local handler for the control's Unloaded event.
                        // This is where we end the lifecycle and allow cleanup.
                        void UnloadedHandler(object? s, EventArgs args)
                        {
                            // Unsubscribe immediately to avoid multiple invocations and memory leaks.
                            control.Unloaded -= UnloadedHandler;

                            // Signal cancellation to any ongoing ViewModel work started in OnLoadedAsync.
                            lifetimeCts.Cancel();

                            // Run cleanup asynchronously without blocking the UI thread.
                            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                            {
                                try
                                {
                                    // If OnUnloadedAsync needs UI access, switch explicitly:
                                    // await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                                    // Perform ViewModel cleanup (stop timers, detach events, persist state, etc.).
                                    // Note: This call is currently not cancellable (CancellationToken.None).
                                    await lifecycle.OnUnloadedAsync(CancellationToken.None);
                                }
                                catch (OperationCanceledException)
                                {
                                    // If you later make OnUnloadedAsync cancellable, you can ignore expected cancellation here.
                                }
                                catch (Exception ex)
                                {
                                    // TODO: Log the exception (avoid silently swallowing unexpected failures).
                                }
                                finally
                                {
                                    // Ensure CTS is disposed regardless of outcome.
                                    lifetimeCts.Dispose();
                                }
                            });
                        }

                        // Subscribe to Unloaded to know when the tool window content is removed from the visual tree.
                        control.Unloaded += UnloadedHandler;
                    }
                }
            });
        }

        /// <summary>
        /// Resolves the tool window ViewModel from Visual Studio's MEF container (SComponentModel).
        /// </summary>
        /// <returns>The composed <see cref="ProjectCreatorToolWindowViewModel"/> instance.</returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown when SComponentModel is unavailable or the ViewModel is not exported.
        /// </exception>
        /// <remarks>
        /// Must run on the UI thread because Visual Studio service retrieval and MEF composition
        /// may require the main thread depending on the service implementation.
        /// </remarks>
        private async Task<ProjectCreatorToolWindowViewModel> ResolveViewModelFromMefAsync()
        {
            // Accessing certain VS services requires the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // SComponentModel is the entry point to Visual Studio's MEF composition container.
            if (await package.GetServiceAsync(typeof(SComponentModel)) is not IComponentModel componentModel)
                throw new InvalidOperationException("SComponentModel not available.");

            // Resolve the ViewModel from MEF.
            // If null: it's likely not exported ([Export]) or not included in the VSIX composition.
            var vm = componentModel.GetService<ProjectCreatorToolWindowViewModel>()
                     ?? throw new InvalidOperationException(
                         "ViewModel not found. Is it exported via [Export] and included in VSIX?");

            return vm;
        }
    }
}