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
    /// Command handler
    /// </summary>
    internal sealed class ProjectCreatorToolWindowCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new("5f8e3682-65e2-4b21-b209-1eecd93d7bbf");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectCreatorToolWindowCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ProjectCreatorToolWindowCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ProjectCreatorToolWindowCommand? Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider => package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ProjectCreatorToolWindowCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            if (await package.GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
                Instance = new ProjectCreatorToolWindowCommand(package, commandService);
        }

        /// <summary>
        /// Shows the tool window when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>

        private void Execute(object sender, EventArgs e)
        {
            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

                var window = await package.ShowToolWindowAsync(
                    typeof(ProjectCreatorToolWindow), 0, true, package.DisposalToken);

                if (window?.Frame == null)
                    throw new NotSupportedException("Cannot create tool window");

                var vm = await ResolveViewModelFromMefAsync();

                if (window.Content is ProjectCreatorToolWindowControl control && window.Frame is IVsWindowFrame vsWindowFrame)
                {
                    ToolWindowCloser.SetFrame(control, vsWindowFrame);

                    control.DataContext = vm;

                    if (vm is IFrameworkElementLifecycleAware lifecycle)
                    {
                        var lifetimeCts = CancellationTokenSource.CreateLinkedTokenSource(package.DisposalToken);

                        try
                        {
                            await lifecycle.OnLoadedAsync(lifetimeCts.Token);
                        }
                        catch (OperationCanceledException)
                        {
                            lifetimeCts.Dispose();
                            return;
                        }
                        catch (Exception ex)
                        {
                            lifetimeCts.Dispose();
                            // TODO: log ex
                            throw;
                        }

                        void UnloadedHandler(object? s, EventArgs args)
                        {
                            control.Unloaded -= UnloadedHandler;
                            lifetimeCts.Cancel();

                            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                            {
                                try
                                {
                                    // If OnUnloadedAsync touches UI, ensure UI thread:
                                    // await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                                    await lifecycle.OnUnloadedAsync(CancellationToken.None);
                                }
                                catch (OperationCanceledException)
                                {
                                    // ignore if you decide to make it cancellable
                                }
                                catch (Exception ex)
                                {
                                    // TODO: log ex (don’t silently swallow)
                                }
                                finally
                                {
                                    lifetimeCts.Dispose();
                                }
                            });
                        }

                        control.Unloaded += UnloadedHandler;
                    }
                }
            });
        }

        private async Task<ProjectCreatorToolWindowViewModel> ResolveViewModelFromMefAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (await package.GetServiceAsync(typeof(SComponentModel)) is not IComponentModel componentModel)
                throw new InvalidOperationException("SComponentModel not available.");

            var vm = componentModel.GetService<ProjectCreatorToolWindowViewModel>() ?? throw new InvalidOperationException("ViewModel not found. Is it exported via [Export] and included in VSIX?");
            return vm;
        }

    }
}
