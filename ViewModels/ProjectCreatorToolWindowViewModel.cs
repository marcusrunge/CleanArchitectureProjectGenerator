using MarcusRunge.CleanArchitectureProjectGenerator.Commands;
using MarcusRunge.CleanArchitectureProjectGenerator.Common;
using MarcusRunge.CleanArchitectureProjectGenerator.Constants;
using MarcusRunge.CleanArchitectureProjectGenerator.Contracts;
using MarcusRunge.CleanArchitectureProjectGenerator.Services;
using Microsoft.VisualStudio.Composition;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using CreationPolicy = System.ComponentModel.Composition.CreationPolicy;
using ExportAttribute = System.ComponentModel.Composition.ExportAttribute;

namespace MarcusRunge.CleanArchitectureProjectGenerator.ViewModels
{
    /// <summary>
    /// ViewModel for the Project Creator tool window.
    /// </summary>
    /// <remarks>
    /// Composed via MEF and created as <see cref="CreationPolicy.NonShared"/> so each tool window instance
    /// receives its own ViewModel instance with isolated state.
    /// <para>
    /// Implements <see cref="IFrameworkElementLifecycleAware"/> to support async initialization and teardown
    /// when the view is loaded/unloaded.
    /// </para>
    /// </remarks>
    [Export(typeof(ProjectCreatorToolWindowViewModel))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    [method: ImportingConstructor]
    internal class ProjectCreatorToolWindowViewModel(IGeneratorService generatorService)
        : BindableBase, IFrameworkElementLifecycleAware
    {
        // Backing field for lazy command initialization (avoid allocating command until first access).
        private ICommand? _buttonCommand;

        // Backing collection for UI-bound framework targets (ObservableCollection notifies UI on changes).
        private ObservableCollection<string> _dotNetVersions = [];

        // Bound flag used by the view (via attached property) to request closing the tool window.
        private bool _isOnCloseRequested;

        // Backing fields for bindable properties used by the UI.
        private string? _projectName, _baseNamespace;

        private string? _selectedDotNetVersion;

        /// <summary>
        /// Gets or sets the base namespace used for generated code (e.g., "Company.Product.Module").
        /// </summary>
        public string? BaseNamespace
        {
            get => _baseNamespace;
            set => SetProperty(ref _baseNamespace, value);
        }

        /// <summary>
        /// Gets the command invoked by the UI buttons (Cancel/Create), using a string command parameter.
        /// </summary>
        /// <remarks>
        /// The command parameter is expected to match values from <see cref="ButtonCommandParameters"/>.
        /// Lazy initialization ensures the command is created only when needed.
        /// </remarks>
        public ICommand ButtonCommand =>
            _buttonCommand ??= new RelayCommand<string>(ExecuteButtonCommand);

        /// <summary>
        /// Gets or sets the list of available target frameworks (TFMs) displayed in the UI.
        /// </summary>
        /// <remarks>
        /// <see cref="ObservableCollection{T}"/> is used so adding items triggers UI updates automatically.
        /// </remarks>
        public ObservableCollection<string> DotNetVersions
        {
            get => _dotNetVersions;
            set => SetProperty(ref _dotNetVersions, value);
        }

        /// <summary>
        /// Gets or sets a flag indicating that the tool window should be closed.
        /// </summary>
        /// <remarks>
        /// This is typically bound to an attached property (e.g., <c>ToolWindowCloser.CloseRequested</c>)
        /// so the ViewModel can request closing without referencing VS shell APIs.
        /// </remarks>
        public bool IsOnCloseRequested
        {
            get => _isOnCloseRequested;
            set => SetProperty(ref _isOnCloseRequested, value);
        }

        /// <summary>
        /// Gets or sets the project name entered by the user.
        /// </summary>
        /// <remarks>
        /// When the project name changes, <see cref="BaseNamespace"/> is updated by combining the
        /// inferred root namespace from <paramref name="generatorService"/> with the project name.
        /// </remarks>
        public string? ProjectName
        {
            get => _projectName;
            set
            {
                // Update backing field and raise PropertyChanged if the value actually changed.
                SetProperty(ref _projectName, value);

                // Derive BaseNamespace from the generator service's inferred namespace plus the project name.
                // Note: This will produce strings like "RootNamespace.ProjectName".
                // If generatorService.Namespace is null, this becomes " .{value}" (string interpolation yields " .x"?),
                // but we keep logic unchanged as requested.
                BaseNamespace = $"{generatorService.Namespace}.{value}";
            }
        }

        /// <summary>
        /// Gets or sets the currently selected target framework (TFM) in the UI.
        /// </summary>
        public string? SelectedDotNetVersion
        {
            get => _selectedDotNetVersion;
            set => SetProperty(ref _selectedDotNetVersion, value);
        }

        /// <summary>
        /// Lifecycle callback invoked when the view/control is loaded.
        /// </summary>
        /// <param name="cancellationToken">
        /// Token propagated from the host to cancel initialization (e.g., tool window closed or package disposed).
        /// </param>
        /// <remarks>
        /// Initializes the generator service (to infer default namespace) and populates the list of available
        /// target frameworks for user selection.
        /// </remarks>
        public async ValueTask OnLoadedAsync(CancellationToken cancellationToken = default)
        {
            // Initialize service state based on current VS context (selection/startup project).
            // Exception callback is currently a no-op; callers may later pass logging/UI error handling.
            await generatorService.InitializeAsync(ex => { }, cancellationToken);

            // Use the inferred namespace as the initial base namespace shown to the user.
            BaseNamespace = generatorService.Namespace;

            // Query available TFMs (e.g., net8.0, net48) for a dropdown/list in the UI.
            var dotNetVersions = await generatorService.GetDotNetVersionsAsync(ex => { }, cancellationToken);

            // Populate observable collection to update UI incrementally.
            foreach (var dotNetVersion in dotNetVersions)
            {
                DotNetVersions.Add(dotNetVersion);
            }

            // Note: You might consider selecting a default (e.g., latest) here,
            // but logic is intentionally unchanged.
        }

        /// <summary>
        /// Lifecycle callback invoked when the view/control is unloaded.
        /// </summary>
        /// <param name="cancellationToken">Token for optional cancellation of teardown logic.</param>
        /// <remarks>
        /// Currently does nothing and completes synchronously. Kept as a <see cref="ValueTask"/> to match the contract.
        /// Implement cleanup here if the ViewModel later acquires resources (timers, subscriptions, CTS, etc.).
        /// </remarks>
        public ValueTask OnUnloadedAsync(CancellationToken cancellationToken = default)
        {
            // No teardown required at the moment.
            return new ValueTask();
        }

        /// <summary>
        /// Executes the button command using a string parameter to decide which action to perform.
        /// </summary>
        /// <param name="parameter">
        /// A string command parameter expected to match <see cref="ButtonCommandParameters.Cancel"/> or
        /// <see cref="ButtonCommandParameters.Create"/>.
        /// </param>
        /// <remarks>
        /// This pattern allows multiple buttons to share one command while differentiating actions via CommandParameter.
        /// </remarks>
        private void ExecuteButtonCommand(string? parameter)
        {
            // Ignore missing parameters to keep command robust against misconfigured bindings.
            if (string.IsNullOrEmpty(parameter))
                return;

            // Route by well-known parameter constants instead of "magic strings".
            if (parameter == ButtonCommandParameters.Cancel)
            {
                // Trigger close via bound flag (typically handled by an attached behavior in the view).
                IsOnCloseRequested = true;
            }
            else if (parameter == ButtonCommandParameters.Create)
            {
                // Placeholder for creation logic.
                // Typical next steps might include validating inputs, calling generatorService.CreateAsync,
                // and then requesting close if successful.
            }
            else
            {
                // Unknown command parameter: ignore (defensive default).
                return;
            }
        }
    }
}