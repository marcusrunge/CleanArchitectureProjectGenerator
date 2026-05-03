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
using System.Xml.Linq;
using CreationPolicy = System.ComponentModel.Composition.CreationPolicy;
using ExportAttribute = System.ComponentModel.Composition.ExportAttribute;

namespace MarcusRunge.CleanArchitectureProjectGenerator.ViewModels
{
    [Export(typeof(ProjectCreatorToolWindowViewModel))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    [method: ImportingConstructor]
    internal class ProjectCreatorToolWindowViewModel(IGeneratorService generatorService) : BindableBase, IFrameworkElementLifecycleAware
    {
        private ICommand? _buttonCommand;
        private ObservableCollection<string> _dotNetVersions = [];
        private string? _projectName, _baseNamespace;

        private string? _selectedDotNetVersion;
        public string? BaseNamespace { get => _baseNamespace; set => SetProperty(ref _baseNamespace, value); }
        public ICommand ButtonCommand => _buttonCommand ??= new RelayCommand<string>(ExecuteButtonCommand);
        public ObservableCollection<string> DotNetVersions { get => _dotNetVersions; set => SetProperty(ref _dotNetVersions, value); }

        public string? ProjectName
        {
            get => _projectName;
            set
            {
                SetProperty(ref _projectName, value);
                BaseNamespace = $"{generatorService.Namespace}.{value}";
            }
        }

        public string? SelectedDotNetVersion { get => _selectedDotNetVersion; set => SetProperty(ref _selectedDotNetVersion, value); }

        public async ValueTask OnLoadedAsync(CancellationToken cancellationToken = default)
        {
            await generatorService.InitializeAsync(ex => { }, cancellationToken);
            BaseNamespace = generatorService.Namespace;
            var dotNetVersions = await generatorService.GetDotNetVersionsAsync(ex => { }, cancellationToken);
            foreach (var dotNetVersion in dotNetVersions)
            {
                DotNetVersions.Add(dotNetVersion);
            }
        }

        public ValueTask OnUnloadedAsync(CancellationToken cancellationToken = default)
        {
            return new ValueTask();
        }

        private void ExecuteButtonCommand(string? parameter)
        {
            if (string.IsNullOrEmpty(parameter))
                return;
            else if (parameter == ButtonCommandParameters.Cancel)
            {
            }
            else if (parameter == ButtonCommandParameters.Create)
            {
            }
            else
                return;
        }

        private async Task ExecuteLoadedCommandAsync()
        {
            await generatorService.InitializeAsync(ex => { }, CancellationToken.None);
        }

        private void ExecuteUnloadedCommand()
        {
        }
    }
}