using MarcusRunge.CleanArchitectureProjectGenerator.Commands;
using MarcusRunge.CleanArchitectureProjectGenerator.Common;
using MarcusRunge.CleanArchitectureProjectGenerator.Constants;
using MarcusRunge.CleanArchitectureProjectGenerator.Services;
using Microsoft.VisualStudio.Composition;
using System.ComponentModel.Composition;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using CreationPolicy = System.ComponentModel.Composition.CreationPolicy;
using ExportAttribute = System.ComponentModel.Composition.ExportAttribute;

namespace MarcusRunge.CleanArchitectureProjectGenerator.ViewModels
{
    [Export(typeof(ProjectCreatorToolWindowViewModel))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    [method: ImportingConstructor]
    internal class ProjectCreatorToolWindowViewModel(IGeneratorService generatorService) : BindableBase
    {
        private ICommand? _buttonCommand, _loadedCommand, _unloadedCommand;
        private bool _isBusy;
        private string? _projectName;

        public ICommand ButtonCommand => _buttonCommand ??= new RelayCommand<string>(ExecuteButtonCommand);
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }
        public ICommand LoadedCommand => _loadedCommand ??= new AsyncRelayCommand(ExecuteLoadedCommandAsync);

        public string? ProjectName { get => _projectName; set => SetProperty(ref _projectName, value); }

        public ICommand UnloadedCommand => _unloadedCommand ??= new RelayCommands(ExecuteUnloadedCommand);

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