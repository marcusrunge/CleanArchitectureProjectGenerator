using MarcusRunge.CleanArchitectureProjectGenerator.Commands;
using MarcusRunge.CleanArchitectureProjectGenerator.Common;
using MarcusRunge.CleanArchitectureProjectGenerator.Constants;
using MarcusRunge.CleanArchitectureProjectGenerator.Contracts;
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
    internal class ProjectCreatorToolWindowViewModel(IGeneratorService generatorService) : BindableBase, IFrameworkElementLifecycleAware
    {
        private ICommand? _buttonCommand;       
        private string? _projectName;

        public ICommand ButtonCommand => _buttonCommand ??= new RelayCommand<string>(ExecuteButtonCommand);
       
        public string? ProjectName { get => _projectName; set => SetProperty(ref _projectName, value); }

        
        public async ValueTask OnLoadedAsync(CancellationToken cancellationToken = default)
        {
            await generatorService.InitializeAsync(ex => { }, cancellationToken);
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