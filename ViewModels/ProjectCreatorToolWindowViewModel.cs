using MarcusRunge.CleanArchitectureProjectGenerator.Commands;
using MarcusRunge.CleanArchitectureProjectGenerator.Constants;
using Microsoft.VisualStudio.Composition;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CreationPolicy = System.ComponentModel.Composition.CreationPolicy;
using ExportAttribute = System.ComponentModel.Composition.ExportAttribute;

namespace MarcusRunge.CleanArchitectureProjectGenerator.ViewModels
{

    [Export(typeof(ProjectCreatorToolWindowViewModel))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    internal class ProjectCreatorToolWindowViewModel : INotifyPropertyChanged
    {
        private ICommand? _buttonCommand;
        private bool _isBusy;
        private string? _projectName;
        public ICommand ButtonCommand => _buttonCommand ??= new RelayCommand<string>(ExecuteButtonCommand);
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }

        public string? ProjectName { get => _projectName; set => SetProperty(ref _projectName, value); }

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

        #region INotifyPropertyChanged Implementation

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        protected bool SetProperty<T>(ref T backingField, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(backingField, value))
                return false;
            backingField = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        #endregion INotifyPropertyChanged Implementation
    }
}