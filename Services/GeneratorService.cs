using MarcusRunge.CleanArchitectureProjectGenerator.Common;
using System;
using System.ComponentModel.Composition;
using System.Threading;
using System.Threading.Tasks;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Services
{
    internal interface IGeneratorService
    {
        string? NameSpace { get; set; }

        Task CreateAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);

        Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);
    }

    [Export(typeof(IGeneratorService))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    internal class GeneratorService : BindableBase, IGeneratorService
    {
        private string? _nameSpace;

        public string? NameSpace { get => _nameSpace; set => SetProperty(ref _nameSpace, value); }

        public Task CreateAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }

        public Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }
}