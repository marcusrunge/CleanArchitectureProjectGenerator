using System.Threading;
using System.Threading.Tasks;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Contracts
{
    internal interface IFrameworkElementLifecycleAware
    {
        ValueTask OnLoadedAsync(CancellationToken cancellationToken = default);

        ValueTask OnUnloadedAsync(CancellationToken cancellationToken = default);
    }
}