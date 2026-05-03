using MarcusRunge.CleanArchitectureProjectGenerator.Common;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Services
{
    internal interface IGeneratorService
    {
        string? Namespace { get; set; }

        Task CreateAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);

        Task<IReadOnlyList<string>> GetDotNetVersionsAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);

        Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);
    }

    [Export(typeof(IGeneratorService))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    [method: ImportingConstructor]
    internal class GeneratorService([Import(typeof(SVsServiceProvider))] IServiceProvider serviceProvider) : BindableBase, IGeneratorService
    {
        private readonly IServiceProvider _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));

        private string? _namespace;

        public string? Namespace { get => _namespace; set => SetProperty(ref _namespace, value); }

        public Task CreateAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
            => Task.CompletedTask;

        public async Task<IReadOnlyList<string>> GetDotNetVersionsAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken) => await Task.Run(() =>
        {
            var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            AddDotNetSdkTargets(results);
            AddNetFrameworkTargets(results);

            return results
                .OrderBy(t => t)
                .ToList()
                .AsReadOnly();
        });

        public async Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                if (_serviceProvider.GetService(typeof(SVsShellMonitorSelection)) is not IVsMonitorSelection monitorSelection)
                {
                    Namespace = null;
                    return;
                }
                var hierarchy = TryGetHierarchyFromCurrentSelection(monitorSelection);
                hierarchy ??= TryGetStartupProjectHierarchy(monitorSelection);
                Namespace = hierarchy != null ? GetRootProjectName(hierarchy) : null;
            }
            catch (Exception ex)
            {
                Namespace = null;
                exceptionCallback?.Invoke(ex);
            }
        }

        private static void AddDotNetSdkTargets(HashSet<string> targets)
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var sdkRoot = Path.Combine(programFiles, "dotnet", "sdk");

            if (!Directory.Exists(sdkRoot))
                return;

            foreach (var dir in Directory.GetDirectories(sdkRoot))
            {
                var name = Path.GetFileName(dir);

                if (!Version.TryParse(name.Split('-')[0], out var version))
                    continue;

                if (version.Major >= 5)
                {
                    targets.Add($"net{version.Major}.0");
                }
            }
        }

        private static void AddNetFrameworkTargets(HashSet<string> targets)
        {
            var referenceRoot =
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Reference Assemblies",
                    "Microsoft",
                    "Framework",
                    ".NETFramework"
                );

            if (!Directory.Exists(referenceRoot))
                return;

            foreach (var dir in Directory.GetDirectories(referenceRoot))
            {
                var name = Path.GetFileName(dir);

                if (name.StartsWith("v"))
                {
                    var version = name.TrimStart('v').Replace(".", "");
                    targets.Add($"net{version}");
                }
            }
        }

        private static string? GetRootProjectName(IVsHierarchy hierarchy)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (ErrorHandler.Succeeded(hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_Caption, out var captionObj))
                && captionObj is string caption
                && !string.IsNullOrWhiteSpace(caption))
            {
                return caption;
            }
            if (ErrorHandler.Succeeded(hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_Name, out var nameObj))
                && nameObj is string name
                && !string.IsNullOrWhiteSpace(name))
            {
                return name;
            }
            return null;
        }

        private static IVsHierarchy? TryGetHierarchyFromCurrentSelection(IVsMonitorSelection monitorSelection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IntPtr hierPtr = IntPtr.Zero;
            IntPtr selContainerPtr = IntPtr.Zero;

            try
            {
                int hr = monitorSelection.GetCurrentSelection(out hierPtr, out uint itemid, out IVsMultiItemSelect? multiSelect, out selContainerPtr);
                if (!ErrorHandler.Succeeded(hr) || hierPtr == IntPtr.Zero)
                    return null;

                return Marshal.GetObjectForIUnknown(hierPtr) as IVsHierarchy;
            }
            finally
            {
                if (hierPtr != IntPtr.Zero) Marshal.Release(hierPtr);
                if (selContainerPtr != IntPtr.Zero) Marshal.Release(selContainerPtr);
            }
        }

        private static IVsHierarchy? TryGetStartupProjectHierarchy(IVsMonitorSelection monitorSelection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int hr = monitorSelection.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_StartupProject, out object value);
            if (!ErrorHandler.Succeeded(hr) || value is null)
                return null;
            if (value is IVsHierarchy hier)
                return hier;
            if (value is IntPtr ptr && ptr != IntPtr.Zero)
                return Marshal.GetObjectForIUnknown(ptr) as IVsHierarchy;
            return null;
        }
    }
}