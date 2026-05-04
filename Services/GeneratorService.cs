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
    /// <summary>
    /// Defines the contract for a service that generates/scaffolds projects and provides
    /// environment-specific information required for generation.
    /// </summary>
    /// <remarks>
    /// This interface is intended for use in a Visual Studio extension context:
    /// <list type="bullet">
    /// <item><description><see cref="InitializeAsync"/> inspects the current VS selection to infer defaults (e.g., namespace).</description></item>
    /// <item><description><see cref="GetDotNetVersionsAsync"/> discovers available target frameworks based on installed SDKs/frameworks.</description></item>
    /// <item><description><see cref="CreateAsync"/> performs the actual generation/scaffolding (not implemented here yet).</description></item>
    /// </list>
    /// </remarks>
    internal interface IGeneratorService
    {
        /// <summary>
        /// Gets or sets the namespace (or root project name) inferred from the current context.
        /// </summary>
        /// <remarks>
        /// This is commonly bound to UI (ViewModel) input so it can be displayed/edited.
        /// </remarks>
        string? Namespace { get; set; }

        /// <summary>
        /// Creates/generates the project artifacts.
        /// </summary>
        /// <param name="exceptionCallback">
        /// A callback invoked when an exception occurs. This allows UI layers to display errors without crashing.
        /// </param>
        /// <param name="cancellationToken">A token used to cancel the operation.</param>
        Task CreateAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);

        /// <summary>
        /// Returns a list of available target framework monikers (TFMs) based on the machine's installed .NET SDKs
        /// and .NET Framework reference assemblies.
        /// </summary>
        /// <param name="exceptionCallback">
        /// A callback invoked when an exception occurs.
        /// </param>
        /// <param name="cancellationToken">A token used to cancel the operation.</param>
        /// <returns>A read-only list of TFM strings (e.g., <c>net8.0</c>, <c>net48</c>).</returns>
        Task<IReadOnlyList<string>> GetDotNetVersionsAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);

        /// <summary>
        /// Initializes the service state from the current Visual Studio context (e.g., current selection).
        /// </summary>
        /// <param name="exceptionCallback">
        /// A callback invoked when an exception occurs.
        /// </param>
        /// <param name="cancellationToken">A token used to cancel the operation.</param>
        Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken);
    }

    /// <summary>
    /// Default implementation of <see cref="IGeneratorService"/> composed via MEF.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Exported as non-shared so each consumer gets a fresh instance (useful when stateful properties like
    /// <see cref="Namespace"/> are bound to UI).
    /// </para>
    /// <para>
    /// Inherits from <see cref="BindableBase"/> to support UI binding through <see cref="System.ComponentModel.INotifyPropertyChanged"/>.
    /// </para>
    /// </remarks>
    [Export(typeof(IGeneratorService))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    [method: ImportingConstructor]
    internal class GeneratorService([Import(typeof(SVsServiceProvider))] IServiceProvider serviceProvider)
        : BindableBase, IGeneratorService
    {
        /// <summary>
        /// Visual Studio service provider (SVsServiceProvider) used to retrieve shell services.
        /// </summary>
        private readonly IServiceProvider _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));

        // Backing field for the bindable Namespace property.
        private string? _namespace;

        /// <inheritdoc/>
        public string? Namespace
        {
            get => _namespace;
            set => SetProperty(ref _namespace, value);
        }

        /// <inheritdoc/>
        public Task CreateAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
            => Task.CompletedTask;

        /// <inheritdoc/>
        public async Task<IReadOnlyList<string>> GetDotNetVersionsAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
            => await Task.Run(() =>
            {
                // Collect targets in a set to avoid duplicates (case-insensitive).
                var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Add TFMs from installed .NET SDK folders (net5.0+).
                AddDotNetSdkTargets(results);

                // Add TFMs from .NET Framework reference assemblies (e.g., net48).
                AddNetFrameworkTargets(results);

                // Return a stable, sorted, read-only list for UI binding and deterministic behavior.
                return results
                    .OrderBy(t => t)
                    .ToList()
                    .AsReadOnly();
            });

        /// <inheritdoc/>
        public async Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            try
            {
                // Respect cancellation as early as possible.
                cancellationToken.ThrowIfCancellationRequested();

                // Accessing VS shell selection services should happen on the UI thread.
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                // SVsShellMonitorSelection provides access to the current selection and startup project.
                if (_serviceProvider.GetService(typeof(SVsShellMonitorSelection)) is not IVsMonitorSelection monitorSelection)
                {
                    // If the service isn't available, clear state and exit gracefully.
                    Namespace = null;
                    return;
                }

                // Try to get the hierarchy from the currently selected item in Solution Explorer.
                var hierarchy = TryGetHierarchyFromCurrentSelection(monitorSelection);

                // If no selection hierarchy is available, fall back to the startup project hierarchy.
                hierarchy ??= TryGetStartupProjectHierarchy(monitorSelection);

                // Convert the hierarchy root into a user-friendly project caption/name.
                Namespace = hierarchy != null ? GetRootProjectName(hierarchy) : null;
            }
            catch (Exception ex)
            {
                // On any failure, reset state so the UI doesn't display stale values.
                Namespace = null;

                // Report error via callback so the caller/UI can decide how to present it.
                exceptionCallback?.Invoke(ex);
            }
        }

        /// <summary>
        /// Adds SDK-based TFMs (net5.0+) by probing the dotnet SDK installation directory.
        /// </summary>
        /// <param name="targets">The target collection to populate.</param>
        private static void AddDotNetSdkTargets(HashSet<string> targets)
        {
            // Typical SDK path: %ProgramFiles%\dotnet\sdk
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var sdkRoot = Path.Combine(programFiles, "dotnet", "sdk");

            // If dotnet isn't installed (or path differs), exit silently.
            if (!Directory.Exists(sdkRoot))
                return;

            foreach (var dir in Directory.GetDirectories(sdkRoot))
            {
                // SDK directories are versioned, e.g. "8.0.203" or "8.0.100-preview.1".
                var name = Path.GetFileName(dir);

                // Parse the leading version portion before any '-' suffix.
                // If parsing fails, ignore this directory.
                if (!Version.TryParse(name.Split('-')[0], out var version))
                    continue;

                // .NET 5+ uses "net{major}.0" TFMs.
                if (version.Major >= 5)
                {
                    targets.Add($"net{version.Major}.0");
                }
            }
        }

        /// <summary>
        /// Adds .NET Framework TFMs by probing reference assemblies installed with Visual Studio / targeting packs.
        /// </summary>
        /// <param name="targets">The target collection to populate.</param>
        private static void AddNetFrameworkTargets(HashSet<string> targets)
        {
            // Typical ref assemblies path:
            // %ProgramFiles(x86)%\Reference Assemblies\Microsoft\Framework\.NETFramework
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
                // Folder names are usually "v4.8", "v4.7.2", etc.
                var name = Path.GetFileName(dir);

                if (name.StartsWith("v"))
                {
                    // Convert "v4.7.2" -> "472" and build TFM "net472".
                    // Note: This matches common TFM formatting for .NET Framework.
                    var version = name.TrimStart('v').Replace(".", "");
                    targets.Add($"net{version}");
                }
            }
        }

        /// <summary>
        /// Gets a displayable name for the root of a Visual Studio hierarchy (typically the project name).
        /// </summary>
        /// <param name="hierarchy">The hierarchy to query.</param>
        /// <returns>The project caption or name if available; otherwise <c>null</c>.</returns>
        /// <remarks>
        /// Queries VS hierarchy properties in order of preference:
        /// <list type="number">
        /// <item><description>Caption (what users usually see in Solution Explorer)</description></item>
        /// <item><description>Name (fallback)</description></item>
        /// </list>
        /// Must be called on the UI thread.
        /// </remarks>
        private static string? GetRootProjectName(IVsHierarchy hierarchy)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Prefer VSHPROPID_Caption since it reflects the displayed project name.
            if (ErrorHandler.Succeeded(
                    hierarchy.GetProperty(
                        VSConstants.VSITEMID_ROOT,
                        (int)__VSHPROPID.VSHPROPID_Caption,
                        out var captionObj))
                && captionObj is string caption
                && !string.IsNullOrWhiteSpace(caption))
            {
                return caption;
            }

            // Fallback to VSHPROPID_Name if caption isn't available.
            if (ErrorHandler.Succeeded(
                    hierarchy.GetProperty(
                        VSConstants.VSITEMID_ROOT,
                        (int)__VSHPROPID.VSHPROPID_Name,
                        out var nameObj))
                && nameObj is string name
                && !string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            return null;
        }

        /// <summary>
        /// Attempts to retrieve an <see cref="IVsHierarchy"/> from the user's current selection in Visual Studio.
        /// </summary>
        /// <param name="monitorSelection">The VS selection service.</param>
        /// <returns>The selected hierarchy, or <c>null</c> if no hierarchy is selected.</returns>
        /// <remarks>
        /// Uses COM interop:
        /// <list type="bullet">
        /// <item><description><see cref="IVsMonitorSelection.GetCurrentSelection"/> returns COM pointers that must be released.</description></item>
        /// <item><description>We convert the hierarchy IUnknown pointer to a managed object via <see cref="Marshal.GetObjectForIUnknown"/>.</description></item>
        /// </list>
        /// Must be called on the UI thread.
        /// </remarks>
        private static IVsHierarchy? TryGetHierarchyFromCurrentSelection(IVsMonitorSelection monitorSelection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IntPtr hierPtr = IntPtr.Zero;
            IntPtr selContainerPtr = IntPtr.Zero;

            try
            {
                // Request current selection info from VS.
                // hierPtr and selContainerPtr are COM pointers and must be released in finally.
                int hr = monitorSelection.GetCurrentSelection(
                    out hierPtr,
                    out uint itemid,
                    out IVsMultiItemSelect? multiSelect,
                    out selContainerPtr);

                // If the call fails or there is no hierarchy pointer, there is no valid selection hierarchy.
                if (!ErrorHandler.Succeeded(hr) || hierPtr == IntPtr.Zero)
                    return null;

                // Convert COM pointer to managed IVsHierarchy object.
                return Marshal.GetObjectForIUnknown(hierPtr) as IVsHierarchy;
            }
            finally
            {
                // Release COM pointers to prevent leaks.
                if (hierPtr != IntPtr.Zero) Marshal.Release(hierPtr);
                if (selContainerPtr != IntPtr.Zero) Marshal.Release(selContainerPtr);
            }
        }

        /// <summary>
        /// Attempts to retrieve the startup project hierarchy when there is no suitable current selection.
        /// </summary>
        /// <param name="monitorSelection">The VS selection service.</param>
        /// <returns>The startup project hierarchy, or <c>null</c> if not available.</returns>
        /// <remarks>
        /// The returned value can be:
        /// <list type="bullet">
        /// <item><description>An <see cref="IVsHierarchy"/> directly</description></item>
        /// <item><description>An <see cref="IntPtr"/> to a COM object</description></item>
        /// </list>
        /// Must be called on the UI thread.
        /// </remarks>
        private static IVsHierarchy? TryGetStartupProjectHierarchy(IVsMonitorSelection monitorSelection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // SEID_StartupProject provides the startup project (if set).
            int hr = monitorSelection.GetCurrentElementValue(
                (uint)VSConstants.VSSELELEMID.SEID_StartupProject,
                out object value);

            if (!ErrorHandler.Succeeded(hr) || value is null)
                return null;

            // Sometimes VS already returns a managed IVsHierarchy.
            if (value is IVsHierarchy hier)
                return hier;

            // In other cases, it returns a COM pointer to the hierarchy.
            if (value is IntPtr ptr && ptr != IntPtr.Zero)
                return Marshal.GetObjectForIUnknown(ptr) as IVsHierarchy;

            return null;
        }
    }
}