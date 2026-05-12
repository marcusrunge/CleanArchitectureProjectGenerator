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
        string? RootNamespace { get; set; }

        /// <summary>
        /// Creates/generates the project artifacts.
        /// </summary>
        /// <param name="safeProjectname">The name of the assembly to create, typically entered by the user.</param>
        /// <param name="rootNamespace">The base root namespace to use for the generated code, often derived from <see cref="RootNamespace"/> and <paramref name="projectName"/>.</param>
        /// <param name="targetFramework">The target framework moniker (TFM) selected by the user (e.g., <c>net8.0</c>, <c>net48</c>).</param>
        /// <param name="exceptionCallback">
        /// A callback invoked when an exception occurs. This allows UI layers to display errors without crashing.
        /// </param>
        /// <param name="cancellationToken">A token used to cancel the operation.</param>
        Task CreateAsync(string safeProjectname, string rootNamespace, string targetFramework, Action<Exception> exceptionCallback, CancellationToken cancellationToken);

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
    /// <see cref="RootNamespace"/> are bound to UI).
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
        public string? RootNamespace
        {
            get => _namespace;
            set => SetProperty(ref _namespace, value);
        }

        /// <inheritdoc/>
        public async Task CreateAsync(string safeProjectname, string rootNamespace, string targetFramework, Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (string.IsNullOrWhiteSpace(safeProjectname))
                    throw new ArgumentException("Project name must not be empty.", nameof(safeProjectname));

                if (string.IsNullOrWhiteSpace(rootNamespace))
                    throw new ArgumentException("Base namespace must not be empty.", nameof(rootNamespace));

                if (string.IsNullOrWhiteSpace(targetFramework))
                    throw new ArgumentException("Target framework (dotNetVersion) must not be empty.", nameof(targetFramework));

                // Get solution directory (must exist and be open)
                var solutionDir = await GetSolutionDirectoryAsync(cancellationToken).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(solutionDir) || !Directory.Exists(solutionDir))
                    throw new InvalidOperationException("No solution is open or the solution directory could not be resolved.");

                // Create project folder under solution directory
                var projectDir = Path.Combine(solutionDir, safeProjectname);
                Directory.CreateDirectory(projectDir);

                // Create project with dotnet new
                await RunDotNetNewAsync(
                    templateShortName: "classlib",
                    projectName: safeProjectname,
                    outputDir: projectDir,
                    targetFramework: targetFramework,
                    cancellationToken: cancellationToken).ConfigureAwait(false);

                cancellationToken.ThrowIfCancellationRequested();

                // Find generated csproj
                var csprojPath = Directory.GetFiles(projectDir, "*.csproj", SearchOption.TopDirectoryOnly)
                                          .FirstOrDefault() ?? throw new FileNotFoundException("Project file (*.csproj) was not created by dotnet new.", projectDir);

                // Ensure RootNamespace matches the requested baseNamespace
                EnsureRootNamespaceInCsproj(csprojPath, rootNamespace);

                cancellationToken.ThrowIfCancellationRequested();

                // Add project to solution (must be on UI thread)
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                var dte = _serviceProvider?.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
                if (dte?.Solution is null)
                    throw new InvalidOperationException("DTE/Solution service is not available.");

                // This adds an existing project file to the solution
                dte.Solution.AddFromFile(csprojPath);

                // Optional: update bindable property so UI can reflect new namespace if desired
                RootNamespace = rootNamespace;
            }
            catch (OperationCanceledException)
            {
                // Don't treat cancellation as an error.
                throw;
            }
            catch (Exception ex)
            {
                exceptionCallback?.Invoke(ex);
                // Depending on your UI flow you can either swallow or rethrow.
                // Swallowing keeps UI responsive; rethrowing lets caller handle.
                // Here we swallow to match the "callback-driven" pattern.
            }
        }

        /// <inheritdoc/>
        public async Task<IReadOnlyList<string>> GetDotNetVersionsAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken) => await Task.Run(() =>
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
                    RootNamespace = null;
                    return;
                }

                // Try to get the hierarchy from the currently selected item in Solution Explorer.
                var hierarchy = TryGetHierarchyFromCurrentSelection(monitorSelection);

                // If no selection hierarchy is available, fall back to the startup project hierarchy.
                hierarchy ??= TryGetStartupProjectHierarchy(monitorSelection);

                // Convert the hierarchy root into a user-friendly project caption/name.
                RootNamespace = hierarchy != null ? GetRootProjectName(hierarchy) : null;
            }
            catch (Exception ex)
            {
                // On any failure, reset state so the UI doesn't display stale values.
                RootNamespace = null;

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

        private static void EnsureRootNamespaceInCsproj(string csprojPath, string rootNamespace)
        {
            // Edits csproj XML to ensure <RootNamespace> is set
            var doc = System.Xml.Linq.XDocument.Load(csprojPath);

            var project = doc.Root;
            if (project is null)
                throw new InvalidOperationException("Invalid csproj XML (missing root element).");

            // SDK-style csproj typically has no XML namespace; but handle both.
            var ns = project.Name.Namespace;

            // Find or create a PropertyGroup
            var propertyGroup = project.Elements(ns + "PropertyGroup").FirstOrDefault()
                                ?? new System.Xml.Linq.XElement(ns + "PropertyGroup");

            if (propertyGroup.Parent is null)
                project.AddFirst(propertyGroup);

            var rootNsElement = propertyGroup.Element(ns + "RootNamespace");
            if (rootNsElement is null)
            {
                propertyGroup.Add(new System.Xml.Linq.XElement(ns + "RootNamespace", rootNamespace));
            }
            else
            {
                rootNsElement.Value = rootNamespace;
            }

            doc.Save(csprojPath);
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

        private static async Task RunDotNetNewAsync(
            string templateShortName,
            string projectName,
            string outputDir,
            string targetFramework,
            CancellationToken cancellationToken)
        {
            // dotnet new classlib -n <name> -o <dir> -f <tfm> --no-restore
            var args =
                $"new {templateShortName} -n \"{projectName}\" -o \"{outputDir}\" -f \"{targetFramework}\" --no-restore";

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = args,
                WorkingDirectory = outputDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = new System.Diagnostics.Process { StartInfo = psi, EnableRaisingEvents = true };

            var stdOut = new System.Text.StringBuilder();
            var stdErr = new System.Text.StringBuilder();

            var tcs = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);

            process.OutputDataReceived += (_, e) => { if (e.Data != null) stdOut.AppendLine(e.Data); };
            process.ErrorDataReceived += (_, e) => { if (e.Data != null) stdErr.AppendLine(e.Data); };
            process.Exited += (_, __) => tcs.TrySetResult(process.ExitCode);

            if (!process.Start())
                throw new InvalidOperationException("Failed to start 'dotnet' process.");

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            using (cancellationToken.Register(() =>
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill();
                }
                catch
                {
                    // Ignore kill failures
                }
            }))
            {
                var exitCode = await tcs.Task.ConfigureAwait(false);

                if (exitCode != 0)
                {
                    throw new InvalidOperationException(
                        $"dotnet new failed with exit code {exitCode}.\n\nSTDOUT:\n{stdOut}\n\nSTDERR:\n{stdErr}");
                }
            }
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

        private async Task<string?> GetSolutionDirectoryAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            if (_serviceProvider.GetService(typeof(SVsSolution)) is not IVsSolution solution)
                return null;

            // If no solution is open, GetSolutionInfo often returns empty strings.
            ErrorHandler.ThrowOnFailure(solution.GetSolutionInfo(out var solutionDir, out _, out _));
            return solutionDir;
        }
    }
}