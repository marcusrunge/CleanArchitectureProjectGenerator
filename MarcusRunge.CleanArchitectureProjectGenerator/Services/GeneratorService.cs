using MarcusRunge.CleanArchitectureProjectGenerator.Common;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

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
        public string? RootNamespace { get => _namespace; set => SetProperty(ref _namespace, value); }

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

                var solutionDir = await GetSolutionDirectoryAsync(cancellationToken).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(solutionDir) || !Directory.Exists(solutionDir))
                    throw new InvalidOperationException("No solution is open or the solution directory could not be resolved.");

                var fullProjectName = $"{rootNamespace}.{safeProjectname}";
                var projectDir = BuildProjectDirectory(solutionDir!, rootNamespace, safeProjectname);
                Directory.CreateDirectory(projectDir);

                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                var dte = _serviceProvider?.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
                if (dte?.Solution is null)
                    throw new InvalidOperationException("DTE/Solution service is not available.");

                var solution2 = (EnvDTE80.Solution2)dte.Solution;

                var extensionDir = Path.GetDirectoryName(typeof(GeneratorService).Assembly.Location)
                    ?? throw new InvalidOperationException("Extension directory could not be resolved.");

                var templateZipPath = Path.Combine(extensionDir, "Resources", "CleanArchitectureModule.zip");

                if (!File.Exists(templateZipPath))
                {
                    throw new FileNotFoundException(
                        $"Template zip was not found at expected path: {templateZipPath}",
                        templateZipPath);
                }

                var tempTemplateDir = Path.Combine(
                    Path.GetTempPath(),
                    "MarcusRunge.CleanArchitectureProjectGenerator",
                    Guid.NewGuid().ToString("N"));

                Directory.CreateDirectory(tempTemplateDir);

                ZipFile.ExtractToDirectory(templateZipPath, tempTemplateDir);

                var vstemplatePath = Directory
                    .GetFiles(tempTemplateDir, "*.vstemplate", SearchOption.AllDirectories)
                    .FirstOrDefault();

                if (string.IsNullOrWhiteSpace(vstemplatePath) || !File.Exists(vstemplatePath))
                {
                    throw new FileNotFoundException(
                        $"No .vstemplate file was found inside template zip: {templateZipPath}",
                        templateZipPath);
                }

                var templateRoot = Path.GetDirectoryName(vstemplatePath)
                    ?? throw new InvalidOperationException("Template root could not be resolved.");

                var templateCsprojPath = Path.Combine(templateRoot, "CleanArchitectureModule.csproj");

                if (!File.Exists(templateCsprojPath))
                {
                    throw new FileNotFoundException(
                        $"CleanArchitectureModule.csproj was not found next to the .vstemplate. Expected: {templateCsprojPath}",
                        templateCsprojPath);
                }
                EnsureTargetFrameworkInCsproj(templateCsprojPath, targetFramework);
                EnsurePropertyInCsproj(templateCsprojPath, "Nullable", "enable");
                EnsureCustomParameterInVstemplate(vstemplatePath, "$rootnamespace$", rootNamespace);
                solution2.AddFromTemplate(vstemplatePath, projectDir, safeProjectname, Exclusive: false);

                cancellationToken.ThrowIfCancellationRequested();

                var csprojPath = await WaitForCsprojAsync(projectDir, cancellationToken).ConfigureAwait(false);

                EnsurePropertyInCsproj(csprojPath, "Nullable", "enable");
                EnsureTargetFrameworkInCsproj(csprojPath, targetFramework);
                EnsurePropertyInCsproj(csprojPath, "RootNamespace", rootNamespace);
                EnsurePropertyInCsproj(csprojPath, "AssemblyName", safeProjectname);

                RootNamespace = rootNamespace;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                exceptionCallback?.Invoke(ex);
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
                    targets.Add($"net{version.Major}.0");
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

                if (!name.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                    continue;

                var version = name.TrimStart('v', 'V').Replace(".", "");
                targets.Add($"net{version}");
            }
        }

        private static string BuildProjectDirectory(string solutionDir, string rootNamespace, string safeProjectname)
        {
            var parts = rootNamespace
                .Split(['.'], StringSplitOptions.RemoveEmptyEntries)
                .Select(ToSafePathSegment)
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .ToArray();

            return Path.Combine([solutionDir, .. parts, .. new[] { ToSafePathSegment(safeProjectname) }]);
        }

        private static void EnsureCustomParameterInVstemplate(
                                            string vstemplatePath,
    string parameterName,
    string value)
        {
            var doc = XDocument.Load(vstemplatePath, LoadOptions.PreserveWhitespace);
            var root = doc.Root ?? throw new InvalidOperationException("Invalid vstemplate XML.");

            XNamespace ns = root.Name.Namespace;

            var templateContent = root.Element(ns + "TemplateContent");
            if (templateContent == null)
                throw new InvalidOperationException("Invalid vstemplate XML. Missing TemplateContent element.");

            var customParameters = templateContent.Element(ns + "CustomParameters");
            if (customParameters == null)
            {
                customParameters = new XElement(ns + "CustomParameters");
                templateContent.AddFirst(customParameters);
            }

            var existing = customParameters
                .Elements(ns + "CustomParameter")
                .FirstOrDefault(e =>
                    string.Equals(
                        (string?)e.Attribute("Name"),
                        parameterName,
                        StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                customParameters.Add(
                    new XElement(
                        ns + "CustomParameter",
                        new XAttribute("Name", parameterName),
                        new XAttribute("Value", value)));
            }
            else
            {
                existing.SetAttributeValue("Value", value);
            }

            doc.Save(vstemplatePath);
        }

        private static void EnsurePropertyInCsproj(string csprojPath, string propertyName, string value)
        {
            var doc = XDocument.Load(csprojPath, LoadOptions.PreserveWhitespace);
            var project = doc.Root ?? throw new InvalidOperationException("Invalid csproj XML (missing root element).");

            XNamespace ns = project.Name.Namespace;

            var propertyGroup =
                project.Elements(ns + "PropertyGroup").FirstOrDefault(pg => pg.Attribute("Condition") == null)
                ?? new XElement(ns + "PropertyGroup");

            if (propertyGroup.Parent is null)
                project.AddFirst(propertyGroup);

            var el = propertyGroup.Element(ns + propertyName);
            if (el == null)
            {
                el = new XElement(ns + propertyName);
                propertyGroup.Add(el);
            }

            el.Value = value;

            doc.Save(csprojPath);
        }

        private static void EnsureTargetFrameworkInCsproj(string csprojPath, string targetFramework)
        {
            var doc = XDocument.Load(csprojPath, LoadOptions.PreserveWhitespace);
            var project = doc.Root ?? throw new InvalidOperationException("Invalid csproj XML (missing root element).");

            XNamespace ns = project.Name.Namespace;

            var propertyGroup =
                project.Elements(ns + "PropertyGroup").FirstOrDefault(pg => pg.Attribute("Condition") == null)
                ?? new XElement(ns + "PropertyGroup");

            if (propertyGroup.Parent is null)
                project.AddFirst(propertyGroup);

            propertyGroup.Element(ns + "TargetFrameworks")?.Remove();

            var tf = propertyGroup.Element(ns + "TargetFramework");
            if (tf == null)
            {
                tf = new XElement(ns + "TargetFramework");
                propertyGroup.AddFirst(tf);
            }

            tf.Value = targetFramework;

            doc.Save(csprojPath);
        }

        private static string? GetRootProjectName(IVsHierarchy hierarchy)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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

        private static string ToSafePathSegment(string segment)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var filtered = new string([.. segment.Where(ch => !invalid.Contains(ch))]).Trim();
            return string.IsNullOrWhiteSpace(filtered) ? "_" : filtered;
        }

        private static IVsHierarchy? TryGetHierarchyFromCurrentSelection(IVsMonitorSelection monitorSelection)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IntPtr hierPtr = IntPtr.Zero;
            IntPtr selContainerPtr = IntPtr.Zero;

            try
            {
                int hr = monitorSelection.GetCurrentSelection(
                    out hierPtr,
                    out uint itemid,
                    out IVsMultiItemSelect? multiSelect,
                    out selContainerPtr);

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

            int hr = monitorSelection.GetCurrentElementValue(
                (uint)VSConstants.VSSELELEMID.SEID_StartupProject,
                out object value);

            if (!ErrorHandler.Succeeded(hr) || value is null)
                return null;

            if (value is IVsHierarchy hier)
                return hier;

            if (value is IntPtr ptr && ptr != IntPtr.Zero)
                return Marshal.GetObjectForIUnknown(ptr) as IVsHierarchy;

            return null;
        }

        private static async Task<string> WaitForCsprojAsync(string projectDir, CancellationToken ct)
        {
            for (int i = 0; i < 30; i++)
            {
                ct.ThrowIfCancellationRequested();

                var csproj = Directory.GetFiles(projectDir, "*.csproj", SearchOption.TopDirectoryOnly)
                                      .FirstOrDefault();

                if (!string.IsNullOrWhiteSpace(csproj))
                    return csproj;

                await Task.Delay(100, ct).ConfigureAwait(false);
            }

            throw new FileNotFoundException("Project file (*.csproj) was not created from template.", projectDir);
        }

        private async Task<string?> GetSolutionDirectoryAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            if (_serviceProvider.GetService(typeof(SVsSolution)) is not IVsSolution solution)
                return null;

            ErrorHandler.ThrowOnFailure(solution.GetSolutionInfo(out var solutionDir, out _, out _));
            return solutionDir;
        }
    }
}