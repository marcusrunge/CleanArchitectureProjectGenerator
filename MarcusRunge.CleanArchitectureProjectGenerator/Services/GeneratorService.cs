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
                // Stop immediately if the caller has already requested cancellation.
                cancellationToken.ThrowIfCancellationRequested();

                // Validate required user/project inputs before touching the file system or Visual Studio services.
                if (string.IsNullOrWhiteSpace(safeProjectname))
                    throw new ArgumentException("Project name must not be empty.", nameof(safeProjectname));

                if (string.IsNullOrWhiteSpace(rootNamespace))
                    throw new ArgumentException("Base namespace must not be empty.", nameof(rootNamespace));

                if (string.IsNullOrWhiteSpace(targetFramework))
                    throw new ArgumentException("Target framework (dotNetVersion) must not be empty.", nameof(targetFramework));

                // Resolve the currently opened solution directory; generation must happen inside an existing solution.
                var solutionDir = await GetSolutionDirectoryAsync(cancellationToken).ConfigureAwait(false);

                if (string.IsNullOrWhiteSpace(solutionDir) || !Directory.Exists(solutionDir))
                    throw new InvalidOperationException("No solution is open or the solution directory could not be resolved.");

                // Build the final project identity and destination path from the selected namespace and project name.
                var fullProjectName = $"{rootNamespace}.{safeProjectname}";
                var projectDir = BuildProjectDirectory(solutionDir!, rootNamespace, safeProjectname);

                // Ensure the physical target directory exists before Visual Studio adds the project from the template.
                Directory.CreateDirectory(projectDir);

                // Switch to the UI thread because DTE and Visual Studio shell services are apartment-threaded.
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                // Retrieve the Visual Studio automation object used to add projects to the current solution.
                var dte = _serviceProvider?.GetService(typeof(SDTE)) as EnvDTE80.DTE2;

                if (dte?.Solution is null)
                    throw new InvalidOperationException("DTE/Solution service is not available.");

                // Solution2 exposes AddFromTemplate, which is required for project-template based generation.
                var solution2 = (EnvDTE80.Solution2)dte.Solution;

                // Locate the installed extension directory so bundled template resources can be loaded.
                var extensionDir = Path.GetDirectoryName(typeof(GeneratorService).Assembly.Location)
                    ?? throw new InvalidOperationException("Extension directory could not be resolved.");

                // The clean architecture project template is packaged as a zip file within the extension resources.
                var templateZipPath = Path.Combine(extensionDir, "Resources", "CleanArchitectureModule.zip");

                if (!File.Exists(templateZipPath))
                {
                    throw new FileNotFoundException(
                        $"Template zip was not found at expected path: {templateZipPath}",
                        templateZipPath);
                }

                // Extract the template into an isolated temporary folder so template XML/project files can be customized.
                var tempTemplateDir = Path.Combine(
                    Path.GetTempPath(),
                    "MarcusRunge.CleanArchitectureProjectGenerator",
                    Guid.NewGuid().ToString("N"));

                Directory.CreateDirectory(tempTemplateDir);

                // Unpack the template archive before locating and modifying its .vstemplate and .csproj files.
                ZipFile.ExtractToDirectory(templateZipPath, tempTemplateDir);

                // Find the Visual Studio template manifest that describes how the project should be created.
                var vstemplatePath = Directory
                    .GetFiles(tempTemplateDir, "*.vstemplate", SearchOption.AllDirectories)
                    .FirstOrDefault();

                if (string.IsNullOrWhiteSpace(vstemplatePath) || !File.Exists(vstemplatePath))
                {
                    throw new FileNotFoundException(
                        $"No .vstemplate file was found inside template zip: {templateZipPath}",
                        templateZipPath);
                }

                // The template root is the folder containing the .vstemplate file and expected project file.
                var templateRoot = Path.GetDirectoryName(vstemplatePath)
                    ?? throw new InvalidOperationException("Template root could not be resolved.");

                // Locate the template project file so framework and compiler options can be adjusted before import.
                var templateCsprojPath = Path.Combine(templateRoot, "CleanArchitectureModule.csproj");

                if (!File.Exists(templateCsprojPath))
                {
                    throw new FileNotFoundException(
                        $"CleanArchitectureModule.csproj was not found next to the .vstemplate. Expected: {templateCsprojPath}",
                        templateCsprojPath);
                }

                // Configure the extracted template project before Visual Studio creates the actual solution project.
                EnsureTargetFrameworkInCsproj(templateCsprojPath, targetFramework);
                EnsurePropertyInCsproj(templateCsprojPath, "Nullable", "enable");
                EnsureCustomParameterInVstemplate(vstemplatePath, "$rootnamespace$", rootNamespace);

                // Add the customized project template to the current solution.
                solution2.AddFromTemplate(vstemplatePath, projectDir, safeProjectname, Exclusive: false);

                // Check again after template creation because project generation may take time.
                cancellationToken.ThrowIfCancellationRequested();

                // Wait until Visual Studio/template generation has produced the project file on disk.
                var csprojPath = await WaitForCsprojAsync(projectDir, cancellationToken).ConfigureAwait(false);

                // Normalize important project properties after creation to ensure the generated project is consistent.
                EnsurePropertyInCsproj(csprojPath, "Nullable", "enable");
                EnsureTargetFrameworkInCsproj(csprojPath, targetFramework);
                EnsurePropertyInCsproj(csprojPath, "RootNamespace", rootNamespace);
                EnsurePropertyInCsproj(csprojPath, "AssemblyName", safeProjectname);

                // Update bindable state so the UI reflects the namespace used for generation.
                RootNamespace = rootNamespace;
            }
            catch (OperationCanceledException)
            {
                // Preserve cancellation semantics so callers can distinguish cancellation from failure.
                throw;
            }
            catch (Exception ex)
            {
                // Report failures through the provided callback instead of throwing into the UI layer.
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
            // .NET SDKs are installed under Program Files\dotnet\sdk on a standard Windows installation.
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var sdkRoot = Path.Combine(programFiles, "dotnet", "sdk");

            // If the SDK directory does not exist, no SDK-based target frameworks can be discovered.
            if (!Directory.Exists(sdkRoot))
                return;

            // Each SDK folder name usually starts with a version number, for example "8.0.100".
            foreach (var dir in Directory.GetDirectories(sdkRoot))
            {
                var name = Path.GetFileName(dir);

                // Ignore preview/suffixed versions that cannot be parsed after removing the suffix.
                if (!Version.TryParse(name.Split('-')[0], out var version))
                    continue;

                // Modern SDK-style target frameworks use net{major}.0, for example net8.0.
                if (version.Major >= 5)
                    targets.Add($"net{version.Major}.0");
            }
        }

        private static void AddNetFrameworkTargets(HashSet<string> targets)
        {
            // .NET Framework reference assemblies are installed in the x86 Program Files folder.
            var referenceRoot =
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Reference Assemblies",
                    "Microsoft",
                    "Framework",
                    ".NETFramework"
                );

            // If the reference assemblies folder is missing, no .NET Framework targets are available.
            if (!Directory.Exists(referenceRoot))
                return;

            // Each framework folder is named with a leading "v", for example "v4.8".
            foreach (var dir in Directory.GetDirectories(referenceRoot))
            {
                var name = Path.GetFileName(dir);

                // Skip folders that do not follow the expected .NET Framework version folder format.
                if (!name.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Convert folder names such as "v4.8" to target framework monikers such as "net48".
                var version = name.TrimStart('v', 'V').Replace(".", "");
                targets.Add($"net{version}");
            }
        }

        private static string BuildProjectDirectory(string solutionDir, string rootNamespace, string safeProjectname)
        {
            // Split the namespace into folder segments so namespaces map naturally to directory structure.
            var parts = rootNamespace
                .Split(['.'], StringSplitOptions.RemoveEmptyEntries)
                .Select(ToSafePathSegment)
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .ToArray();

            // Combine the solution path, namespace folders, and project name into the final project directory.
            return Path.Combine([solutionDir, .. parts, .. new[] { ToSafePathSegment(safeProjectname) }]);
        }

        private static void EnsureCustomParameterInVstemplate(
                                            string vstemplatePath,
    string parameterName,
    string value)
        {
            // Load the .vstemplate while preserving formatting as much as possible.
            var doc = XDocument.Load(vstemplatePath, LoadOptions.PreserveWhitespace);

            // A valid Visual Studio template must have a root XML element.
            var root = doc.Root ?? throw new InvalidOperationException("Invalid vstemplate XML.");

            // Use the document namespace so element lookups work for namespaced template files.
            XNamespace ns = root.Name.Namespace;

            // Custom parameters must be located inside the TemplateContent section.
            var templateContent = root.Element(ns + "TemplateContent") ?? throw new InvalidOperationException("Invalid vstemplate XML. Missing TemplateContent element.");

            // Create the CustomParameters container if the template does not already define one.
            var customParameters = templateContent.Element(ns + "CustomParameters");

            if (customParameters == null)
            {
                customParameters = new XElement(ns + "CustomParameters");
                templateContent.AddFirst(customParameters);
            }

            // Look for an existing parameter to avoid adding duplicates.
            var existing = customParameters
                .Elements(ns + "CustomParameter")
                .FirstOrDefault(e =>
                    string.Equals(
                        (string?)e.Attribute("Name"),
                        parameterName,
                        StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                // Add the missing custom parameter so Visual Studio can replace it during template instantiation.
                customParameters.Add(
                    new XElement(
                        ns + "CustomParameter",
                        new XAttribute("Name", parameterName),
                        new XAttribute("Value", value)));
            }
            else
            {
                // Update the existing value to match the current generation request.
                existing.SetAttributeValue("Value", value);
            }

            // Persist the customized template manifest before it is used by Visual Studio.
            doc.Save(vstemplatePath);
        }

        private static void EnsurePropertyInCsproj(string csprojPath, string propertyName, string value)
        {
            // Load the project file while preserving whitespace to minimize formatting changes.
            var doc = XDocument.Load(csprojPath, LoadOptions.PreserveWhitespace);

            // A valid project file must have a root Project element.
            var project = doc.Root ?? throw new InvalidOperationException("Invalid csproj XML (missing root element).");

            // Respect the XML namespace used by the project file, if any.
            XNamespace ns = project.Name.Namespace;

            // Prefer an unconditional PropertyGroup so the property applies to all configurations.
            var propertyGroup =
                project.Elements(ns + "PropertyGroup").FirstOrDefault(pg => pg.Attribute("Condition") == null)
                ?? new XElement(ns + "PropertyGroup");

            // If no suitable PropertyGroup existed, add one at the top of the project file.
            if (propertyGroup.Parent is null)
                project.AddFirst(propertyGroup);

            // Find or create the requested property element.
            var el = propertyGroup.Element(ns + propertyName);

            if (el == null)
            {
                el = new XElement(ns + propertyName);
                propertyGroup.Add(el);
            }

            // Set the requested property value and save the project file.
            el.Value = value;
            doc.Save(csprojPath);
        }

        private static void EnsureTargetFrameworkInCsproj(string csprojPath, string targetFramework)
        {
            // Load the project file so target framework information can be normalized.
            var doc = XDocument.Load(csprojPath, LoadOptions.PreserveWhitespace);

            // A valid project file must have a root Project element.
            var project = doc.Root ?? throw new InvalidOperationException("Invalid csproj XML (missing root element).");

            // Use the project namespace for reliable XML element lookup.
            XNamespace ns = project.Name.Namespace;

            // Prefer an unconditional PropertyGroup because the selected framework should apply globally.
            var propertyGroup =
                project.Elements(ns + "PropertyGroup").FirstOrDefault(pg => pg.Attribute("Condition") == null)
                ?? new XElement(ns + "PropertyGroup");

            // Add a new PropertyGroup if the project did not contain an unconditional one.
            if (propertyGroup.Parent is null)
                project.AddFirst(propertyGroup);

            // Remove multi-targeting because this generator creates a project for one selected target framework.
            propertyGroup.Element(ns + "TargetFrameworks")?.Remove();

            // Find or create the single TargetFramework element.
            var tf = propertyGroup.Element(ns + "TargetFramework");

            if (tf == null)
            {
                tf = new XElement(ns + "TargetFramework");
                propertyGroup.AddFirst(tf);
            }

            // Apply the selected target framework and persist the project file.
            tf.Value = targetFramework;
            doc.Save(csprojPath);
        }

        private static string? GetRootProjectName(IVsHierarchy hierarchy)
        {
            // Visual Studio hierarchy properties must be read on the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            // Prefer the project caption because it is the display name shown in Solution Explorer.
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

            // Fall back to the internal project name if no caption is available.
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

            // Return null when no usable project name could be read from the hierarchy.
            return null;
        }

        private static string ToSafePathSegment(string segment)
        {
            // Remove characters that are invalid in file or directory names.
            var invalid = Path.GetInvalidFileNameChars();
            var filtered = new string([.. segment.Where(ch => !invalid.Contains(ch))]).Trim();

            // Use an underscore as a safe fallback when the segment becomes empty after filtering.
            return string.IsNullOrWhiteSpace(filtered) ? "_" : filtered;
        }

        private static IVsHierarchy? TryGetHierarchyFromCurrentSelection(IVsMonitorSelection monitorSelection)
        {
            // Current selection information is provided by the Visual Studio shell and requires the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            // Native pointers returned by GetCurrentSelection must be released after use.
            IntPtr hierPtr = IntPtr.Zero;
            IntPtr selContainerPtr = IntPtr.Zero;

            try
            {
                // Read the currently selected hierarchy item from Solution Explorer.
                int hr = monitorSelection.GetCurrentSelection(
                    out hierPtr,
                    out uint itemid,
                    out IVsMultiItemSelect? multiSelect,
                    out selContainerPtr);

                // If no hierarchy is selected, there is no project context to infer.
                if (!ErrorHandler.Succeeded(hr) || hierPtr == IntPtr.Zero)
                    return null;

                // Convert the COM hierarchy pointer to the managed IVsHierarchy interface.
                return Marshal.GetObjectForIUnknown(hierPtr) as IVsHierarchy;
            }
            finally
            {
                // Release COM pointers to avoid leaking Visual Studio shell references.
                if (hierPtr != IntPtr.Zero) Marshal.Release(hierPtr);
                if (selContainerPtr != IntPtr.Zero) Marshal.Release(selContainerPtr);
            }
        }

        private static IVsHierarchy? TryGetStartupProjectHierarchy(IVsMonitorSelection monitorSelection)
        {
            // Startup project information is part of the Visual Studio selection context and requires the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            // Query the shell for the current startup project element.
            int hr = monitorSelection.GetCurrentElementValue(
                (uint)VSConstants.VSSELELEMID.SEID_StartupProject,
                out object value);

            // No startup project is available if the shell call fails or returns no value.
            if (!ErrorHandler.Succeeded(hr) || value is null)
                return null;

            // The shell may return the hierarchy directly.
            if (value is IVsHierarchy hier)
                return hier;

            // Some shell implementations return a COM pointer that must be converted to IVsHierarchy.
            if (value is IntPtr ptr && ptr != IntPtr.Zero)
                return Marshal.GetObjectForIUnknown(ptr) as IVsHierarchy;

            // Unsupported return type; no hierarchy can be inferred.
            return null;
        }

        private static async Task<string> WaitForCsprojAsync(string projectDir, CancellationToken ct)
        {
            // Poll briefly because Visual Studio template creation may write the project file asynchronously.
            for (int i = 0; i < 30; i++)
            {
                // Respect cancellation while waiting for the project file to appear.
                ct.ThrowIfCancellationRequested();

                // Look for the generated project file in the root of the target project directory.
                var csproj = Directory.GetFiles(projectDir, "*.csproj", SearchOption.TopDirectoryOnly)
                                      .FirstOrDefault();

                // Return as soon as the project file is available.
                if (!string.IsNullOrWhiteSpace(csproj))
                    return csproj;

                // Wait a short interval before checking again.
                await Task.Delay(100, ct).ConfigureAwait(false);
            }

            // Fail explicitly if the template did not create a project file within the expected time.
            throw new FileNotFoundException("Project file (*.csproj) was not created from template.", projectDir);
        }

        private async Task<string?> GetSolutionDirectoryAsync(CancellationToken cancellationToken)
        {
            // Stop immediately if the caller no longer needs the solution directory.
            cancellationToken.ThrowIfCancellationRequested();

            // IVsSolution is a Visual Studio shell service and must be accessed on the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Retrieve the currently loaded solution service.
            if (_serviceProvider.GetService(typeof(SVsSolution)) is not IVsSolution solution)
                return null;

            // Ask Visual Studio for the solution directory and throw if the shell reports failure.
            ErrorHandler.ThrowOnFailure(solution.GetSolutionInfo(out var solutionDir, out _, out _));

            // Return the physical folder containing the open solution.
            return solutionDir;
        }
    }
}