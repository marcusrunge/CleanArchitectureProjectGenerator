using EnvDTE;
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
    internal class GeneratorService([Import(typeof(SVsServiceProvider))] IServiceProvider serviceProvider) : BindableBase, IGeneratorService
    {
        /// <summary>
        /// Visual Studio service provider (SVsServiceProvider) used to retrieve shell services.
        /// </summary>
        private readonly IServiceProvider _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));

        // Backing field for the bindable Namespace property.
        private string? _namespace;

        // Backing field for the selected solution folder relative path, used to determine where to place the generated project within the solution.
        private string? _selectedSolutionFolderRelativePath;

        /// <inheritdoc/>
        public string? RootNamespace { get => _namespace; set => SetProperty(ref _namespace, value); }

        /// <inheritdoc/>
        public async Task CreateAsync(string safeProjectname, string rootNamespace, string targetFramework, Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            // Initialize a variable to hold the path to the temporary template directory, which will be used for extracting and customizing the project template.
            string? tempTemplateDir = null;
            try
            {
                // Stop immediately if the caller has already requested cancellation.
                cancellationToken.ThrowIfCancellationRequested();

                // Validate required user/project inputs before touching the file system or Visual Studio services.
                if (string.IsNullOrWhiteSpace(safeProjectname))
                    throw new ArgumentException("Project name must not be empty.", nameof(safeProjectname));
                // Validate that the root namespace is provided, as it is essential for generating the project structure and namespaces.
                if (string.IsNullOrWhiteSpace(rootNamespace))
                    throw new ArgumentException("Base namespace must not be empty.", nameof(rootNamespace));
                // Validate that the target framework is provided, as it is essential for generating the project with the correct framework settings.
                if (string.IsNullOrWhiteSpace(targetFramework))
                    throw new ArgumentException("Target framework (dotNetVersion) must not be empty.", nameof(targetFramework));

                // Resolve the currently opened solution directory; generation must happen inside an existing solution.
                var solutionDir = await GetSolutionDirectoryAsync(cancellationToken).ConfigureAwait(false);
                // If the solution directory could not be resolved, throw an exception to indicate that project generation cannot proceed.
                if (string.IsNullOrWhiteSpace(solutionDir) || !Directory.Exists(solutionDir))
                    // If the solution directory is not available, it indicates that no solution is open or the path could not be determined.
                    throw new InvalidOperationException("No solution is open or the solution directory could not be resolved.");

                // Build the full project name and target directory based on the current solution-folder context.
                var selectedSolutionFolderRelativePath = _selectedSolutionFolderRelativePath;
                // Normalize the root namespace to ensure it is a valid namespace segment, removing any invalid characters or formatting issues.
                var baseNamespace = NormalizeNamespace(rootNamespace);

                // Extract the last segment of the user-entered project name to use as the short project name for class and namespace generation.
                var projectNamePart = ExtractLastPathOrNamespaceSegment(safeProjectname);
                // Build the short project name by removing the root namespace prefix from the user-entered project name, if it exists, to allow for cleaner class and namespace generation within the template.
                var shortProjectName = BuildShortProjectName(baseNamespace, projectNamePart);
                // Normalize the short project name to ensure it is a valid namespace segment, removing any invalid characters or formatting issues.
                shortProjectName = NormalizeNamespaceSegment(shortProjectName);
                // Build the full project name by combining the base namespace and the short project name, ensuring proper formatting and avoiding duplication.
                var fullProjectName = BuildFullProjectName(baseNamespace, shortProjectName);
                // Determine the target project directory based on the solution directory, optional solution folder path, and the full project name.
                var projectDir = BuildProjectDirectory(solutionDir!, selectedSolutionFolderRelativePath, fullProjectName);

                // Ensure the physical target directory exists before Visual Studio adds the project from the template.
                Directory.CreateDirectory(projectDir);

                // Switch to the UI thread because DTE and Visual Studio shell services are apartment-threaded.
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                // Retrieve the Visual Studio automation object used to add projects to the current solution.
                var dte = _serviceProvider?.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
                // Check that the DTE service and solution are available; if not, throw an exception to indicate that project generation cannot proceed.
                if (dte?.Solution is null)
                    throw new InvalidOperationException("DTE/Solution service is not available.");

                // Solution2 exposes AddFromTemplate, which is required for project-template based generation.
                var solution2 = (EnvDTE80.Solution2)dte.Solution;

                // Locate the installed extension directory so bundled template resources can be loaded.
                var extensionDir = Path.GetDirectoryName(typeof(GeneratorService).Assembly.Location) ?? throw new InvalidOperationException("Extension directory could not be resolved.");

                // The clean architecture project template is packaged as a zip file within the extension resources.
                var templateZipPath = Path.Combine(extensionDir, "Resources", "CleanArchitectureModule.zip");
                // Check that the template zip file exists; if not, throw a FileNotFoundException to indicate that the required template resource is missing.
                if (!File.Exists(templateZipPath))
                {
                    // If the template zip file is missing, throw a FileNotFoundException to indicate that the required template resource is not available for project generation.
                    throw new FileNotFoundException($"Template zip was not found at expected path: {templateZipPath}", templateZipPath);
                }

                // Extract the template into an isolated temporary folder so template XML/project files can be customized.
                tempTemplateDir = Path.Combine(Path.GetTempPath(), "MarcusRunge.CleanArchitectureProjectGenerator", Guid.NewGuid().ToString("N"));
                // Ensure the temporary template directory exists before extracting the template zip file.
                Directory.CreateDirectory(tempTemplateDir);

                // Unpack the template archive before locating and modifying its .vstemplate and .csproj files.
                ZipFile.ExtractToDirectory(templateZipPath, tempTemplateDir);

                // Find the Visual Studio template manifest that describes how the project should be created.
                var vstemplatePath = Directory.GetFiles(tempTemplateDir, "*.vstemplate", SearchOption.AllDirectories).FirstOrDefault();
                // Check that the .vstemplate file was found; if not, throw a FileNotFoundException to indicate that the required template manifest is missing from the extracted template.
                if (string.IsNullOrWhiteSpace(vstemplatePath) || !File.Exists(vstemplatePath))
                {
                    // If the .vstemplate file is missing,
                    throw new FileNotFoundException($"No .vstemplate file was found inside template zip: {templateZipPath}", templateZipPath);
                }

                // The template root is the folder containing the .vstemplate file and expected project file.
                var templateRoot = Path.GetDirectoryName(vstemplatePath) ?? throw new InvalidOperationException("Template root could not be resolved.");

                // Locate the template project file so framework and compiler options can be adjusted before import.
                var templateCsprojPath = Path.Combine(templateRoot, "CleanArchitectureModule.csproj");
                // Check that the template .csproj file exists; if not, throw a FileNotFoundException to indicate that the required project file is missing from the extracted template.
                if (!File.Exists(templateCsprojPath))
                {
                    // If the template .csproj file is missing,
                    throw new FileNotFoundException($"CleanArchitectureModule.csproj was not found next to the .vstemplate. Expected: {templateCsprojPath}", templateCsprojPath);
                }

                // Configure the extracted template project before Visual Studio creates the actual solution project.
                EnsureTargetFrameworkInCsproj(templateCsprojPath, targetFramework);
                EnsurePropertyInCsproj(templateCsprojPath, "Nullable", "enable");

                // Pass important information to the template through custom parameters so it can be used in template variable replacements.
                EnsureCustomParameterInVstemplate(vstemplatePath, "$rootnamespace$", fullProjectName);
                EnsureCustomParameterInVstemplate(vstemplatePath, "$projectnamespace$", fullProjectName);
                EnsureCustomParameterInVstemplate(vstemplatePath, "$shortprojectname$", shortProjectName);
                EnsureCustomParameterInVstemplate(vstemplatePath, "$basenamespace$", baseNamespace);

                // Add the customized project template to the solution or to the selected solution folder.
                AddProjectFromTemplate(dte, solution2, vstemplatePath, projectDir, fullProjectName, _selectedSolutionFolderRelativePath);

                // Check again after template creation because project generation may take time.
                cancellationToken.ThrowIfCancellationRequested();

                // Wait until Visual Studio/template generation has produced the project file on disk.
                var csprojPath = await WaitForCsprojAsync(projectDir, cancellationToken).ConfigureAwait(false);

                // Normalize important project properties after creation to ensure the generated project is consistent.
                EnsurePropertyInCsproj(csprojPath, "Nullable", "enable");
                EnsureTargetFrameworkInCsproj(csprojPath, targetFramework);

                // Important:
                // These should match the actual generated project identity.
                EnsurePropertyInCsproj(csprojPath, "RootNamespace", fullProjectName);
                EnsurePropertyInCsproj(csprojPath, "AssemblyName", fullProjectName);

                // Update bindable state so the UI reflects the namespace used for generation.
                //
                // Keep this as the root/base namespace, not the generated project namespace.
                // Example:
                // RootNamespace = SmallBusinessOperations
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
            finally
            {
                // Clean up the temporary template directory to avoid leaving behind extracted files.
                TryDeleteDirectory(tempTemplateDir);
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

            // Add TFMs from .NET Standard reference assemblies and SDK packs (e.g., netstandard2.0, netstandard2.1).
            AddNetStandardTargets(results);

            // Return a stable, sorted, read-only list for UI binding and deterministic behavior.
            return results.OrderBy(TargetFrameworkSortKey).ToList().AsReadOnly();
        });

        /// <inheritdoc/>
        public async Task InitializeAsync(Action<Exception> exceptionCallback, CancellationToken cancellationToken)
        {
            // Ensure this method is called on the UI thread because it accesses DTE project properties.
            try
            {
                // Stop immediately if the caller has already requested cancellation.
                cancellationToken.ThrowIfCancellationRequested();
                // Switch to the UI thread because DTE and Visual Studio shell services are apartment-threaded.
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                // Retrieve the Visual Studio automation object used to inspect the current solution and selection.
                if (_serviceProvider.GetService(typeof(SDTE)) is not EnvDTE80.DTE2 dte || dte.Solution == null || string.IsNullOrWhiteSpace(dte.Solution.FullName))
                {
                    // If the DTE service or solution is not available, clear the namespace and selected folder state.
                    RootNamespace = null;
                    // Clear the selected solution folder relative path since no solution is open.
                    _selectedSolutionFolderRelativePath = null;
                    // Exit early since there is no solution context to initialize from.
                    return;
                }
                // Extract the solution name from the full solution path, removing the file extension to get a clean base name.
                var solutionName = Path.GetFileNameWithoutExtension(dte.Solution.FullName);
                // If the solution name is null, empty, or whitespace, clear the namespace and selected folder state.
                if (string.IsNullOrWhiteSpace(solutionName))
                {
                    // Clear the RootNamespace property since no valid solution name could be determined.
                    RootNamespace = null;
                    // Clear the selected solution folder relative path since no valid solution name could be determined.
                    _selectedSolutionFolderRelativePath = null;
                    // Exit early since there is no valid solution name to derive a namespace from.
                    return;
                }
                // Normalize the solution name to ensure it is a valid namespace segment, removing any invalid characters or formatting issues.
                solutionName = NormalizeNamespace(solutionName);
                // Attempt to determine the relative path of the currently selected solution folder in the Solution Explorer, if any.
                _selectedSolutionFolderRelativePath = TryGetSelectedSolutionFolderRelativePath(dte);
                // If no solution folder is selected, use the solution name as the root namespace.
                if (string.IsNullOrWhiteSpace(_selectedSolutionFolderRelativePath))
                {
                    // If no solution folder is selected, use the solution name as the root namespace.
                    RootNamespace = solutionName;
                    // Exit early since there is no selected solution folder to derive a namespace suffix from.
                    return;
                }

                // Convert the selected solution folder relative path into a namespace-like suffix by replacing directory separators with dots and normalizing each segment.
                var folderNamespace = ToNamespaceSuffix(_selectedSolutionFolderRelativePath);
                // If the folder namespace is empty or whitespace, use the solution name as the root namespace.
                if (string.IsNullOrWhiteSpace(folderNamespace))
                {
                    // If the folder namespace is empty, use the solution name as the root namespace.
                    RootNamespace = solutionName;
                    // Exit early since there is no valid folder namespace to append to the solution name.
                    return;
                }
                // Combine the solution name and folder namespace to form the full root namespace for the generated project.
                RootNamespace = BuildFullProjectName(solutionName, folderNamespace);
            }
            // Catch any exceptions that occur during initialization to prevent the extension from crashing and to allow the caller to handle errors gracefully.
            catch (Exception ex)
            {
                // If an exception occurs, clear the namespace and selected folder state to avoid using potentially invalid data.
                RootNamespace = null;
                // Clear the selected solution folder relative path since an error occurred during initialization.
                _selectedSolutionFolderRelativePath = null;
                // Invoke the provided exception callback to report the error to the caller, allowing for logging or user notification.
                exceptionCallback?.Invoke(ex);
            }
        }

        // Discovers installed .NET SDKs by inspecting the standard SDK installation directory and inferring target frameworks from folder names.
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

        // .NET Framework reference assemblies are installed in the x86 Program Files folder, even on 64-bit machines.
        private static void AddNetFrameworkTargets(HashSet<string> targets)
        {
            // .NET Framework reference assemblies are installed in the x86 Program Files folder.
            var referenceRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Reference Assemblies", "Microsoft", "Framework", ".NETFramework");

            // If the reference assemblies folder is missing, no .NET Framework targets are available.
            if (!Directory.Exists(referenceRoot))
                return;

            // Each framework folder is named with a leading "v", for example "v4.8".
            foreach (var dir in Directory.GetDirectories(referenceRoot))
            {
                // Extract the folder name to parse the version.
                var name = Path.GetFileName(dir);

                // Skip folders that do not follow the expected .NET Framework version folder format.
                if (!name.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Convert folder names such as "v4.8" to target framework monikers such as "net48".
                var version = name.TrimStart('v', 'V').Replace(".", "");
                targets.Add($"net{version}");
            }
        }

        // Adds .NET Standard target frameworks to the provided set by checking installed SDKs and reference assemblies, ensuring that only supported versions are included.
        private static void AddNetStandardTargets(HashSet<string> targets)
        {
            // Conservative baseline: netstandard2.0 is broadly supported by SDK-style projects.
            targets.Add("netstandard2.0");

            // netstandard2.1 is available with newer SDKs, but not supported by .NET Framework consumers.
            if (IsNetStandard21Available())
                targets.Add("netstandard2.1");

            // Optional: discover older installed .NET Standard reference folders if present.
            AddNetStandardTargetsFromReferenceAssemblies(targets);
            AddNetStandardTargetsFromSdkPacks(targets);
        }

        // .NET Standard reference assemblies are installed in the x86 Program Files folder, even on 64-bit machines.
        private static void AddNetStandardTargetsFromReferenceAssemblies(HashSet<string> targets)
        {
            // .NET Standard reference assemblies are installed in the x86 Program Files folder, even on 64-bit machines.
            var referenceRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Reference Assemblies", "Microsoft", "Framework", ".NETStandard");

            // If the reference assemblies folder is missing, no .NET Standard targets can be discovered.
            if (!Directory.Exists(referenceRoot))
                return;

            // Each .NET Standard folder is named with a leading "v", for example "v2.0" or "v2.1".
            foreach (var dir in Directory.GetDirectories(referenceRoot))
            {
                // Extract the folder name to parse the version.
                var name = Path.GetFileName(dir);

                // Skip folders that do not follow the expected .NET Standard version folder format.
                if (!name.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Convert folder names such as "v2.0" to target framework monikers such as "netstandard2.0".
                var version = name.TrimStart('v', 'V');

                // Only add valid, non-empty versions to the target list.
                if (!string.IsNullOrWhiteSpace(version))
                    targets.Add($"netstandard{version}");
            }
        }

        // .NET Standard reference packs are installed under Program Files\dotnet\packs\NETStandard.Library.Ref on Windows.
        private static void AddNetStandardTargetsFromSdkPacks(HashSet<string> targets)
        {
            // .NET Standard reference packs are installed under Program Files\dotnet\packs\NETStandard.Library.Ref on Windows.
            var packsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "packs", "NETStandard.Library.Ref");

            // If the packs directory does not exist, no .NET Standard targets can be discovered from SDK packs.
            if (!Directory.Exists(packsRoot))
                return;

            // Each version folder contains a "ref" subfolder with target framework directories (e.g., netstandard2.0, netstandard2.1).
            foreach (var versionDir in Directory.GetDirectories(packsRoot))
            {
                // The "ref" subdirectory contains the reference assemblies for each target framework.
                var refDir = Path.Combine(versionDir, "ref");

                // If the "ref" directory does not exist, skip this version folder.
                if (!Directory.Exists(refDir))
                    continue;

                // Each subdirectory under "ref" corresponds to a target framework moniker (TFM).
                foreach (var tfmDir in Directory.GetDirectories(refDir))
                {
                    // Extract the TFM folder name, which is typically in the format "netstandard2.0" or "netstandard2.1".
                    var tfm = Path.GetFileName(tfmDir);

                    // Only add TFMs that start with "netstandard" to the target list, ignoring any unexpected folders.
                    if (tfm.StartsWith("netstandard", StringComparison.OrdinalIgnoreCase))
                        targets.Add(tfm);
                }
            }
        }

        // Adds a new project to the current solution from a specified Visual Studio template, optionally placing it inside a selected solution folder.
        private static Project AddProjectFromTemplate(EnvDTE80.DTE2 dte, EnvDTE80.Solution2 solution2, string vstemplatePath, string projectDir, string fullProjectName, string? selectedSolutionFolderRelativePath)
        {
            // Ensure this method is called on the UI thread because DTE and Visual Studio shell services are apartment-threaded.
            ThreadHelper.ThrowIfNotOnUIThread();
            // If a solution folder is selected, find the corresponding SolutionFolder project to add the new project into it.
            if (!string.IsNullOrWhiteSpace(selectedSolutionFolderRelativePath))
            {
                // Attempt to locate the solution folder project by its relative path within the solution.
                if (!TryFindSolutionFolderProject(dte.Solution, selectedSolutionFolderRelativePath, out var solutionFolderProject))
                {
                    // If the solution folder project could not be found, throw an exception to indicate that the specified folder does not exist in the solution.
                    throw new InvalidOperationException($"The selected solution folder could not be found: {selectedSolutionFolderRelativePath}");
                }
                // Ensure that the located project is indeed a SolutionFolder; if not, throw an exception to indicate that the selected project is not a valid solution folder.
                if (solutionFolderProject.Object is not EnvDTE80.SolutionFolder solutionFolder)
                {
                    // If the located project is not a SolutionFolder, throw an exception to indicate that the selected project is not a valid solution folder.
                    throw new InvalidOperationException($"The selected project is not a valid solution folder: {selectedSolutionFolderRelativePath}");
                }
                // Add the new project from the template into the located solution folder, specifying the target directory and full project name.
                return solutionFolder.AddFromTemplate(vstemplatePath, projectDir, fullProjectName);
            }
            // If no solution folder is selected, add the new project directly to the root of the solution.
            return solution2.AddFromTemplate(vstemplatePath, projectDir, fullProjectName, Exclusive: false);
        }

        // Builds the full project name (used for namespaces and assembly names) by combining the root namespace and the user-entered project name, ensuring proper formatting and avoiding duplication.
        private static string BuildFullProjectName(string rootNamespace, string projectName)
        {
            // Validate inputs to ensure the resulting project name can be constructed meaningfully.
            if (string.IsNullOrWhiteSpace(rootNamespace))
                throw new ArgumentException("Root namespace must not be empty.", nameof(rootNamespace));
            // The project name is required to build the full project name, even if it ends up being the same as the root namespace.
            if (string.IsNullOrWhiteSpace(projectName))
                throw new ArgumentException("Project name must not be empty.", nameof(projectName));
            // Clean up the inputs by trimming whitespace and dots to ensure consistent formatting and comparison.
            var cleanRootNamespace = rootNamespace.Trim().Trim('.');
            var cleanProjectName = projectName.Trim().Trim('.');
            // If the project name is exactly the same as the root namespace, return it directly to avoid duplication.
            if (cleanProjectName.Equals(cleanRootNamespace, StringComparison.OrdinalIgnoreCase))
                return cleanRootNamespace;
            // If the project name already starts with the root namespace followed by a dot, return it as is to avoid adding the root namespace twice.
            if (cleanProjectName.StartsWith(cleanRootNamespace + ".", StringComparison.OrdinalIgnoreCase))
                return cleanProjectName;
            // Otherwise, combine the root namespace and project name with a dot separator to form the full project name.
            var shortProjectName = BuildShortProjectName(rootNamespace, cleanProjectName);
            // The full project name is the root namespace followed by the short project name, ensuring that the root namespace is not duplicated if it was included in the project name.
            return $"{cleanRootNamespace}.{shortProjectName}";
        }

        // Builds the target project directory by combining the solution directory, optional relative solution folder path, and the full project name, ensuring that the resulting path is valid and safe for use in the file system.
        private static string BuildProjectDirectory(string solutionDir, string? relativeSolutionFolderPath, string fullProjectName)
        {
            // Validate inputs to ensure the resulting project directory can be constructed meaningfully.
            if (string.IsNullOrWhiteSpace(solutionDir))
                throw new ArgumentException("Solution directory must not be empty.", nameof(solutionDir));
            // The full project name is required to determine the final project directory; it must not be empty or whitespace.
            if (string.IsNullOrWhiteSpace(fullProjectName))
                throw new ArgumentException("Full project name must not be empty.", nameof(fullProjectName));
            // Normalize the solution directory to ensure it is a valid path and does not contain any trailing directory separators that could affect path combination.
            var baseDirectory = solutionDir;
            // If a relative solution folder path is provided, split it into safe segments and combine them with the base directory to form the target project directory.
            if (!string.IsNullOrWhiteSpace(relativeSolutionFolderPath))
            {
                // Split the relative solution folder path into segments, normalize them to safe path segments, and filter out any empty or whitespace segments to ensure a valid directory structure.
                var safeSegments = relativeSolutionFolderPath?.Replace('\\', '/').Split(['/'], StringSplitOptions.RemoveEmptyEntries).Select(ToSafePathSegment).Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
                // If there are any valid safe segments, combine them with the base directory to form the target project directory.
                if (safeSegments?.Length > 0)
                {
                    // Combine each safe segment with the base directory to build the full path to the target project directory.
                    foreach (var segment in safeSegments)
                        // Combine the current base directory with the next safe segment to progressively build the full path.
                        baseDirectory = Path.Combine(baseDirectory, segment);
                }
            }
            // Combine the base directory (which may include solution folder segments) with the safe version of the full project name to determine the final project directory for generation.
            var projectDirectory = Path.Combine(baseDirectory, ToSafePathSegment(fullProjectName));
            // Validate that the computed project directory is indeed a subdirectory of the solution directory to prevent accidental generation outside the solution scope.
            EnsurePathIsInsideDirectory(solutionDir, projectDirectory);
            // Return the computed project directory for use in project generation.
            return projectDirectory;
        }

        // Builds a short project name for template parameters by removing the root namespace prefix from the user-entered project name, if it exists, to allow for cleaner class and namespace generation within the template.
        private static string BuildShortProjectName(string rootNamespace, string projectName)
        {
            // Validate inputs to ensure the method can perform string manipulations safely.
            if (string.IsNullOrWhiteSpace(rootNamespace))
                throw new ArgumentException("Root namespace must not be empty.", nameof(rootNamespace));
            // The project name is required to determine the short name, even if it ends up being the same as the root namespace.
            if (string.IsNullOrWhiteSpace(projectName))
                throw new ArgumentException("Project name must not be empty.", nameof(projectName));
            // Clean up the inputs by trimming whitespace and dots to ensure consistent comparison and formatting.
            var cleanRootNamespace = rootNamespace.Trim().Trim('.');
            var cleanProjectName = projectName.Trim().Trim('.');
            // If the project name is exactly the same as the root namespace, return it as the short name to avoid empty strings.
            if (cleanProjectName.Equals(cleanRootNamespace, StringComparison.OrdinalIgnoreCase))
            {
                // The short name is the same as the root namespace in this case, so return it directly.
                var lastDotIndex = cleanProjectName.LastIndexOf('.');
                // If there is a dot in the project name, return the substring after the last dot to get the short name. This handles cases where the project name has multiple segments but is identical to the root namespace, for example "Company.Product" with root namespace "Company" would yield "Product".
                if (lastDotIndex >= 0 && lastDotIndex < cleanProjectName.Length - 1)
                    return cleanProjectName.Substring(lastDotIndex + 1);

                return cleanProjectName;
            }
            // If the project name starts with the root namespace followed by a dot, remove that prefix to get the short name.
            if (cleanProjectName.StartsWith(cleanRootNamespace + ".", StringComparison.OrdinalIgnoreCase))
            {
                // The short name is the part of the project name that comes after the root namespace and the following dot.
                return cleanProjectName.Substring(cleanRootNamespace.Length + 1);
            }
            // If the project name does not start with the root namespace, return the last segment after the final dot as the short name, or the full project name if there are no dots.
            var projectNameLastDotIndex = cleanProjectName.LastIndexOf('.');
            // If there is a dot and it's not the last character, return the substring after the last dot as the short name.
            if (projectNameLastDotIndex >= 0 && projectNameLastDotIndex < cleanProjectName.Length - 1)
            {
                // This handles cases where the project name has multiple segments but does not start with the root namespace, for example "Company.Product.Module" would yield "Module".
                return cleanProjectName.Substring(projectNameLastDotIndex + 1);
            }
            // If there are no dots, the short name is the same as the project name.
            return cleanProjectName;
        }

        // Ensures a custom parameter with the specified name and value exists in the .vstemplate file, which allows passing information to the template during project creation.
        private static void EnsureCustomParameterInVstemplate(string vstemplatePath, string parameterName, string value)
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
            var existing = customParameters.Elements(ns + "CustomParameter").FirstOrDefault(e => string.Equals((string?)e.Attribute("Name"), parameterName, StringComparison.OrdinalIgnoreCase));

            if (existing == null)
            {
                // Add the missing custom parameter so Visual Studio can replace it during template instantiation.
                customParameters.Add(new XElement(ns + "CustomParameter", new XAttribute("Name", parameterName), new XAttribute("Value", value)));
            }
            else
            {
                // Update the existing value to match the current generation request.
                existing.SetAttributeValue("Value", value);
            }

            // Persist the customized template manifest before it is used by Visual Studio.
            doc.Save(vstemplatePath);
        }

        // Ensures that the candidate path is a subdirectory of the specified root directory, throwing an exception if it is not, to prevent project generation outside the solution folder.
        private static void EnsurePathIsInsideDirectory(string rootDirectory, string candidatePath)
        {
            // Normalize both paths to absolute paths and ensure they end with a directory separator for accurate prefix comparison.
            var root = Path.GetFullPath(rootDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;

            // Normalize the candidate path similarly to ensure consistent comparison.
            var candidate = Path.GetFullPath(candidatePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            // Check if the candidate path starts with the root path, ignoring case, to determine if it is a subdirectory.
            if (!candidate.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project directory must be inside the solution directory.");
        }

        // Ensures the specified property is set to the given value in the project file, creating the property if it does not exist.
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

        // Ensures the specified target framework is set in the project file, replacing any existing TargetFramework or TargetFrameworks properties.
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

        // Extracts the last segment of a path or namespace-like string, which is useful for determining the project name from user input that may include directories or namespace separators.
        private static string ExtractLastPathOrNamespaceSegment(string value)
        {
            // If the input is null, empty, or whitespace, return an empty string to avoid processing invalid data.
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;
            // Normalize the input by trimming whitespace, replacing backslashes with forward slashes, and removing leading/trailing slashes or dots to prepare for splitting.
            var normalized = value.Trim().Replace('\\', '/').Trim('/', '.');
            // Split the normalized input into segments based on slashes and dots, trimming whitespace and filtering out empty segments to isolate the last meaningful segment.
            var parts = normalized.Split(['/', '.'], StringSplitOptions.RemoveEmptyEntries).Select(p => p.Trim()).Where(p => !string.IsNullOrWhiteSpace(p)).ToArray();
            // If no valid segments are found, return the normalized input as a fallback.
            if (parts.Length == 0)
                return normalized;
            // Return the last segment, which represents the most specific part of the path or namespace, suitable for use as a project name or namespace segment.
            return parts[parts.Length - 1];
        }

        // Recursively constructs the relative path of a solution folder project within the solution hierarchy by traversing its parent projects.
        private static string GetSolutionFolderRelativePath(Project solutionFolderProject)
        {
            // Ensure this method is called on the UI thread because it accesses DTE project properties.
            ThreadHelper.ThrowIfNotOnUIThread();
            // Initialize a list to hold the segments of the solution folder path, starting from the innermost folder and working outward.
            var segments = new List<string>();
            // Start with the provided solution folder project and traverse up the hierarchy to build the full relative path.
            var current = solutionFolderProject;
            // Continue traversing up the hierarchy as long as the current project is a solution folder and has a parent project item.
            while (current != null && IsSolutionFolder(current))
            {
                // Add the current solution folder's name to the beginning of the segments list to build the path in the correct order.
                segments.Insert(0, current.Name);
                // Get the parent project item of the current solution folder project to continue traversing up the hierarchy.
                var parentProjectItem = current.ParentProjectItem;
                // If there is no parent project item or the parent project is null, we have reached the top of the hierarchy and should stop traversing.
                if (parentProjectItem == null || parentProjectItem.ContainingProject == null)
                {
                    // We have reached the top of the hierarchy, so we can stop traversing.
                    break;
                }
                // Move up to the parent project for the next iteration of the loop.
                current = parentProjectItem.ContainingProject;
            }
            // Join the segments using the system's directory separator character to form the final relative path string and return it.
            return string.Join(Path.DirectorySeparatorChar.ToString(), segments.ToArray());
        }

        // Checks if any installed .NET SDK supports netstandard2.1 by inspecting the SDK installation directory and parsing version numbers.
        private static bool IsNetStandard21Available()
        {
            // Check for installed .NET SDKs that support netstandard2.1 by inspecting the standard SDK installation directory.
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);

            // The .NET SDKs are typically installed under Program Files\dotnet\sdk on Windows.
            var sdkRoot = Path.Combine(programFiles, "dotnet", "sdk");

            // If the SDK directory does not exist, netstandard2.1 cannot be supported.
            if (!Directory.Exists(sdkRoot))
                return false;

            // Iterate through each SDK folder to check its version. The folder names usually start with a version number, such as "3.1.100" or "5.0.200".
            foreach (var dir in Directory.GetDirectories(sdkRoot))
            {
                // Extract the folder name to parse the version.
                var name = Path.GetFileName(dir);

                // Ignore any preview or suffixed versions that cannot be parsed after removing the suffix.
                if (!Version.TryParse(name.Split('-')[0], out var version))
                    continue;

                // .NET Core 3.x SDK and newer can build netstandard2.1 libraries.
                if (version.Major >= 3)
                    return true;
            }
            // If no suitable SDKs were found, netstandard2.1 is not available.
            return false;
        }

        // Determines whether the specified project is a solution folder by comparing its kind to the known solution folder GUID.
        private static bool IsSolutionFolder(Project project) => string.Equals(project?.Kind, EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder, StringComparison.OrdinalIgnoreCase);

        // Normalizes a namespace string by trimming whitespace and dots, splitting into segments, and ensuring each segment is a valid C# identifier.
        private static string NormalizeNamespace(string value)
        {
            // Validate input to ensure a meaningful namespace can be constructed.
            var normalized = value.Trim().Trim('.');

            // Split the namespace into segments by dots, removing any empty segments and normalizing each segment to ensure it is a valid C# identifier.
            var segments = normalized.Split(['.'], StringSplitOptions.RemoveEmptyEntries).Select(NormalizeNamespaceSegment).ToArray();

            // Ensure that at least one valid segment exists after normalization; otherwise, throw an exception to indicate that the namespace is invalid.
            if (segments.Length == 0)
                throw new ArgumentException("Namespace must contain at least one valid segment.", nameof(value));

            // Recombine the normalized segments into a single namespace string, using dots as separators, which is suitable for use in C# code.
            return string.Join(".", segments);
        }

        // Normalizes a single segment of a namespace by removing invalid characters and ensuring it starts with a letter or underscore, which is required for valid C# identifiers.
        private static string NormalizeNamespaceSegment(string segment)
        {
            // If the segment is empty or whitespace, return a safe placeholder to avoid invalid namespace segments.
            if (string.IsNullOrWhiteSpace(segment))
                return "_";
            // Remove characters that are not letters, digits, or underscores to ensure the segment is a valid C# identifier.
            var chars = segment.Where(ch => char.IsLetterOrDigit(ch) || ch == '_').ToArray();
            // If the resulting segment is empty after filtering, return a safe placeholder to avoid invalid namespace segments.
            var result = new string(chars);
            // Ensure the first character is a letter or underscore, as required for valid C# identifiers. If it is not, prepend an underscore to make it valid.
            if (string.IsNullOrWhiteSpace(result))
                return "_";
            // If the first character is not a letter or underscore, prepend an underscore to make it a valid identifier.
            if (!char.IsLetter(result[0]) && result[0] != '_')
                result = "_" + result;
            // Return the normalized segment, which is guaranteed to be a valid C# identifier for use in namespaces.
            return result;
        }

        // Normalizes a version string into a sortable format by padding major and minor version numbers with leading zeros, ensuring consistent lexicographical sorting.
        private static string NormalizeVersionForSort(string versionText)
        {
            // If the version string is empty or whitespace, return a default sort key that represents the lowest possible version.
            if (string.IsNullOrWhiteSpace(versionText))
                return "0000_0000";
            // Split the version string into parts using the dot as a separator, ignoring empty entries to handle cases like "1..0".
            var parts = versionText.Split(['.'], StringSplitOptions.RemoveEmptyEntries);
            // Initialize major and minor version numbers to zero, which will be used for sorting if the version string does not contain valid numbers.
            int major = 0;
            int minor = 0;
            // Attempt to parse the major version number from the first part of the split version string, if it exists.
            if (parts.Length > 0)
                int.TryParse(parts[0], out major);
            // Attempt to parse the minor version number from the second part of the split version string, if it exists.
            if (parts.Length > 1)
                int.TryParse(parts[1], out minor);
            // Return a formatted string with leading zeros for both major and minor versions, ensuring that the sort key is consistent and sortable lexicographically.
            return major.ToString("D4") + "_" + minor.ToString("D4");
        }

        // Generates a sort key for target frameworks to ensure consistent ordering in the UI, prioritizing .NET Standard, then .NET Framework, and finally .NET Core/5+.
        private static string TargetFrameworkSortKey(string tfm)
        {
            // Use a prefix to ensure .NET Standard comes first, then .NET Framework, then .NET Core/5+, and finally any unknown frameworks.
            if (string.IsNullOrWhiteSpace(tfm))
                return "9_";
            // Normalize the target framework string to ensure consistent comparison and sorting.
            if (tfm.StartsWith("netstandard", StringComparison.OrdinalIgnoreCase))
                return "1_" + NormalizeVersionForSort(tfm.Substring("netstandard".Length));
            // .NET Framework target frameworks start with "net" followed by a version number without a dot, e.g., "net48".
            if (tfm.StartsWith("net", StringComparison.OrdinalIgnoreCase) && tfm.Contains("."))
                return "2_" + NormalizeVersionForSort(tfm.Substring("net".Length));
            // .NET Framework target frameworks without a dot, e.g., "net48", are sorted after .NET Standard but before .NET Core/5+.
            if (tfm.StartsWith("net", StringComparison.OrdinalIgnoreCase))
                return "3_" + NormalizeVersionForSort(tfm.Substring("net".Length));
            // .NET Core and .NET 5+ target frameworks start with "net" followed by a version number with a dot, e.g., "net5.0", "net6.0", "net7.0".
            return "9_" + tfm;
        }

        // Converts a relative path (e.g., "Folder1/Folder2") into a namespace-like suffix (e.g., "Folder1.Folder2") by normalizing each segment for namespace compatibility.
        private static string ToNamespaceSuffix(string? relativePath)
        {
            // Ensure this method is called on the UI thread because it may be used in conjunction with DTE project properties.
            if (string.IsNullOrWhiteSpace(relativePath))
                return string.Empty;
            // Normalize the relative path by replacing backslashes with forward slashes, splitting it into segments, normalizing each segment for namespace compatibility, and filtering out any empty or whitespace segments.
            var segments = relativePath?.Replace('\\', '/').Split(['/'], StringSplitOptions.RemoveEmptyEntries).Select(NormalizeNamespaceSegment).Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
            // If there are no valid segments, return an empty string to indicate that there is no namespace suffix.
            return string.Join(".", segments);
        }

        // Converts an arbitrary string into a safe path segment by removing invalid characters and trimming whitespace.
        private static string ToSafePathSegment(string segment)
        {
            // Validate input to ensure a meaningful path segment can be constructed.
            var invalid = Path.GetInvalidFileNameChars();
            // Remove any characters that are not valid in file or directory names to ensure the segment can be used safely in the file system.
            var filtered = new string([.. segment.Where(ch => !invalid.Contains(ch))]).Trim();
            // If the resulting segment is empty or whitespace after filtering, return a safe placeholder to avoid invalid path segments.
            return string.IsNullOrWhiteSpace(filtered) ? "_" : filtered;
        }

        // Tries to delete a directory and its contents, ignoring any exceptions that may occur (e.g., if files are locked by Visual Studio).
        private static void TryDeleteDirectory(string? directoryPath)
        {
            // If the directory path is null, empty, or whitespace, there is nothing to delete, so return early.
            if (string.IsNullOrWhiteSpace(directoryPath))
                return;
            // Attempt to delete the directory and its contents recursively, but catch and ignore any exceptions that may occur during deletion.
            try
            {
                // Check if the directory exists before attempting to delete it to avoid unnecessary exceptions.
                if (Directory.Exists(directoryPath))
                    // Delete the directory and all its contents recursively.
                    Directory.Delete(directoryPath, recursive: true);
            }
            catch
            {
                // Ignore cleanup errors.
                // Visual Studio or the project system may still hold temporary files briefly.
            }
        }

        // Recursively searches for a solution folder project that matches the specified relative path within the solution.
        private static bool TryFindSolutionFolderProject(Solution solution, string? relativeSolutionFolderPath, out Project solutionFolderProject)
        {
            // Ensure this method is called on the UI thread because it accesses DTE project properties.
            ThreadHelper.ThrowIfNotOnUIThread();
            // Initialize the out parameter to null to indicate that no solution folder project has been found yet.
            solutionFolderProject = null!;
            // Split the relative solution folder path into segments, normalizing them to ensure they are valid and non-empty for comparison against project names.
            var segments = relativeSolutionFolderPath?.Replace('\\', '/').Split(['/'], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
            // If there are no valid segments, the path is invalid and cannot match any solution folder.
            if (segments == null || segments.Length == 0)
                // If the relative solution folder path is null or empty, there is no solution folder to find, so return false.
                return false;
            // Iterate through each top-level project in the solution to find a matching solution folder project based on the provided path segments.
            foreach (Project project in solution.Projects)
            {
                // Recursively check if the current project or any of its subprojects match the desired solution folder path. If a match is found, return true and set the out parameter.
                if (TryFindSolutionFolderProject(project, segments, 0, out solutionFolderProject))
                    // If a matching solution folder project is found, return true to indicate success.
                    return true;
            }
            // If no matching solution folder project was found after checking all top-level projects, return false to indicate that the desired solution folder path does not exist in the solution.
            return false;
        }

        // Recursively searches for a solution folder project that matches the specified path segments.
        private static bool TryFindSolutionFolderProject(Project project, string[]? segments, int index, out Project solutionFolderProject)
        {
            // Ensure this method is called on the UI thread because it accesses DTE project properties.
            ThreadHelper.ThrowIfNotOnUIThread();
            // Initialize the out parameter to null to indicate that no solution folder project has been found yet.
            solutionFolderProject = null!;
            // Check if the current project is a solution folder. If not, it cannot match the desired path.
            if (!IsSolutionFolder(project))
                return false;
            // If the segments array is null or the index is out of bounds, the path is invalid and cannot match any solution folder.
            if (segments == null || index >= segments.Length)
                return false;
            // Compare the current project's name with the expected segment at the current index. If they don't match, this path is not valid.
            if (!string.Equals(project.Name, segments[index], StringComparison.OrdinalIgnoreCase))
                return false;
            // If this is the last segment in the path, we have found the matching solution folder project.
            if (index == segments.Length - 1)
            {
                solutionFolderProject = project;
                return true;
            }
            // If there are more segments to check, we need to look into the subprojects of the current solution folder.
            if (project.ProjectItems == null)
                return false;
            // Iterate through each project item in the current solution folder to find subprojects that may match the next segment in the path.
            foreach (ProjectItem item in project.ProjectItems)
            {
                // Check if the project item has a subproject. If it does not, skip to the next item.
                var subProject = item.SubProject;
                // If the subproject is null, it means this item does not represent a nested project, so we continue to the next item.
                if (subProject == null)
                    continue;
                // Recursively call this method to check if the subproject matches the next segment in the path. If a match is found, return true.
                if (TryFindSolutionFolderProject(subProject, segments, index + 1, out solutionFolderProject))
                    return true;
            }
            // If no matching subproject was found for the next segment, return false to indicate that the desired solution folder path does not exist.
            return false;
        }

        private static string? TryGetSelectedSolutionFolderRelativePath(EnvDTE80.DTE2 dte)
        {
            // Accessing the DTE and its properties must be done on the UI thread to avoid cross-thread operation exceptions.
            ThreadHelper.ThrowIfNotOnUIThread();
            // If any of the required DTE properties are null, we cannot determine the selected solution folder.
            if (dte.ToolWindows == null || dte.ToolWindows.SolutionExplorer == null || dte.ToolWindows.SolutionExplorer.SelectedItems == null)
            {
                // If the DTE or its properties are not available, we cannot determine the selected solution folder.
                return null;
            }

            // The SelectedItems property returns an array of selected items in Solution Explorer. If no items are selected, we cannot determine the solution folder.
            if (dte.ToolWindows.SolutionExplorer.SelectedItems is not Array selectedItems || selectedItems.Length == 0)
                // If no items are selected, there is no solution folder context to infer.
                return null;

            // The first selected item is typically the one we want to inspect for solution folder context.
            if (selectedItems.GetValue(0) is not UIHierarchyItem selectedItem)
                // If the selected item is not a UIHierarchyItem, we cannot determine the solution folder context.
                return null;
            // Initialize a variable to hold the selected project, which may be a solution folder or a regular project.
            Project? selectedProject = null;

            // Case 1: Visual Studio delivers a Project directly when a solution folder is selected in Solution Explorer.
            selectedProject = selectedItem.Object as Project;

            // Case 2: If the selected item is a ProjectItem, it may represent a subproject within a solution folder. In that case, we can access the SubProject property to get the actual project.
            if (selectedProject == null)
            {
                // If the selected item is a ProjectItem, check if it has a SubProject, which would indicate that it is part of a solution folder.
                if (selectedItem.Object is ProjectItem selectedProjectItem)
                    selectedProject = selectedProjectItem.SubProject;
            }
            // If we still don't have a selected project, it means the selection does not correspond to a solution folder or a project item with a subproject.
            if (selectedProject == null)
                // If no project is selected, we cannot determine the solution folder context.
                return null;
            // Check if the selected project is a solution folder. If it is not, we cannot determine a relative path for it.
            if (!IsSolutionFolder(selectedProject))
                // If the selected project is not a solution folder, we cannot determine a relative path for it.
                return null;
            // If we have a valid solution folder project, we can compute its relative path within the solution by traversing its parent hierarchy.
            return GetSolutionFolderRelativePath(selectedProject);
        }

        // Waits asynchronously for a .csproj file to appear in the specified project directory, which indicates that Visual Studio has completed project creation from the template.
        private static async Task<string> WaitForCsprojAsync(string projectDir, CancellationToken ct)
        {
            // Poll briefly because Visual Studio template creation may write the project file asynchronously.
            for (int i = 0; i < 30; i++)
            {
                // Respect cancellation while waiting for the project file to appear.
                ct.ThrowIfCancellationRequested();

                // Look for the generated project file in the root of the target project directory.
                var csproj = Directory.GetFiles(projectDir, "*.csproj", SearchOption.TopDirectoryOnly).FirstOrDefault();

                // Return as soon as the project file is available.
                if (!string.IsNullOrWhiteSpace(csproj))
                    return csproj;

                // Wait a short interval before checking again.
                await Task.Delay(100, ct).ConfigureAwait(false);
            }

            // Fail explicitly if the template did not create a project file within the expected time.
            throw new FileNotFoundException("Project file (*.csproj) was not created from template.", projectDir);
        }

        // Retrieves the directory of the currently loaded solution in Visual Studio.
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