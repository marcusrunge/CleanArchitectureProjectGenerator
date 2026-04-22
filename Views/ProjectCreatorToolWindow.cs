using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Views
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("5892c607-dab4-446d-be29-eb33671c0475")]
    public class ProjectCreatorToolWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectCreatorToolWindow"/> class.
        /// </summary>
        public ProjectCreatorToolWindow() : base(null)
        {
            this.Caption = "Clean Architecture Project Generator";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new ProjectCreatorToolWindowControl();
        }
    }
}
