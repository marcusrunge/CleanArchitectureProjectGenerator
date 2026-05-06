namespace MarcusRunge.CleanArchitectureProjectGenerator.Constants
{
    /// <summary>
    /// Provides strongly centralized string constants used as button command parameters.
    /// </summary>
    /// <remarks>
    /// Using constants avoids "magic strings" scattered across the codebase and reduces
    /// the risk of typos when comparing command parameters (e.g., in ICommand handlers).
    /// <para>
    /// These values are typically passed as <c>CommandParameter</c> from the UI and then
    /// interpreted in the ViewModel/command logic to decide which action to execute.
    /// </para>
    /// </remarks>
    public class ButtonCommandParameters
    {
        /// <summary>
        /// Command parameter value representing a "Cancel" action.
        /// </summary>
        /// <remarks>
        /// Use this when the user intends to abort/close a dialog, wizard, or operation.
        /// </remarks>
        public const string Cancel = "Cancel";

        /// <summary>
        /// Command parameter value representing a "Create" action.
        /// </summary>
        /// <remarks>
        /// Use this when the user confirms creation (e.g., creating a project or scaffold).
        /// </remarks>
        public const string Create = "Create";
    }
}