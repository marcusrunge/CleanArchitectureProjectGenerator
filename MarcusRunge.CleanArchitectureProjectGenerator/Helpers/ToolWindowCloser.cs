using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows;
using System.Windows.Data;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Helpers
{
    /// <summary>
    /// Provides attached properties to allow a WPF control to request closing its hosting
    /// Visual Studio tool window without directly referencing VS shell APIs.
    /// </summary>
    /// <remarks>
    /// This helper enables an MVVM-friendly "close request" pattern:
    /// <list type="bullet">
    /// <item><description>The view (tool window control) stores an <see cref="IVsWindowFrame"/> via an attached property.</description></item>
    /// <item><description>The ViewModel can trigger closing by setting <c>CloseRequested</c> to <c>true</c> (typically via binding).</description></item>
    /// <item><description>When the flag flips to true, the helper closes the frame on the UI thread and resets the flag.</description></item>
    /// </list>
    /// This avoids passing shell objects into the ViewModel and keeps VS-specific code in the UI layer.
    /// </remarks>
    internal static class ToolWindowCloser
    {
        /// <summary>
        /// Attached property used as a boolean "signal" indicating that the tool window should close.
        /// </summary>
        /// <remarks>
        /// When set to <c>true</c>, <see cref="OnCloseRequestedChanged"/> is invoked.
        /// The property is reset back to <c>false</c> after processing to allow repeated close requests.
        /// </remarks>
        public static readonly DependencyProperty CloseRequestedProperty =
            DependencyProperty.RegisterAttached(
                "CloseRequested",
                typeof(bool),
                typeof(ToolWindowCloser),
                new PropertyMetadata(false, OnCloseRequestedChanged));

        /// <summary>
        /// Attached property used to store the <see cref="IVsWindowFrame"/> instance associated with a tool window.
        /// </summary>
        /// <remarks>
        /// The frame is the VS shell object that provides the actual close operation (<see cref="IVsWindowFrame.CloseFrame(uint)"/>).
        /// It is stored on the control so that closing can be triggered indirectly via <see cref="CloseRequestedProperty"/>.
        /// </remarks>
        public static readonly DependencyProperty FrameProperty =
            DependencyProperty.RegisterAttached(
                "Frame",
                typeof(IVsWindowFrame),
                typeof(ToolWindowCloser),
                new PropertyMetadata(null));

        /// <summary>
        /// Gets the current value of the <see cref="CloseRequestedProperty"/> attached property.
        /// </summary>
        /// <param name="obj">The dependency object that stores the attached property value.</param>
        /// <returns><c>true</c> if a close was requested; otherwise <c>false</c>.</returns>
        public static bool GetCloseRequested(DependencyObject obj) =>
            (bool)obj.GetValue(CloseRequestedProperty);

        /// <summary>
        /// Gets the <see cref="IVsWindowFrame"/> stored on the specified dependency object.
        /// </summary>
        /// <param name="obj">The dependency object that stores the attached property value.</param>
        /// <returns>The associated <see cref="IVsWindowFrame"/>.</returns>
        /// <remarks>
        /// This method will return <c>null</c> if no frame was assigned.
        /// </remarks>
        public static IVsWindowFrame GetFrame(DependencyObject obj) =>
            (IVsWindowFrame)obj.GetValue(FrameProperty);

        /// <summary>
        /// Sets the <see cref="CloseRequestedProperty"/> attached property.
        /// </summary>
        /// <param name="obj">The dependency object that stores the attached property value.</param>
        /// <param name="value">The value indicating whether closing is requested.</param>
        public static void SetCloseRequested(DependencyObject obj, bool value) =>
            obj.SetValue(CloseRequestedProperty, value);

        /// <summary>
        /// Sets the <see cref="FrameProperty"/> attached property.
        /// </summary>
        /// <param name="obj">The dependency object that stores the attached property value.</param>
        /// <param name="value">The <see cref="IVsWindowFrame"/> to store.</param>
        public static void SetFrame(DependencyObject obj, IVsWindowFrame value) =>
            obj.SetValue(FrameProperty, value);

        /// <summary>
        /// Callback invoked when <see cref="CloseRequestedProperty"/> changes.
        /// </summary>
        /// <param name="d">The dependency object that owns the property.</param>
        /// <param name="e">Change details (old/new values).</param>
        /// <remarks>
        /// This method performs the actual tool window close operation when the flag becomes <c>true</c>.
        /// </remarks>
        private static void OnCloseRequestedChanged(
    DependencyObject d,
    DependencyPropertyChangedEventArgs e)
        {
            // Only consume true as a close request.
            if (e.NewValue is not true)
                return;

            var frame = GetFrame(d);

            // Even if there is no frame, consume/reset the request.
            // Otherwise the flag may remain true forever.
            if (frame == null)
            {
                ResetCloseRequested(d);
                return;
            }

            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                try
                {
                    frame.CloseFrame((uint)__FRAMECLOSE.FRAMECLOSE_NoSave);
                }
                finally
                {
                    // Reset only after CloseFrame has been called.
                    ResetCloseRequested(d);
                }
            })
            .FileAndForget("ToolWindowCloser/CloseFrame");
        }

        private static void ResetCloseRequested(DependencyObject d)
        {
            // Prefer SetCurrentValue so an existing binding is not replaced.
            d.SetCurrentValue(CloseRequestedProperty, false);

            // If CloseRequested is bound TwoWay, push the reset back to the source VM property.
            var binding = BindingOperations.GetBindingExpression(d, CloseRequestedProperty);
            binding?.UpdateSource();
        }
    }
}