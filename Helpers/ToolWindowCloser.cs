using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Helpers
{
    internal static class ToolWindowCloser
    {
        public static readonly DependencyProperty CloseRequestedProperty = DependencyProperty.RegisterAttached("CloseRequested", typeof(bool), typeof(ToolWindowCloser), new PropertyMetadata(false, OnCloseRequestedChanged));

        public static readonly DependencyProperty FrameProperty = DependencyProperty.RegisterAttached("Frame", typeof(IVsWindowFrame), typeof(ToolWindowCloser), new PropertyMetadata(null));

        public static bool GetCloseRequested(DependencyObject obj) => (bool)obj.GetValue(CloseRequestedProperty);

        public static IVsWindowFrame GetFrame(DependencyObject obj) => (IVsWindowFrame)obj.GetValue(FrameProperty);

        public static void SetCloseRequested(DependencyObject obj, bool value) => obj.SetValue(CloseRequestedProperty, value);

        public static void SetFrame(DependencyObject obj, IVsWindowFrame value) => obj.SetValue(FrameProperty, value);

        private static void OnCloseRequestedChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue is true)
            {
                var frame = GetFrame(d);
                if (frame == null)
                    return;

                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    frame.CloseFrame(
                        (uint)__FRAMECLOSE.FRAMECLOSE_NoSave);
                }).FileAndForget("ToolWindowCloser/CloseFrame");

                d.SetValue(CloseRequestedProperty, false);
            }
        }
    }
}