using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MarcusRunge.CleanArchitectureProjectGenerator.Common
{
    /// <summary>
    /// Base class for ViewModels (or other bindable objects) that provides
    /// <see cref="INotifyPropertyChanged"/> support.
    /// </summary>
    /// <remarks>
    /// Implements the standard pattern used in MVVM to notify the UI when a property value changes.
    /// <para>
    /// Key goals:
    /// <list type="bullet">
    /// <item><description>Reduce boilerplate by centralizing <see cref="PropertyChanged"/> raising.</description></item>
    /// <item><description>Avoid redundant notifications by checking equality before updating values.</description></item>
    /// <item><description>Automatically infer the property name via <see cref="CallerMemberNameAttribute"/>.</description></item>
    /// </list>
    /// </para>
    /// </remarks>
    internal abstract class BindableBase : INotifyPropertyChanged
    {
        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        /// <remarks>
        /// UI frameworks (e.g., WPF) listen to this event to refresh bindings when a property changes.
        /// </remarks>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// Raises the <see cref="PropertyChanged"/> event for the specified property.
        /// </summary>
        /// <param name="propertyName">
        /// The name of the property that changed. When omitted, the compiler supplies the caller member name.
        /// </param>
        /// <remarks>
        /// Marked as <c>virtual</c> to allow derived classes to intercept or extend notification behavior
        /// (e.g., raise dependent property notifications, add logging, etc.).
        /// </remarks>
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            // Use null-conditional invocation to safely raise the event only if there are subscribers.
            // PropertyChangedEventArgs carries the property name so the UI can update only affected bindings.
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Sets a property's backing field and raises <see cref="PropertyChanged"/> only if the value actually changed.
        /// </summary>
        /// <typeparam name="T">The property type.</typeparam>
        /// <param name="backingField">A reference to the field that stores the property's current value.</param>
        /// <param name="value">The new value to assign.</param>
        /// <param name="propertyName">
        /// The name of the property that changed. When omitted, the compiler supplies the caller member name.
        /// </param>
        /// <returns>
        /// <c>true</c> if the value was changed and a notification was raised; otherwise <c>false</c>.
        /// </returns>
        /// <remarks>
        /// This helper prevents unnecessary UI refreshes by:
        /// <list type="bullet">
        /// <item><description>Comparing old/new values using <see cref="EqualityComparer{T}.Default"/>.</description></item>
        /// <item><description>Only assigning and raising <see cref="PropertyChanged"/> when a real change occurs.</description></item>
        /// </list>
        /// The boolean return value is useful for conditional follow-up logic (e.g., recompute derived state only when changed).
        /// </remarks>
        protected bool SetProperty<T>(ref T backingField, T value, [CallerMemberName] string? propertyName = null)
        {
            // Use the default equality comparer for T to correctly handle:
            // - value types
            // - reference types
            // - custom equality implementations (IEquatable<T>, overridden Equals, etc.)
            if (EqualityComparer<T>.Default.Equals(backingField, value))
                return false;

            // Update the storage field first so any observers reading the property after notification see the new value.
            backingField = value;

            // Notify bindings/listeners that the property value changed.
            OnPropertyChanged(propertyName);

            return true;
        }
    }
}