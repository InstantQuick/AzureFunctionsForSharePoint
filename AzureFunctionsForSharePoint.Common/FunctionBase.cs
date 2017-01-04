namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Functionality common to individual functions
    /// </summary>
    public class FunctionBase
    {
        /// <summary>
        /// A simple notification event. Listeners to this event can log messages as needed.
        /// </summary>
        /// <remarks>
        /// If you need more than just the simple Message property, extend <see cref="FunctionNotificationEventArgs"/>
        /// </remarks>
        public event FunctionNotificationEventHandler FunctionNotify;

        /// <summary>
        /// Fires the FunctionNotify event with a message
        /// </summary>
        /// <param name="message">The notification message</param>
        public void Log(string message)
        {
            FunctionNotify?.Invoke(this, new FunctionNotificationEventArgs { Message = message });
        }
    }
}
