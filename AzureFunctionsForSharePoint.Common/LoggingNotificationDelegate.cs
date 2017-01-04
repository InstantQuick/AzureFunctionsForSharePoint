using System;

namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Delegate backing the <see cref="FunctionBase.FunctionNotify"/> event
    /// </summary>
    /// <param name="sender">The function handler</param>
    /// <param name="eventArgs">The event data</param>
    public delegate void FunctionNotificationEventHandler(
        object sender, FunctionNotificationEventArgs eventArgs);

    /// <summary>
    /// A message to send to <see cref="FunctionNotificationEventHandler"/>
    /// </summary>
    public class FunctionNotificationEventArgs : EventArgs
    {
        /// <summary>
        /// The message text
        /// </summary>
        public string Message { get; set; }
    }
}
