using System;

namespace FunctionsCore
{
    public delegate void FunctionNotificationEventHandler(
        object sender, FunctionNotificationEventArgs eventArgs);

    public class FunctionNotificationEventArgs : EventArgs
    {
        public string Message { get; set; }
    }
}
