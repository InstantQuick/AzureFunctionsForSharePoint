namespace AzureFunctionsForSharePoint.Common
{
    public class FunctionBase
    {
        public event FunctionNotificationEventHandler FunctionNotify;

        public void Log(string message)
        {
            FunctionNotify?.Invoke(this, new FunctionNotificationEventArgs { Message = message });
        }
    }
}
