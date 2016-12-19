namespace AzureFunctionsForSharePoint.Common
{
    public class QueuedSharePointProcessEvent : QueuedSharePointEvent
    {
        public SharePointRemoteEventAdapter SharePointRemoteEventAdapter { get; set; }
    }
}
