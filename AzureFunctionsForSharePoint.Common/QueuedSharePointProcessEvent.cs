namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Service Bus queue message sent by EventDispatch in response to receiving a remote event notification from SharePoint
    /// </summary>
    public class QueuedSharePointProcessEvent : QueuedSharePointEvent
    {
        /// <summary>
        /// Details about the remote event
        /// </summary>
        public SharePointRemoteEventAdapter SharePointRemoteEventAdapter { get; set; }
    }
}
