namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Service Bus queue message sent by AppLaunch to notify the client a user launched an app
    /// </summary>
    /// <remarks>
    /// The messaging system uses a object's type as a message's content type to allow clients to distinguish different message types in the same queue. 
    /// Subclassing the base QueuedSharePointEvent provides this type information.
    /// </remarks>
    public class QueuedAppLaunchEvent : QueuedSharePointEvent
    {
    }
}
