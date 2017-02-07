using Microsoft.SharePoint.Client;

namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Base properties for queued service bus messages
    /// </summary>
    public class QueuedSharePointEvent
    {
        /// <summary>
        /// The id of the client. This must match a previously created client configuration in Azure storage. 
        /// </summary>
        public string ClientId { get; set; }
        /// <summary>
        /// The URL of the SharePoint <see cref="Web"/> to which the message applies
        /// </summary>
        public string AppWebUrl { get; set; }
        /// <summary>
        /// Security access token for the app-only identity
        /// </summary>
        public string AppAccessToken { get; set; }
        /// <summary>
        /// Security access token for the user's identity
        /// </summary>
        public string UserAccessToken { get; set; }
        /// <summary>
        /// Number of times processing this message has failed previously
        /// </summary>
        public int RetryCount { get; set; }
    }
}
