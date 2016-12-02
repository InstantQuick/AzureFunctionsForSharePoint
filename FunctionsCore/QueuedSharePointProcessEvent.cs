using System.Collections.Generic;

namespace FunctionsCore
{
    public class QueuedSharePointProcessEvent : QueuedSharePointEvent
    {
        public SharePointRemoteEventAdapter SharePointRemoteEventAdapter { get; set; }
    }
}
