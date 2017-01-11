# Understanding the EventDispatch Function
The EventDispatch function receives a remote event from SharePoint as a WCF SOAP message and 
parses it using [SharePointRemoteEventAdapter](../api/AzureFunctionsForSharePoint.Common.SharePointRemoteEventAdapter.html). 
Based on the event type, the received information may be augmented by reading additional information 
from SharePoint. EventDispatch sends the resulting [QueuedSharePointProcessEvent](../api/AzureFunctionsForSharePoint.Common.SharePointRemoteEventAdapter.html) 
to the client's service bus queue as JSON.
