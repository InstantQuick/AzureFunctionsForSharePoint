using System;
using System.Runtime.Serialization;
using System.Web.Script.Serialization;
using AzureFunctionsForSharePoint.Common;
using Microsoft.ServiceBus.Messaging;
using Microsoft.SharePoint.Client;
using static AzureFunctionsForSharePoint.Core.Security.TokenHelper;

namespace IQAppBackgroundJobs
{
    public class BackgroundJobHandlerArgs : AzureFunctionArgs { }
    public class BackgroundJobHandler : FunctionBase
    {
        public void Execute(BrokeredMessage receivedMessage, BackgroundJobHandlerArgs storageConfig)
        {
            var eventJSON = receivedMessage.GetBody<string>(new DataContractSerializer(typeof(string)));
            var baseEvent = (new JavaScriptSerializer()).Deserialize<QueuedSharePointEvent>(eventJSON);

            var appOnlyContext = GetClientContext(baseEvent.AppWebUrl,
                baseEvent.AppAccessToken);

            appOnlyContext.Web.EnsureProperties(w => w.Url, w => w.ServerRelativeUrl, w => w.AllProperties);
            Log($"Connected to {appOnlyContext.Web.Url}");

            try
            {
                switch (receivedMessage.ContentType)
                {
                    case "AzureFunctionsForSharePoint.Common.QueuedAppLaunchEvent":
                        Log(receivedMessage.ContentType);
                        break;
                    case "AzureFunctionsForSharePoint.Common.QueuedSharePointProvisioningEvent":
                        Log(receivedMessage.ContentType);
                        break;
                    case "AzureFunctionsForSharePoint.Common.QueuedSharePointProcessEvent":
                        Log(receivedMessage.ContentType);
                        var actualEvent = (new JavaScriptSerializer()).Deserialize<QueuedSharePointProcessEvent>(eventJSON);

                        foreach (var prop in actualEvent.SharePointRemoteEventAdapter.EventProperties)
                        {
                            Log($"Event {prop.Key}={prop.Value}");
                        }
                        foreach (var prop in actualEvent.SharePointRemoteEventAdapter.ItemAfterProperties)
                        {
                            Log($"After {prop.Key}={prop.Value}");
                        }
                        foreach (var prop in actualEvent.SharePointRemoteEventAdapter.ItemBeforeProperties)
                        {
                            Log($"Before {prop.Key}={prop.Value}");
                        }
                        Log($"Finished {actualEvent.SharePointRemoteEventAdapter.EventType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                Log($"{ex}");
            }
        }
    }
}
