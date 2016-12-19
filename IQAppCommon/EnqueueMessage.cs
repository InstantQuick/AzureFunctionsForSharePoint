using System;
using System.Runtime.Serialization;
using System.Web.Script.Serialization;
using AzureFunctionsForSharePoint.Common;
using Microsoft.ServiceBus;
using Microsoft.ServiceBus.Messaging;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;

namespace AzureFunctionsForSharePoint.Core
{
    public class EnqueueMessage
    {
        public static void SendQueueMessage(QueuedSharePointEvent eventData)
        {
            var clientConfig = GetConfiguration(eventData.ClientId);

            if (string.IsNullOrEmpty(clientConfig.NotificationQueueName) || string.IsNullOrEmpty(clientConfig.ServiceBusConnectionString)) return;

            QueueDescription qd = new QueueDescription(clientConfig.NotificationQueueName)
            {
                MaxSizeInMegabytes = 5120,
                DefaultMessageTimeToLive = new TimeSpan(5, 0, 0, 0)
            };

            string connectionString = clientConfig.ServiceBusConnectionString;

            var namespaceManager = NamespaceManager.CreateFromConnectionString(connectionString);

            if (!namespaceManager.QueueExists(clientConfig.NotificationQueueName))
            {
                namespaceManager.CreateQueue(qd);
            }

            var client = QueueClient.CreateFromConnectionString(connectionString, clientConfig.NotificationQueueName);
            BrokeredMessage message = new BrokeredMessage(ToJSON(eventData), new DataContractSerializer(typeof(string)));
            message.ContentType = eventData.GetType().ToString();
            client.Send(message);
        }

        private static string ToJSON(Object e)
        {
            JavaScriptSerializer js = new JavaScriptSerializer();
            return js.Serialize(e);
        }
    }
}
