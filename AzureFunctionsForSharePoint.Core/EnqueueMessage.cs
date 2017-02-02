using System;
using System.Runtime.Serialization;
using System.Web.Script.Serialization;
using AzureFunctionsForSharePoint.Common;
using Microsoft.ServiceBus;
using Microsoft.ServiceBus.Messaging;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;

namespace AzureFunctionsForSharePoint.Core
{
    /// <summary>
    /// Sends a message to a client's service bus queue
    /// </summary>
    public class EnqueueMessage
    {
        /// <summary>
        /// Send the event data as serialized json to the service bus queue specified in the client config
        /// </summary>
        /// <param name="eventData">The message to send</param>
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

            //It has been observed that it is possible for the message to get to the handler
            //before SP is done processing synchronous events like ItemUpdating 
            //causing handlers that fetch items to get old (un-updated) values for list items.
            //
            //Wait a little bit before the message gets delivered to the handler so SP can get its work done
            message.ScheduledEnqueueTimeUtc = DateTime.Now.AddSeconds(15);
            client.Send(message);
        }

        /// <summary>
        /// General purpose method for sending arbitrary messages to a service bus queue
        /// </summary>
        /// <param name="eventData">The message to send</param>
        /// <param name="serviceBusConnectionString">The message to send</param>
        /// <param name="notificationQueueName">The message to send</param>
        /// <param name="delaySeconds">The message to send</param>
        public static void SendQueueMessage(object eventData, string serviceBusConnectionString, string notificationQueueName, int delaySeconds = 0)
        {
            if (string.IsNullOrEmpty(serviceBusConnectionString) || string.IsNullOrEmpty(notificationQueueName))
            {
                throw new ArgumentException("Sending a message requires a connection and a queue name");
            };

            QueueDescription qd = new QueueDescription(notificationQueueName)
            {
                MaxSizeInMegabytes = 5120,
                DefaultMessageTimeToLive = new TimeSpan(5, 0, 0, 0)
            };

            string connectionString = serviceBusConnectionString;

            var namespaceManager = NamespaceManager.CreateFromConnectionString(connectionString);

            if (!namespaceManager.QueueExists(notificationQueueName))
            {
                namespaceManager.CreateQueue(qd);
            }

            var client = QueueClient.CreateFromConnectionString(connectionString, notificationQueueName);
            BrokeredMessage message = new BrokeredMessage(ToJSON(eventData), new DataContractSerializer(typeof(string)));
            message.ContentType = eventData.GetType().ToString();

            //Delay the delivery as instructed
            message.ScheduledEnqueueTimeUtc = DateTime.Now.AddSeconds(delaySeconds);
            client.Send(message);
        }

        private static string ToJSON(Object e)
        {
            JavaScriptSerializer js = new JavaScriptSerializer();
            return js.Serialize(e);
        }
    }
}
