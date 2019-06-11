using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using AzureFunctionsForSharePoint.Common.ProcessEvent;
using AzureFunctionsForSharePoint.Common.ProcessOneWayEvent;
using System.Xml.Linq;
using Microsoft.SharePoint.Client.EventReceivers;

namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Represents a remote SharePoint event notification parsed to be easy to handle with additional context to augment event processing.
    /// Sent by the EventDispatch function to a client's service bus queue in response to receipt of a remote event notification
    /// </summary>
    /// <seealso cref="SPRemoteEventProperties"/>
    public class SharePointRemoteEventAdapter
    {
        /// <summary>
        /// Corresponds to <see cref="SPRemoteEventProperties.CorrelationId"/>
        /// </summary>
        public string CorrelationId { get; set; }
        /// <summary>
        /// Corresponds to <see cref="SPRemoteEventProperties.CultureLCID"/>
        /// </summary>
        public string CultureLCID { get; set; }
        /// <summary>
        /// Corresponds to <see cref="SPRemoteEventProperties.ErrorCode"/>
        /// </summary>
        public string ErrorCode { get; set; }
        /// <summary>
        /// Corresponds to <see cref="SPRemoteEventProperties.ErrorMessage"/>
        /// </summary>
        public string ErrorMessage { get; set; }
        /// <summary>
        /// Corresponds to <see cref="SPRemoteEventProperties.EventType"/>
        /// </summary>
        public string EventType { get; set; }
        /// <summary>
        /// Corresponds to <see cref="SPRemoteEventProperties.UICultureLCID"/>
        /// </summary>
        public string UICultureLCID { get; set; }
        /// <summary>
        /// Name value pairs of the event data provided with the remote event. These vary by event types.
        /// </summary>
        public Dictionary<string, string> EventProperties = new Dictionary<string, string>();
        /// <summary>
        /// If applicable and available, the properties of a list item before the event. Event dispatch attempts to augment this data when possible.
        /// </summary>
        public Dictionary<string, string> ItemBeforeProperties = new Dictionary<string, string>();
        /// <summary>
        /// If applicable and available, the properties of a list item after the event
        /// </summary>
        public Dictionary<string, string> ItemAfterProperties = new Dictionary<string, string>();
        private string _contextToken = string.Empty;

        /// <summary>
        /// Provides the context token from the event
        /// </summary>
        /// <returns></returns>
        public string GetContextToken()
        {
            return _contextToken;
        }

        /// <summary>
        /// Transforms the text of a remote event WCF SOAP message into a new SharePointRemoteEventAdapter
        /// </summary>
        /// <param name="soapBody">Message received from SharePoint</param>
        /// <returns>A new SharePointRemoteEventAdapter instance from the SOAP message</returns>
        public static SharePointRemoteEventAdapter GetSharePointRemoteEventAdapter(string soapBody)
        {
            if (soapBody.Contains("ProcessOneWayEvent"))
            {
                ProcessOneWayEvent.Envelope soapEnvelope;
                using (var reader = new StringReader(soapBody))
                {
                    var serializer = new XmlSerializer(typeof(ProcessOneWayEvent.Envelope));
                    soapEnvelope = (ProcessOneWayEvent.Envelope)serializer.Deserialize(reader);
                }
                var processOneWayEventProperties = soapEnvelope.Items[0].ProcessOneWayEvent[0];
                var adapter = GetAdapterFromOneWayEventProperties(processOneWayEventProperties, soapBody);
                adapter._contextToken = processOneWayEventProperties.ContextToken;
                return adapter;
            }
            else if (soapBody.Contains("ProcessEvent"))
            {
                ProcessEvent.Envelope soapEnvelope;
                using (var reader = new StringReader(soapBody))
                {
                    var serializer = new XmlSerializer(typeof(ProcessEvent.Envelope));
                    soapEnvelope = (ProcessEvent.Envelope)serializer.Deserialize(reader);
                }
                var processEventProperties = soapEnvelope.Items[0].ProcessEvent[0];
                var adapter = GetAdapterFromEventProperties(processEventProperties, soapBody);
                adapter._contextToken = processEventProperties.ContextToken;
                return adapter;
            }
            return new SharePointRemoteEventAdapter();
        }

        private static SharePointRemoteEventAdapter GetAdapterFromEventProperties(ProcessEventProperties eventProperties, string soapBody)
        {
            var adapter = new SharePointRemoteEventAdapter()
            {
                EventType = eventProperties.EventType,
                CorrelationId = eventProperties.CorrelationId,
                CultureLCID = eventProperties.CultureLCID,
                ErrorCode = eventProperties.ErrorCode,
                ErrorMessage = eventProperties.ErrorMessage,
                UICultureLCID = eventProperties.UICultureLCID
            };

            if (eventProperties.AppEventProperties.Length > 0 &&
                eventProperties.AppEventProperties[0].ProductId != default(Guid))
            {
                adapter.EventProperties["AppWebFullUrl"] = eventProperties.AppEventProperties[0].AppWebFullUrl;
                adapter.EventProperties["AssetId"] = eventProperties.AppEventProperties[0].AssetId;
                adapter.EventProperties["ContentMarket"] = eventProperties.AppEventProperties[0].ContentMarket;
                adapter.EventProperties["HostWebFullUrl"] = eventProperties.AppEventProperties[0].HostWebFullUrl;
                adapter.EventProperties["ProductId"] = eventProperties.AppEventProperties[0].ProductId.ToString();
                adapter.EventProperties["PreviousVersion"] = eventProperties.AppEventProperties[0].PreviousVersion.ToString();
                adapter.EventProperties["Version"] = eventProperties.AppEventProperties[0].Version.ToString();
            }

            else if (eventProperties.EntityInstanceEventProperties.Length > 0 &&
                eventProperties.EntityInstanceEventProperties[0].LobSystemInstanceName != null)
            {
                adapter.EventProperties["LobSystemInstanceName"] = eventProperties.EntityInstanceEventProperties[0].LobSystemInstanceName;
                adapter.EventProperties["EntityName"] = eventProperties.EntityInstanceEventProperties[0].EntityName;
                adapter.EventProperties["EntityNamespace"] = eventProperties.EntityInstanceEventProperties[0].EntityNamespace;
                adapter.EventProperties["NotificationContext"] = eventProperties.EntityInstanceEventProperties[0].NotificationContext;
                adapter.EventProperties["NotificationMessage"] = Convert.ToBase64String(eventProperties.EntityInstanceEventProperties[0].NotificationMessage);
            }

            else if (eventProperties.ItemEventProperties.Length > 0 &&
                eventProperties.ItemEventProperties[0].ListItemId != null)
            {
                adapter.EventProperties["ListItemId"] = eventProperties.ItemEventProperties[0].ListItemId;
                adapter.EventProperties["AfterUrl"] = eventProperties.ItemEventProperties[0].AfterUrl;
                adapter.EventProperties["BeforeUrl"] = eventProperties.ItemEventProperties[0].BeforeUrl;
                adapter.EventProperties["CurrentUserId"] = eventProperties.ItemEventProperties[0].CurrentUserId;
                adapter.EventProperties["ExternalNotificationMessage"] = eventProperties.ItemEventProperties[0].ExternalNotificationMessage;
                adapter.EventProperties["IsBackgroundSave"] = eventProperties.ItemEventProperties[0].IsBackgroundSave;
                adapter.EventProperties["ListId"] = eventProperties.ItemEventProperties[0].ListId;
                adapter.EventProperties["ListTitle"] = eventProperties.ItemEventProperties[0].ListTitle;
                adapter.EventProperties["UserDisplayName"] = eventProperties.ItemEventProperties[0].UserDisplayName;
                adapter.EventProperties["UserLoginName"] = eventProperties.ItemEventProperties[0].UserLoginName;
                adapter.EventProperties["Versionless"] = eventProperties.ItemEventProperties[0].Versionless;
                adapter.EventProperties["WebUrl"] = eventProperties.ItemEventProperties[0].WebUrl;
                adapter = GetListItemBeforeAndAfterProperties(adapter, soapBody);
            }

            else if (eventProperties.ListEventProperties.Length > 0 &&
                eventProperties.ListEventProperties[0].ListId != null)
            {
                adapter.EventProperties["ListId"] = eventProperties.ListEventProperties[0].ListId;
                adapter.EventProperties["FeatureId"] = eventProperties.ListEventProperties[0].FeatureId;
                adapter.EventProperties["FieldName"] = eventProperties.ListEventProperties[0].FieldName;
                adapter.EventProperties["FieldXml"] = eventProperties.ListEventProperties[0].FieldXml;
                adapter.EventProperties["ListTitle"] = eventProperties.ListEventProperties[0].ListTitle;
                adapter.EventProperties["TemplateId"] = eventProperties.ListEventProperties[0].TemplateId;
                adapter.EventProperties["WebUrl"] = eventProperties.ListEventProperties[0].WebUrl;
            }

            else if (eventProperties.SecurityEventProperties.Length > 0 &&
                eventProperties.SecurityEventProperties[0].WebId != default(Guid))
            {
                adapter.EventProperties["UserLoginName"] = eventProperties.SecurityEventProperties[0].UserLoginName;
                adapter.EventProperties["GroupName"] = eventProperties.SecurityEventProperties[0].GroupName;
                adapter.EventProperties["GroupUserLoginName"] = eventProperties.SecurityEventProperties[0].GroupUserLoginName;
                adapter.EventProperties["RoleDefinitionName"] = eventProperties.SecurityEventProperties[0].RoleDefinitionName;
                adapter.EventProperties["ScopeUrl"] = eventProperties.SecurityEventProperties[0].ScopeUrl;
                adapter.EventProperties["UserDisplayName"] = eventProperties.SecurityEventProperties[0].UserDisplayName;
                adapter.EventProperties["WebFullUrl"] = eventProperties.SecurityEventProperties[0].WebFullUrl;
                adapter.EventProperties["WebId"] = eventProperties.SecurityEventProperties[0].WebId.ToString();
                adapter.EventProperties["GroupId"] = eventProperties.SecurityEventProperties[0].GroupId.ToString();
                adapter.EventProperties["GroupNewOwnerId"] = eventProperties.SecurityEventProperties[0].GroupNewOwnerId.ToString();
                adapter.EventProperties["GroupOwnerId"] = eventProperties.SecurityEventProperties[0].GroupOwnerId.ToString();
                adapter.EventProperties["GroupUserId"] = eventProperties.SecurityEventProperties[0].GroupUserId.ToString();
                adapter.EventProperties["ObjectType"] = eventProperties.SecurityEventProperties[0].ObjectType.ToString();
                adapter.EventProperties["PrincipalId"] = eventProperties.SecurityEventProperties[0].PrincipalId.ToString();
                adapter.EventProperties["RoleDefinitionId"] = eventProperties.SecurityEventProperties[0].RoleDefinitionId.ToString();
                adapter.EventProperties["RoleDefinitionPermissions"] = eventProperties.SecurityEventProperties[0].RoleDefinitionPermissions.ToString();
            }
            else if (eventProperties.WebEventProperties.Length > 0 &&
                eventProperties.WebEventProperties[0].FullUrl != null)
            {
                adapter.EventProperties["FullUrl"] = eventProperties.WebEventProperties[0].FullUrl;
                adapter.EventProperties["NewServerRelativeUrl"] = eventProperties.WebEventProperties[0].NewServerRelativeUrl;
                adapter.EventProperties["ServerRelativeUrl"] = eventProperties.WebEventProperties[0].ServerRelativeUrl;
            }
            return adapter;
        }
        private static SharePointRemoteEventAdapter GetAdapterFromOneWayEventProperties(ProcessOneWayEventProperties eventProperties, string soapBody)
        {
            var adapter = new SharePointRemoteEventAdapter()
            {
                EventType = eventProperties.EventType,
                CorrelationId = eventProperties.CorrelationId,
                CultureLCID = eventProperties.CultureLCID,
                ErrorCode = eventProperties.ErrorCode,
                ErrorMessage = eventProperties.ErrorMessage,
                UICultureLCID = eventProperties.UICultureLCID
            };

            if (eventProperties.AppEventProperties.Length > 0 &&
                eventProperties.AppEventProperties[0].ProductId != default(Guid))
            {
                adapter.EventProperties["AppWebFullUrl"] = eventProperties.AppEventProperties[0].AppWebFullUrl;
                adapter.EventProperties["AssetId"] = eventProperties.AppEventProperties[0].AssetId;
                adapter.EventProperties["ContentMarket"] = eventProperties.AppEventProperties[0].ContentMarket;
                adapter.EventProperties["HostWebFullUrl"] = eventProperties.AppEventProperties[0].HostWebFullUrl;
                adapter.EventProperties["ProductId"] = eventProperties.AppEventProperties[0].ProductId.ToString();
                adapter.EventProperties["PreviousVersion"] = eventProperties.AppEventProperties[0].PreviousVersion.ToString();
                adapter.EventProperties["Version"] = eventProperties.AppEventProperties[0].Version.ToString();
            }

            else if (eventProperties.EntityInstanceEventProperties.Length > 0 &&
                eventProperties.EntityInstanceEventProperties[0].LobSystemInstanceName != null)
            {
                adapter.EventProperties["LobSystemInstanceName"] = eventProperties.EntityInstanceEventProperties[0].LobSystemInstanceName;
                adapter.EventProperties["EntityName"] = eventProperties.EntityInstanceEventProperties[0].EntityName;
                adapter.EventProperties["EntityNamespace"] = eventProperties.EntityInstanceEventProperties[0].EntityNamespace;
                adapter.EventProperties["NotificationContext"] = eventProperties.EntityInstanceEventProperties[0].NotificationContext;
                adapter.EventProperties["NotificationMessage"] = Convert.ToBase64String(eventProperties.EntityInstanceEventProperties[0].NotificationMessage);
            }

            else if (eventProperties.ItemEventProperties.Length > 0 &&
                eventProperties.ItemEventProperties[0].ListItemId != null)
            {
                adapter.EventProperties["ListItemId"] = eventProperties.ItemEventProperties[0].ListItemId;
                adapter.EventProperties["AfterUrl"] = eventProperties.ItemEventProperties[0].AfterUrl;
                adapter.EventProperties["BeforeUrl"] = eventProperties.ItemEventProperties[0].BeforeUrl;
                adapter.EventProperties["CurrentUserId"] = eventProperties.ItemEventProperties[0].CurrentUserId;
                adapter.EventProperties["ExternalNotificationMessage"] = eventProperties.ItemEventProperties[0].ExternalNotificationMessage;
                adapter.EventProperties["IsBackgroundSave"] = eventProperties.ItemEventProperties[0].IsBackgroundSave;
                adapter.EventProperties["ListId"] = eventProperties.ItemEventProperties[0].ListId;
                adapter.EventProperties["ListTitle"] = eventProperties.ItemEventProperties[0].ListTitle;
                adapter.EventProperties["UserDisplayName"] = eventProperties.ItemEventProperties[0].UserDisplayName;
                adapter.EventProperties["UserLoginName"] = eventProperties.ItemEventProperties[0].UserLoginName;
                adapter.EventProperties["Versionless"] = eventProperties.ItemEventProperties[0].Versionless;
                adapter.EventProperties["WebUrl"] = eventProperties.ItemEventProperties[0].WebUrl;
                adapter = GetListItemBeforeAndAfterProperties(adapter, soapBody);
            }

            else if (eventProperties.ListEventProperties.Length > 0 &&
                eventProperties.ListEventProperties[0].ListId != null)
            {
                adapter.EventProperties["ListId"] = eventProperties.ListEventProperties[0].ListId;
                adapter.EventProperties["FeatureId"] = eventProperties.ListEventProperties[0].FeatureId;
                adapter.EventProperties["FieldName"] = eventProperties.ListEventProperties[0].FieldName;
                adapter.EventProperties["FieldXml"] = eventProperties.ListEventProperties[0].FieldXml;
                adapter.EventProperties["ListTitle"] = eventProperties.ListEventProperties[0].ListTitle;
                adapter.EventProperties["TemplateId"] = eventProperties.ListEventProperties[0].TemplateId;
                adapter.EventProperties["WebUrl"] = eventProperties.ListEventProperties[0].WebUrl;
            }

            else if (eventProperties.SecurityEventProperties.Length > 0 &&
                eventProperties.SecurityEventProperties[0].WebId != default(Guid))
            {
                adapter.EventProperties["UserLoginName"] = eventProperties.SecurityEventProperties[0].UserLoginName;
                adapter.EventProperties["GroupName"] = eventProperties.SecurityEventProperties[0].GroupName;
                adapter.EventProperties["GroupUserLoginName"] = eventProperties.SecurityEventProperties[0].GroupUserLoginName;
                adapter.EventProperties["RoleDefinitionName"] = eventProperties.SecurityEventProperties[0].RoleDefinitionName;
                adapter.EventProperties["ScopeUrl"] = eventProperties.SecurityEventProperties[0].ScopeUrl;
                adapter.EventProperties["UserDisplayName"] = eventProperties.SecurityEventProperties[0].UserDisplayName;
                adapter.EventProperties["WebFullUrl"] = eventProperties.SecurityEventProperties[0].WebFullUrl;
                adapter.EventProperties["WebId"] = eventProperties.SecurityEventProperties[0].WebId.ToString();
                adapter.EventProperties["GroupId"] = eventProperties.SecurityEventProperties[0].GroupId.ToString();
                adapter.EventProperties["GroupNewOwnerId"] = eventProperties.SecurityEventProperties[0].GroupNewOwnerId.ToString();
                adapter.EventProperties["GroupOwnerId"] = eventProperties.SecurityEventProperties[0].GroupOwnerId.ToString();
                adapter.EventProperties["GroupUserId"] = eventProperties.SecurityEventProperties[0].GroupUserId.ToString();
                adapter.EventProperties["ObjectType"] = eventProperties.SecurityEventProperties[0].ObjectType.ToString();
                adapter.EventProperties["PrincipalId"] = eventProperties.SecurityEventProperties[0].PrincipalId.ToString();
                adapter.EventProperties["RoleDefinitionId"] = eventProperties.SecurityEventProperties[0].RoleDefinitionId.ToString();
                adapter.EventProperties["RoleDefinitionPermissions"] = eventProperties.SecurityEventProperties[0].RoleDefinitionPermissions.ToString();
            }
            else if (eventProperties.WebEventProperties.Length > 0 &&
                eventProperties.WebEventProperties[0].FullUrl != null)
            {
                adapter.EventProperties["FullUrl"] = eventProperties.WebEventProperties[0].FullUrl;
                adapter.EventProperties["NewServerRelativeUrl"] = eventProperties.WebEventProperties[0].NewServerRelativeUrl;
                adapter.EventProperties["ServerRelativeUrl"] = eventProperties.WebEventProperties[0].ServerRelativeUrl;
            }
            return adapter;
        }
        private static SharePointRemoteEventAdapter GetListItemBeforeAndAfterProperties(SharePointRemoteEventAdapter adapter, string soapBody)
        {
            var afterPropertiesXml =
                soapBody.GetInnerText(
                    "<AfterProperties xmlns:a=\"http://schemas.microsoft.com/2003/10/Serialization/Arrays\">", "</AfterProperties>");

            afterPropertiesXml = CleanXml(afterPropertiesXml);

            var beforePropertiesXml =
                soapBody.GetInnerText(
                    "<BeforeProperties xmlns:a=\"http://schemas.microsoft.com/2003/10/Serialization/Arrays\">", "</BeforeProperties>");

            beforePropertiesXml = CleanXml(beforePropertiesXml);

            SetListItemBeforeAndAfterProperties(adapter.ItemAfterProperties, afterPropertiesXml);
            SetListItemBeforeAndAfterProperties(adapter.ItemBeforeProperties, beforePropertiesXml);

            return adapter;
        }
        private static void SetListItemBeforeAndAfterProperties(Dictionary<string, string> properties, string propertiesXml)
        {
            if (propertiesXml == string.Empty) return;

            var doc = XDocument.Parse(propertiesXml);
            var items = doc.Descendants("item");
            foreach (var xmlItem in items)
            {
                var newKey = xmlItem.Descendants("Key").FirstOrDefault()?.Value;
                if (newKey != null)
                {
                    properties[newKey] = xmlItem.Descendants("Value").FirstOrDefault()?.Value;
                }
            }
        }
        private static string CleanXml(string xml)
        {
            if (xml == string.Empty) return xml;
            //xml = xml.Replace("\"a:", "\"").Replace("i:type=\"b:int\"", "").Replace("i:type=\"b:string\"", "").Replace("xmlns:b=\"http://www.w3.org/2001/XMLSchema\"", "").Replace("KeyValueOfstringanyType", "item").Replace("\"i:", "\"").Replace("\"b:", "\"");
            xml = xml.Replace("<a:", "<").Replace("</a:", "</").Replace("xmlns:b=\"http://www.w3.org/2001/XMLSchema\"", "").Replace("KeyValueOfstringanyType", "item").Replace(" i:", " ").Replace("\"b:", "\"");
            return $"<items>{xml}</items>";
        }
    }
}
