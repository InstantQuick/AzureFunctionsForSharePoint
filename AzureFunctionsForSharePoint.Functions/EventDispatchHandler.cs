using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Script.Serialization;
using AzureFunctionsForSharePoint.Core;
using AzureFunctionsForSharePoint.Common;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.SharePoint.Client;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using static AzureFunctionsForSharePoint.Core.SecurityTokens;
using static AzureFunctionsForSharePoint.Core.EnqueueMessage;
using static AzureFunctionsForSharePoint.Core.ContextUtility;

namespace AzureFunctionsForSharePoint.Functions
{
    /// <summary>
    /// Function specific configuration elements should be added as properties here to extend the <see cref="AzureFunctionArgs" /> class.
    /// </summary>
    public class EventDispatchFunctionArgs : AzureFunctionArgs { }

    /// <summary>
    /// The EventDispatch function receives a remote event from SharePoint as a WCF SOAP message and  parses it using <see cref="SharePointRemoteEventAdapter"/>.
    /// Based on the event type, the received information may be augmented by reading additional information from SharePoint.
    /// EventDispatch sends the resulting <see cref="QueuedSharePointProcessEvent"/> to the client's service bus queue.  
    /// The EventDispatch function receives a remote event from SharePoint as a WCF SOAP message, parses it into something that is easier to consume
    /// using <see cref="SharePointRemoteEventAdapter"/>. Based on the event type, the received information may be augmented by reading additional information from SharePoint. 
    /// EventDispatch sends the resulting <see cref="QueuedSharePointProcessEvent"/> to the client's service bus queue as JSON.
    /// </summary>
    /// <remarks>
    /// This class inherits <see cref="FunctionBase"/> for its simple logging notification event. 
    /// </remarks>
    public class EventDispatchHandler : FunctionBase
    {
        private const string AppOnlyPrincipalId = "00000003-0000-0ff1-ce00-000000000000";
        private readonly Dictionary<string, string> _queryParams;
        private readonly string _requestAuthority;
        private readonly HttpResponseMessage _response;
        private readonly SharePointRemoteEventAdapter _eventInfo;
        private ClientConfiguration _clientConfiguration;

        /// <summary>
        /// Initializes the handler for a given HttpRequestMessage received from the function trigger and parses the incoming WCF message
        /// </summary>
        /// <param name="request">The current request</param>
        public EventDispatchHandler(HttpRequestMessage request)
        {
            try
            {
                var soapBody = request.Content.ReadAsStringAsync().Result;
                _eventInfo = SharePointRemoteEventAdapter.GetSharePointRemoteEventAdapter(soapBody);

                _queryParams = request.GetQueryNameValuePairs()?
                    .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                _requestAuthority = request.RequestUri.Authority;
                _response = request.CreateResponse();
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                throw;
            }
        }

        /// <summary>
        /// Processes the received event and sends the result to the client's service bus queue.
        /// 
        /// SharePoint's remote event notification lacks the current item state for ItemDeleting and ItemUpdating events.
        /// For these event types, it attempts to fetch the current (unchanged) item and populate the ItemBeforeProperties. It is possible for the attempt to fail if the item is already deleted. If the attempt fails, the event is forwarded with the available information.
        /// </summary>
        /// <param name="args">An <see cref="EventDispatchFunctionArgs"/> instance specifying the location of the client configuration in Azure storage.</param>
        /// <remarks>The event is ignored if it is the result of an action taken by an app only identity</remarks>
        /// <returns>HttpStatusCode.OK if all is well or 500.</returns>
        public HttpResponseMessage Execute(EventDispatchFunctionArgs args)
        {
            try
            {
                _response.StatusCode = HttpStatusCode.OK;

                //Ignore the event if it is the result of an action taken by an app only identity
                if (_eventInfo.EventProperties.ContainsKey("UserLoginName") && _eventInfo.EventProperties["UserLoginName"].Contains(AppOnlyPrincipalId))
                {
                    Log("Event source is an app not a user. Ignoring");
                    return _response;
                }

                var clientId = GetClientId();

                if (clientId==null)
                {
                    Log("Request has no client ID. Ignoring");
                    return _response;
                }

                //Connect to the SharePoint site and get access tokens
                try
                {
                    _clientConfiguration = GetConfiguration(clientId, args.StorageAccount, args.StorageAccountKey);
                }
                catch
                {
                    Log("Failed to get client configuration");
                    Log($"Client Id is {clientId}");
                    Log(args.StorageAccount);
                    Log(args.StorageAccountKey);
                    throw;
                }

                var spContextToken = TokenHelper.ReadAndValidateContextToken(ContextToken, _requestAuthority, clientId,
                    _clientConfiguration.AcsClientConfig.ClientSecret);
                var encodedCacheKey = TokenHelper.Base64UrlEncode(spContextToken.CacheKey);
                var spHostUri = new Uri(SPWebUrl);

                var accessToken = TokenHelper.GetACSAccessTokens(spContextToken, spHostUri.Authority,
                   _clientConfiguration.ClientId,
                   _clientConfiguration.AcsClientConfig.ClientSecret);

                var ctx = ConnectToSPWeb(accessToken);

                var securityTokens = new SecurityTokens()
                {
                    ClientId = clientId,
                    AccessToken = accessToken.AccessToken,
                    AccessTokenExpires = accessToken.ExpiresOn,
                    AppWebUrl = SPWebUrl,
                    Realm = spContextToken.Realm,
                    RefreshToken = spContextToken.RefreshToken
                };

                Log($"Storing tokens for {clientId}/{encodedCacheKey}");
                StoreSecurityTokens(securityTokens, encodedCacheKey, args.StorageAccount, args.StorageAccountKey);

                //Create the event message to send to the client's service bus queue
                var eventMessage = new QueuedSharePointProcessEvent()
                {
                    SharePointRemoteEventAdapter = _eventInfo,
                    ClientId = _clientConfiguration.ClientId,
                    AppWebUrl = SPWebUrl,
                    UserAccessToken = accessToken.AccessToken,
                    AppAccessToken = GetACSAccessTokens(clientId, encodedCacheKey, true),
                };

                //SharePoint's remote event notification lacks the current item state for ItemDeleting and ItemUpdating events
                //For these event types, attempt to fetch the current (unchanged) item and populate the ItemBeforeProperties
                if (_eventInfo.EventType == "ItemDeleting" || _eventInfo.EventType == "ItemUpdating")
                {
                    //SharePoint feature provisioning sometimes raises this event
                    //and deletes some things in the process with no ListId given
                    var listId = Guid.Parse(_eventInfo.EventProperties["ListId"]);
                    if (listId != default(Guid))
                    {
                        var item =
                            ctx.Web.Lists.GetById(Guid.Parse(_eventInfo.EventProperties["ListId"]))
                                .GetItemById(_eventInfo.EventProperties["ListItemId"]);
                        ctx.Load(item, i => i.FieldValuesAsText);
                        try
                        {
                            ctx.ExecuteQueryRetry();
                            _eventInfo.ItemBeforeProperties = item.FieldValuesAsText.FieldValues;
                        }
                        catch
                        {
                            //The query depends on timing and there are a number of things that can go wrong. 
                            //If the BeforeProperties can't be read, forward the event anyway with the info that is available
                        }
                    }
                }

                //Send the event to the client's service bus queue
                SendQueueMessage(eventMessage);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                throw;
            }

            return _response;
        }

        private string ContextToken
        {
            get { return _eventInfo.GetContextToken(); }
        }

        private List<string> WebPropertyNames = new List<string>() { "AppWebFullUrl", "HostWebFullUrl", "WebUrl", "WebFullUrl", "FullUrl" };
        private string SPWebUrl
        {
            get
            {
                var urlKey = _eventInfo.EventProperties.Keys.FirstOrDefault(k => WebPropertyNames.Contains(k) && _eventInfo.EventProperties[k] != null);

                if (urlKey == null)
                {
                    Log("");
                }
                return urlKey != null
                  ? _eventInfo.EventProperties[urlKey]
                  : string.Empty;
            }
        }

        private string GetClientId()
        {
            if (_queryParams != null && _queryParams.ContainsKey("clientId") &&
                !string.IsNullOrEmpty(_queryParams?["clientId"])) return _queryParams["clientId"];
            else
            {
                var contextTokenParts = ContextToken?.Split('.');
                if (contextTokenParts != null && contextTokenParts.Length > 1)
                {
                    var mainPart = contextTokenParts[1];
                    try
                    {
                        var jwt = TokenHelper.Base64DecodeJwtToken(mainPart);
                        var deserializer = new JavaScriptSerializer();
                        var tokenProperties = deserializer.Deserialize<Dictionary<string, string>>(jwt);
                        if (tokenProperties.ContainsKey("aud"))
                        {
                            return tokenProperties["aud"].Split('/')[0];
                        }
                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
            return null;
        }

        private ClientContext ConnectToSPWeb(OAuth2AccessTokenResponse accessToken)
        {
            var ctx = TokenHelper.GetClientContext(SPWebUrl, accessToken.AccessToken);
            ctx.Load(ctx.Web);
            ctx.ExecuteQueryRetry();
            Log($"Connected to {ctx.Web.Url}");
            return ctx;
        }
    }
}
