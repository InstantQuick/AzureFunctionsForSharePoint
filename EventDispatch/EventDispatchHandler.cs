using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Script.Serialization;
using AzureFunctionsForSharePoint.Core;
using AzureFunctionsForSharePoint.Common;
using AzureFunctionsForSharePoint.Core.Security;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.SharePoint.Client;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using static AzureFunctionsForSharePoint.Core.SecurityTokens;
using static AzureFunctionsForSharePoint.Core.EnqueueMessage;
using static AzureFunctionsForSharePoint.Core.ContextUtility;

namespace EventDispatch
{
    /// <summary>
    /// Function specific configuration elements should be added as properties here to extend the <see cref="AzureFunctionArgs" /> class.
    /// </summary>
    public class EventDispatchFunctionArgs : AzureFunctionArgs { }

    /// <summary>
    /// This function receive a remote event from SharePoint as a WCF SOAP message and parses it using <see cref="SharePointRemoteEventAdapter"/>.
    /// Based on the event type, the received information may be augmented by reading additional information from SharePoint.
    ///  and sends the resulting <see cref="QueuedSharePointProcessEvent"/> to the client's service bus queue.  
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
        private ClientConfiguration _clientClientConfiguration;

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
        /// 
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public HttpResponseMessage Execute(EventDispatchFunctionArgs args)
        {
            try
            {
                _response.StatusCode = HttpStatusCode.OK;
                if (_eventInfo.EventProperties.ContainsKey("UserLoginName") && _eventInfo.EventProperties["UserLoginName"].Contains(AppOnlyPrincipalId)) return _response;

                _clientClientConfiguration = GetConfiguration(ClientId, args.StorageAccount, args.StorageAccountKey);
                var spContextToken = TokenHelper.ReadAndValidateContextToken(ContextToken, _requestAuthority, ClientId,
                    _clientClientConfiguration.ClientSecret);
                var encodedCacheKey = TokenHelper.Base64UrlEncode(spContextToken.CacheKey);
                var spHostUri = new Uri(SPWebUrl);

                var accessToken = TokenHelper.GetAccessToken(spContextToken, spHostUri.Authority,
                   _clientClientConfiguration.ClientId,
                   _clientClientConfiguration.ClientSecret);

                var ctx = ConnectToSPWeb(accessToken);

                var securityTokens = new SecurityTokens()
                {
                    ClientId = ClientId,
                    AccessToken = accessToken.AccessToken,
                    AccessTokenExpires = accessToken.ExpiresOn,
                    AppWebUrl = SPWebUrl,
                    Realm = spContextToken.Realm,
                    RefreshToken = spContextToken.RefreshToken
                };

                Log($"Storing tokens for {ClientId}/{encodedCacheKey}");
                StoreSecurityTokens(securityTokens, encodedCacheKey, args.StorageAccount, args.StorageAccountKey);

                var eventMessage = new QueuedSharePointProcessEvent()
                {
                    SharePointRemoteEventAdapter = _eventInfo,
                    ClientId = _clientClientConfiguration.ClientId,
                    AppWebUrl = SPWebUrl,
                    UserAccessToken = accessToken.AccessToken,
                    AppAccessToken = GetAccessToken(ClientId, encodedCacheKey, true),
                };


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
                SendQueueMessage(eventMessage);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                throw;
            }

            return _response;
        }

        public string ContextToken
        {
            get { return _eventInfo.GetContextToken(); }
        }

        private List<string> WebPropertyNames = new List<string>() { "AppWebFullUrl", "HostWebFullUrl", "WebUrl", "WebFullUrl", "FullUrl" };
        public string SPWebUrl
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

        public string ClientId
        {
            get
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
                            var jwt = TokenHelper.Base64Decode(mainPart);
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
