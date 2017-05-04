using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
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
    public class AppLauncherFunctionArgs : AzureFunctionArgs { }

    /// <summary>
    /// This function is called when SharePoint POSTs an ACS token for a SharePoint add-in.
    /// The SharePoint add-in's manifest XML must specify the function URL as the value of the <see href="https://msdn.microsoft.com/en-us/library/office/jj583318.aspx">StartPage element</see>. 
    /// A valid client configuration is required.
    /// 
    /// Once connected to a SharePoint site, the function checks the add-in's install status and provisions as indicated by the bootstrapmanifest.json located in the client's configuration storage container. If provisioning occurs a message is sent to the service bus queue specified in the client configuration to notify the client for additional processing as desired.
    /// Finally,  a message is sent to the service bus queue specified in the client configuration to notify the client of the add-in's launch.
    /// </summary>
    /// <remarks>
    /// This class inherits <see cref="FunctionBase"/> for its simple logging notification event. 
    /// </remarks>
    public class AppLaunchHandler : FunctionBase
    {
        private readonly NameValueCollection _formParams;
        private readonly Dictionary<string, string> _queryParams;
        private readonly string _requestAuthority;
        private readonly HttpResponseMessage _response;
        private ClientConfiguration _clientClientConfiguration;

        /// <summary>
        /// Initializes the handler for a given HttpRequestMessage received from the function trigger
        /// </summary>
        /// <param name="request">The current request</param>
        public AppLaunchHandler(HttpRequestMessage request)
        {
            if (request.Content.IsFormData())
            {
                _formParams = request.Content.ReadAsFormDataAsync().Result;
            }

            _queryParams = request.GetQueryNameValuePairs()?
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            _requestAuthority = request.RequestUri.Authority;
            _response = request.CreateResponse();
        }

        /// <summary>
        /// Performs the app launch flow for the current request
        /// </summary>
        /// <param name="args">An <see cref="AppLauncherFunctionArgs"/> instance specifying the location of the client configuration in Azure storage.</param>
        /// <returns>If launch succeeds the response is a 302 redirect back to the SharePoint site's home page.</returns>
        public HttpResponseMessage Execute(AppLauncherFunctionArgs args)
        {
            try
            {
                _clientClientConfiguration = GetConfiguration(ClientId, args.StorageAccount, args.StorageAccountKey);
                var spContextToken = TokenHelper.ReadAndValidateContextToken(ContextToken, _requestAuthority, ClientId,
                    _clientClientConfiguration.AcsClientConfig.ClientSecret);
                var spHostUri = new Uri(SPWebUrl);

                var accessToken = TokenHelper.GetACSAccessTokens(spContextToken, spHostUri.Authority,
                    _clientClientConfiguration.ClientId,
                    _clientClientConfiguration.AcsClientConfig.ClientSecret);


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

                var encodedCacheKey = TokenHelper.Base64UrlEncode(spContextToken.CacheKey);
                Log($"Storing tokens for {ClientId}/{encodedCacheKey}");
                StoreSecurityTokens(securityTokens, encodedCacheKey, args.StorageAccount, args.StorageAccountKey);

                Log($"Ensuring web properties for {ctx.Web.Url}");
                EnsureBaseConfiguration(encodedCacheKey);

                Log($"Sending app launch event for {ctx.Web.Url}");
                SendQueueMessage(new QueuedAppLaunchEvent()
                {
                    ClientId = ClientId,
                    AppWebUrl = ctx.Web.Url,
                    UserAccessToken = securityTokens.AccessToken,
                    AppAccessToken = GetACSAccessTokens(ClientId, encodedCacheKey, true),
                    RetryCount = 5
                });

                _response.StatusCode = HttpStatusCode.Moved;
                _response.Headers.Location = new Uri($"{ctx.Web.Url}?cId={ClientId}&cKey={encodedCacheKey}");

                return _response;
            }
            catch (Exception ex)
            {
                _response.StatusCode = HttpStatusCode.OK;
                _response.Content = new StringContent(GetErrorPage(ex.ToString()));
                _response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                return _response;
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

        private string ContextToken
        {
            get
            {
                string[] paramNames = { "AppContext", "AppContextToken", "AccessToken", "SPAppToken" };

                foreach (string paramName in paramNames)
                {
                    if (_formParams != null && !string.IsNullOrEmpty(_formParams[paramName])) return _formParams[paramName];
                    if (_queryParams != null && _queryParams.ContainsKey(paramName) &&
                        !string.IsNullOrEmpty(_queryParams?[paramName])) return _queryParams[paramName];
                }
                return null;
            }
        }

        private string SPWebUrl
        {
            get
            {
                if (_queryParams != null && _queryParams.ContainsKey("SPAppWebUrl") &&
                    !string.IsNullOrEmpty(_queryParams?["SPAppWebUrl"])) return _queryParams["SPAppWebUrl"];

                if (_queryParams != null && _queryParams.ContainsKey("SPHostUrl") &&
                    !string.IsNullOrEmpty(_queryParams?["SPHostUrl"])) return _queryParams["SPHostUrl"];

                throw new ArgumentException("No app web or host web in query string!");
            }
        }

        private string SPHostUrl
        {
            get
            {
                if (_queryParams != null && _queryParams.ContainsKey("SPHostUrl") &&
                    !string.IsNullOrEmpty(_queryParams?["SPHostUrl"])) return _queryParams["SPHostUrl"];

                throw new ArgumentException("No host web in query string!");
            }
        }

        private string ClientId
        {
            get
            {
                if (_queryParams != null && _queryParams.ContainsKey("clientId") &&
                    !string.IsNullOrEmpty(_queryParams?["clientId"])) return _queryParams["clientId"].ToLower();
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
                                return tokenProperties["aud"].Split('/')[0].ToLower();
                            }
                        }
                        catch
                        {
                            //ignore
                        }
                    }
                }
                return null;
            }
        }

        private bool IsHostWeb
        {
            get
            {
                if (_queryParams != null && _queryParams.ContainsKey("SPAppWebUrl") &&
                    !string.IsNullOrEmpty(_queryParams?["SPAppWebUrl"])) return false;

                return true;
            }
        }

        private void EnsureBaseConfiguration(string cacheKey)
        {
            var appOnlyContext = GetClientContext(ClientId, cacheKey, true);

            appOnlyContext.Web.EnsureProperty(w => w.AllProperties);
            var propKeys = appOnlyContext.Web.AllProperties.FieldValues;
            var props = appOnlyContext.Web.AllProperties;
            if (!propKeys.ContainsKey($"appRedirectUrl.{ClientId}"))
            {
                SetAppRedirectUrlProperty(appOnlyContext, props);
                InstallBaseManifest(appOnlyContext);
                SendQueueMessage(new QueuedSharePointProvisioningEvent()
                {
                    ClientId = ClientId,
                    AppWebUrl = SPWebUrl,
                    AppAccessToken = GetAppOnlyAccessToken(ClientId, cacheKey),
                    UserAccessToken = GetACSAccessTokens(ClientId, cacheKey),
                    RetryCount = 5,
                    Action = ProvisioningAction.Install,
                    ProvisioningStep = ProvisioningSteps.NotStarted
                });
            }
        }

        private void InstallBaseManifest(ClientContext clientContext)
        {
            Log("Applying base manifest");
            var manifest = GetBootstrapManifest(ClientId, _clientClientConfiguration.GetStorageAccount(),
                _clientClientConfiguration.GetStorageAccountKey());

            var provisioner = new BootstrapProvisioner();
            provisioner.Notify += (sender, eventArgs) => { Log(eventArgs.Detail); };

            provisioner.Provision(clientContext, clientContext.Web, manifest);
        }


        private void SetAppRedirectUrlProperty(ClientContext clientContext, PropertyValues props)
        {
            var instanceId = GetAppInstanceId(clientContext);
            if (string.IsNullOrEmpty(instanceId))
            {
                //Don't throw because this is optional
                return;
            }
            props[$"appRedirectUrl.{ClientId}"] =
                    $"{SPHostUrl}/_layouts/15/appredirect.aspx?instance_id={instanceId}";
            try
            {
                clientContext.Web.Update();
                clientContext.Load(clientContext.Web, w => w.Id, w => w.ServerRelativeUrl, w => w.Url,
                    w => w.AllProperties, w => w.AppInstanceId);
                clientContext.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {
                Log($"AppInit error setting web properties: {ex}");
                throw;
            }
        }

        private string GetAppInstanceId(ClientContext clientContext)
        {
            //The AppInstanceId is required when redirecting for cache key
            //It is directly available if this is an app web
            if (!IsHostWeb)
            {
                return clientContext.Web.AppInstanceId.ToString();
            }

            var productId = _clientClientConfiguration.AcsClientConfig.ProductId;
            if (string.IsNullOrEmpty(productId))
            {
                return null;
            }

            //Otherwise fetch it from the host web using the productId from the client catalog
            var instances = clientContext.Web.GetAppInstancesByProductId(Guid.Parse(productId));
            clientContext.Load(instances);
            clientContext.ExecuteQueryRetry();

            if (instances.Count == 0)
            {
                throw new InvalidOperationException("Unable to get the app instance ID!");
            }
            return instances[0].Id.ToString();
        }

        private string GetErrorPage(string errorText)
        {
            try
            {
                return
                    GetFile("AppLaunch.Resources.Error.html", Assembly.GetExecutingAssembly())
                        .Replace("{{Exception}}", errorText).Replace("{{AppId}}", ClientId);
            }
            catch
            {
                return "";
            }
        }

        private string GetFile(string key, Assembly assembly)
        {
            if (assembly == null) return null;
            var stream = assembly.GetManifestResourceStream(key);
            if (stream == null) return null;
            using (var streamReader = new StreamReader(stream))
            {
                return Encoding.UTF8.GetString(ReadFully(streamReader.BaseStream));
            }
        }

        private byte[] ReadFully(Stream input)
        {
            var buffer = new byte[16 * 1024];
            using (var ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}
