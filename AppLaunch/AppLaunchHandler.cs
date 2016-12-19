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
using AzureFunctionsForSharePoint.Core.Security;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.SharePoint.Client;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using static AzureFunctionsForSharePoint.Core.SecurityTokens;
using static AzureFunctionsForSharePoint.Core.EnqueueMessage;
using static AzureFunctionsForSharePoint.Core.ContextUtility;

namespace AppLaunch
{
    public class AppLauncherFunctionArgs
    {
        public string StorageAccount { get; set; }
        public string StorageAccountKey { get; set; }
    }

    public class AppLaunchHandler : FunctionBase
    {
        private readonly NameValueCollection _formParams;
        private readonly Dictionary<string, string> _queryParams;
        private readonly string _requestAuthority;
        private readonly HttpResponseMessage _response;
        private ClientConfiguration _clientClientConfiguration;

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

        public HttpResponseMessage Execute(AppLauncherFunctionArgs args)
        {
            try
            {
                _clientClientConfiguration = GetConfiguration(ClientId, args.StorageAccount, args.StorageAccountKey);
                var spContextToken = TokenHelper.ReadAndValidateContextToken(ContextToken, _requestAuthority, ClientId,
                    _clientClientConfiguration.ClientSecret);
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
                    AppAccessToken = GetAccessToken(ClientId, encodedCacheKey, true),
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

        public string ContextToken
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

        public string SPWebUrl
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

        public string SPHostUrl
        {
            get
            {
                if (_queryParams != null && _queryParams.ContainsKey("SPHostUrl") &&
                    !string.IsNullOrEmpty(_queryParams?["SPHostUrl"])) return _queryParams["SPHostUrl"];

                throw new ArgumentException("No host web in query string!");
            }
        }

        public string ClientId
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
                            var jwt = TokenHelper.Base64Decode(mainPart);
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
                    AppAccessToken = GetAccessToken(ClientId, cacheKey, true),
                    UserAccessToken = GetAccessToken(ClientId, cacheKey),
                    RetryCount = 5,
                    Action = ProvisioningAction.Install,
                    ProvisioningStep = ProvisioningSteps.NotStarted
                });
            }
        }

        private void InstallBaseManifest(ClientContext clientContext)
        {
            Log("Applying base manifest");
            var manifest = GetBaseManifest(ClientId, _clientClientConfiguration.GetStorageAccount(),
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

            var productId = _clientClientConfiguration.ProductId;
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
