using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using AzureFunctionsForSharePoint.Core;
using AzureFunctionsForSharePoint.Common;
using AzureFunctionsForSharePoint.Core.Security;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using static AzureFunctionsForSharePoint.Core.SecurityTokens;

namespace GetAccessToken
{
    public class GetAccessTokenArgs
    {
        public string StorageAccount { get; set; }
        public string StorageAccountKey { get; set; }
    }
    public class GetAccessTokenHandler : FunctionBase
    {
        private static string targetPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        private readonly Dictionary<string, string> _queryParams;
        private readonly HttpResponseMessage _response;

        public GetAccessTokenHandler(HttpRequestMessage request)
        {
            _queryParams = request.GetQueryNameValuePairs()?
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            _response = request.CreateResponse();
        }

        public HttpResponseMessage Execute(GetAccessTokenArgs args)
        {
            try
            {
                var cacheKey = _queryParams["cacheKey"];
                var clientId = _queryParams["clientId"];

                var clientConfig = GetConfiguration(clientId);
                var tokens = GetSecurityTokens(cacheKey, clientId);

                Uri hostUri = new Uri(tokens.AppWebUrl);

                //Always try to get access as the user. If the user has no access, this should
                //never return an app only context
                var userAccessToken = GetUserAccessToken(cacheKey, tokens, hostUri, clientConfig);
                _response.StatusCode = HttpStatusCode.OK;
                _response.Content = new StringContent($"{{\"token\":\"{userAccessToken.AccessToken}\"}}");
                _response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            }
            catch
            {
                _response.StatusCode = HttpStatusCode.NotFound;
            }
            return _response;
        }

        private static OAuth2AccessTokenResponse GetUserAccessToken(string cacheKey, SecurityTokens tokens, Uri hostUri, ClientConfiguration clientConfig)
        {
            var userAccessToken = TokenHelper.GetAccessToken(tokens.RefreshToken, targetPrincipal, hostUri.Authority,
                tokens.Realm, tokens.ClientId, clientConfig.ClientSecret);

            tokens.AccessToken = userAccessToken.AccessToken;
            tokens.AccessTokenExpires = userAccessToken.ExpiresOn;
            StoreSecurityTokens(tokens, cacheKey);
            return userAccessToken;
        }
    }
}
