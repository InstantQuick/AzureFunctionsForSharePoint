using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using AzureFunctionsForSharePoint.Core;
using AzureFunctionsForSharePoint.Common;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using static AzureFunctionsForSharePoint.Core.SecurityTokens;

namespace AzureFunctionsForSharePoint.Functions
{
    /// <summary>
    /// Function specific configuration elements should be added as properties here to extend the <see cref="AzureFunctionArgs" /> class.
    /// </summary>
    public class GetACSAccessTokensArgs : AzureFunctionArgs { }

    /// <summary>
    /// Returns a JSON string containing userAccessToken and appOnlyAccessToken properties for a given clientId and cacheKey combo.
    /// </summary>
    /// <remarks>
    /// This class inherits <see cref="FunctionBase"/> for its simple logging notification event. 
    /// </remarks>
    public class GetACSAccessTokensHandler : FunctionBase
    {
        private static string targetPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        private readonly Dictionary<string, string> _queryParams;
        private readonly HttpResponseMessage _response;

        /// <summary>
        /// Initialize the function and populate the query params collection.
        /// </summary>
        /// <param name="request">The current request</param>
        public GetACSAccessTokensHandler(HttpRequestMessage request)
        {
            _queryParams = request.GetQueryNameValuePairs()?
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            _response = request.CreateResponse();
        }

        /// <summary>
        /// Returns application/json containing userAccessToken and appOnlyAccessToken properties for a valid clientId and cacheKey combo or a 404 for invalid input.
        /// </summary>
        /// <param name="args">An <see cref="GetACSAccessTokensArgs"/> instance specifying the location of the client configuration in Azure storage.</param>
        /// <returns>JSON or 404</returns>
        public HttpResponseMessage Execute(GetACSAccessTokensArgs args)
        {
            try
            {
                var cacheKey = _queryParams["cacheKey"];
                var clientId = _queryParams["clientId"];

                var clientConfig = GetConfiguration(clientId);
                var tokens = GetSecurityTokens(cacheKey, clientId);

                Uri hostUri = new Uri(tokens.AppWebUrl);

                var userAccessToken = GetUserAccessToken(cacheKey, tokens, hostUri, clientConfig);
                var appOnlyAccessToken = ContextUtility.GetAppOnlyAccessToken(clientId, cacheKey);
                _response.StatusCode = HttpStatusCode.OK;
                _response.Content = new StringContent($"{{\"userAccessToken\":\"{userAccessToken.AccessToken}\", \"appOnlyAccessToken\":\"{appOnlyAccessToken}\"}}");
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
            var userAccessToken = TokenHelper.GetACSAccessTokens(tokens.RefreshToken, targetPrincipal, hostUri.Authority,
                tokens.Realm, tokens.ClientId, clientConfig.AcsClientConfig.ClientSecret);

            tokens.AccessToken = userAccessToken.AccessToken;
            tokens.AccessTokenExpires = userAccessToken.ExpiresOn;
            StoreSecurityTokens(tokens, cacheKey);
            return userAccessToken;
        }
    }
}
