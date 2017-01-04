using System;
using AzureFunctionsForSharePoint.Core.Security;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.SharePoint.Client;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using static AzureFunctionsForSharePoint.Core.SecurityTokens;

namespace AzureFunctionsForSharePoint.Core
{
    /// <summary>
    /// Methods for connecting to SharePoint
    /// </summary>
    public class ContextUtility
    {
        private static string targetPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        /// Gets a SharePoint <see cref="ClientContext"/> for a given client id and cache key combination
        /// </summary>
        /// <param name="clientId">The id of the client. This must match the client id from a SharePoint add-in manifest and a valid <see cref="ClientConfiguration"/>.</param>
        /// <param name="cacheKey">A valid cache key</param>
        /// <param name="appOnly">Set to true for an app only context</param>
        /// <param name="fallbackToUser">Be the user if app only is not allowed</param>
        /// <returns>A ClientContext object</returns>
        public static ClientContext GetClientContext(string clientId, string cacheKey, bool appOnly = false, bool fallbackToUser = true)
        {
            try
            {
                ClientContext userClientContext;
                var clientConfig = GetConfiguration(clientId);
                var tokens = GetSecurityTokens(cacheKey, clientId);
                if (tokens == null) return null;
                Uri hostUri = new Uri(tokens.AppWebUrl);

                //Always try to get access as the user. If the user has no access, this should
                //never return an app only context
                var userAccessToken = GetUserAccessToken(cacheKey, tokens, hostUri, clientConfig);
                userClientContext = TokenHelper.GetClientContext(tokens.AppWebUrl, userAccessToken.AccessToken);
                //Never! If the user hasn't got access
                if (!ContextHasAccess(userClientContext)) return null;
                
                string accessToken = GetAppOnlyAccessToken(targetPrincipal, hostUri.Authority, tokens.Realm,
                    tokens.ClientId, clientConfig.ClientSecret);

                var appOnlyContext = TokenHelper.GetClientContext(tokens.AppWebUrl, accessToken);
                //If an app only context isn't available this is an older version
                //Fall back to the user's context
                if (!ContextHasAccess(appOnlyContext))
                {
                    if (fallbackToUser) return userClientContext;
                    else return null;
                }
                return appOnlyContext;
            }
            catch (Exception ex)
            {
                var detailedException = new Exception($"Unable to get client context for cId={clientId} cacheKey={cacheKey}", ex);
                throw (detailedException);
            }
        }

        /// <summary>
        /// Gets an access token for a given client id and cache key combination
        /// </summary>
        /// <param name="clientId">The id of the client. This must match the client id from a SharePoint add-in manifest and a valid <see cref="ClientConfiguration"/>.</param>
        /// <param name="cacheKey">A valid cache key</param>
        /// <param name="appOnly">Set to true for an app only token</param>
        /// <param name="fallbackToUser">Be the user if app only is not allowed</param>
        /// <returns>An access token string</returns>
        public static string GetAccessToken(string clientId, string cacheKey, bool appOnly = false, bool fallbackToUser = true)
        {
            try
            {
                var clientConfig = GetConfiguration(clientId);
                var tokens = GetSecurityTokens(cacheKey, clientId);
                if (tokens == null) return null;
                Uri hostUri = new Uri(tokens.AppWebUrl);
                
                if (!appOnly) return GetUserAccessToken(cacheKey, tokens, hostUri, clientConfig).AccessToken;
                return GetAppOnlyAccessToken(targetPrincipal, hostUri.Authority, tokens.Realm,
                                    tokens.ClientId, clientConfig.ClientSecret);
            }
            catch (Exception ex)
            {
                var detailedException = new Exception($"Unable to get client context for cId={clientId} cacheKey={cacheKey}", ex);
                throw (detailedException);
            }
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

        private static string GetAppOnlyAccessToken(string targetPrincipalName, string authority, string realm, string clientId, string clientSecret)
        {
            return TokenHelper.GetAppOnlyAccessToken(targetPrincipalName, authority, realm, clientId, clientSecret).AccessToken;
        }

        public static string GetAppOnlyAccessToken(string clientId, string cacheKey)
        {
            return GetAccessToken(clientId, cacheKey, true, false);
        }

        private static bool ContextHasAccess(ClientContext ctx)
        {
            try
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();
            }
            catch (ServerUnauthorizedAccessException)
            {
                return false;
            }
            return true;
        }
    }
}
