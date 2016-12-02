using System;
using ClientConfiguration;
using IQAppCommon.Security;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.SharePoint.Client;
using TokenStorage;
using static ClientConfiguration.Configuration;
using static TokenStorage.BlobStorage;

namespace IQAppCommon
{
    public class ContextUtility
    {
        private static string targetPrincipal = "00000003-0000-0ff1-ce00-000000000000";

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
                System.Diagnostics.Trace.TraceError($"Unable to get client context for {cacheKey}|{ex.ToString()}");
                throw (ex);
            }
        }

        public static string GetAccessToken(string clientId, string cacheKey, bool appOnly = false, bool fallbackToUser = true)
        {
            try
            {
                var clientConfig = GetConfiguration(clientId);
                var tokens = GetSecurityTokens(cacheKey, clientId);
                if (tokens == null) return null;
                Uri hostUri = new Uri(tokens.AppWebUrl);

                //Always try to get access as the user. If the user has no access, this should

                if (!appOnly) return GetUserAccessToken(cacheKey, tokens, hostUri, clientConfig).AccessToken;
                return GetAppOnlyAccessToken(targetPrincipal, hostUri.Authority, tokens.Realm,
                                    tokens.ClientId, clientConfig.ClientSecret);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError($"Unable to get client context for {cacheKey}|{ex.ToString()}");
                throw (ex);
            }
        }

        private static OAuth2AccessTokenResponse GetUserAccessToken(string cacheKey, SecurityTokens tokens, Uri hostUri, Configuration clientConfig)
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

        private static bool ContextHasAccess(ClientContext ctx)
        {
            try
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();
            }
            catch (ServerUnauthorizedAccessException ex)
            {
                return false;
            }
            return true;
        }
    }
}
