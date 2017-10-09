using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IdentityModel.Selectors;
using System.IdentityModel.Tokens;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.IdentityModel;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.SharePoint.Client;
using System.Web.Script.Serialization;

namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Methods for working with auth tokens and creating authorized client contexts
    /// </summary>
    public class TokenHelper
    {
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
        private static string _globalEndPointPrefix = "accounts";
        private static string _acsHostUrl = "accesscontrol.windows.net";
        private const string S2SProtocol = "OAuth2";

        /// <summary>
        /// Uses the specified access token to create a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="accessToken">Access token to be used when calling the specified targetUrl</param>
        /// <returns>A ClientContext ready to call targetUrl with the specified access token</returns>
        public static ClientContext GetClientContext(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);

            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        /// <summary>
        /// Reads and validates a contextTokenString sent from SharePoint as the result of app launch or a remote event
        /// </summary>
        /// <param name="contextTokenString">The string sent by SharePoint</param>
        /// <param name="appHostName">The app host (host part of the site URL)</param>
        /// <param name="clientId">A valid client id</param>
        /// <param name="clientSecret">A valid client secret</param>
        /// <returns></returns>
        public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName, string clientId, string clientSecret)
        {
            JsonWebSecurityTokenHandler tokenHandler = CreateJsonWebSecurityTokenHandler(clientSecret);
            SecurityToken securityToken = tokenHandler.ReadToken(contextTokenString);
            JsonWebSecurityToken jsonToken = securityToken as JsonWebSecurityToken;
            SharePointContextToken token = SharePointContextToken.Create(jsonToken);

            string stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
            int firstDot = stsAuthority.IndexOf('.');

            _globalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            _acsHostUrl = stsAuthority.Substring(firstDot + 1);

            tokenHandler.ValidateToken(jsonToken);

            string realm = token.Realm;
            string principal = GetFormattedPrincipal(clientId, appHostName, realm);
            if (!StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal))
            {
                throw new Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException(
                    String.Format(CultureInfo.CurrentCulture,
                    "\"{0}\" is not the intended audience \"{1}\"", principal, token.Audience));
            }

            return token;
        }

        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!String.IsNullOrEmpty(hostName))
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }
            else
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
            }
        }

        private static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }

        private static JsonWebSecurityTokenHandler CreateJsonWebSecurityTokenHandler(string clientSecret)
        {
            JsonWebSecurityTokenHandler handler = new JsonWebSecurityTokenHandler();
            handler.Configuration = new Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration();
            handler.Configuration.AudienceRestriction = new Microsoft.IdentityModel.Tokens.AudienceRestriction(AudienceUriMode.Never);
            handler.Configuration.CertificateValidator = X509CertificateValidator.None;

            List<byte[]> securityKeys = new List<byte[]>();
            securityKeys.Add(Convert.FromBase64String(clientSecret));

            List<SecurityToken> securityTokens = new List<SecurityToken>();
            securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

            handler.Configuration.IssuerTokenResolver =
                SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
                new ReadOnlyCollection<SecurityToken>(securityTokens),
                false);
            SymmetricKeyIssuerNameRegistry issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
            foreach (byte[] securitykey in securityKeys)
            {
                issuerNameRegistry.AddTrustedIssuer(securitykey, GetAcsPrincipalName(""));
            }
            handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
            return handler;
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            return String.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", _globalEndPointPrefix, _acsHostUrl);
        }


        /// <summary>
        /// Retrieves an access token from ACS to call the source of the specified context token at the specified 
        /// targetHost. The targetHost must be registered for principal the that sent the context token.
        /// </summary>
        /// <param name="contextToken">Context token issued by the intended access token audience</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="clientId">ACS client id</param>
        /// <param name="clientSecret">ACS client secret</param>
        /// <returns>An access token with an audience matching the context token's source</returns>
        public static OAuth2AccessTokenResponse GetACSAccessTokens(SharePointContextToken contextToken, string targetHost, string clientId, string clientSecret)
        {

            string targetPrincipalName = contextToken.TargetPrincipalName;

            // Extract the refreshToken from the context token
            string refreshToken = contextToken.RefreshToken;

            if (String.IsNullOrEmpty(refreshToken))
            {
                return null;
            }

            string realm = contextToken.Realm;

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, realm);
            string formattedPrincipal = GetFormattedPrincipal(clientId, null, realm);

            OAuth2AccessTokenRequest oauth2Request =
                OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(
                    formattedPrincipal,
                    clientSecret,
                    refreshToken,
                    resource);

            // Get token
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(GetStsUrl(realm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                if (wex.Response == null) throw;
                var stream = wex.Response.GetResponseStream();
                if (stream == null) throw;
                using (StreamReader sr = new StreamReader(stream))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// Uses the specified refresh token to retrieve an access token from ACS to call the specified principal 
        /// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
        /// null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="refreshToken">Refresh token to exchange for access token</param>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <param name="clientId">ACS client id</param>
        /// <param name="clientSecret">Client secret</param>
        /// <returns>An access token with an audience of the target principal</returns>
        public static OAuth2AccessTokenResponse GetACSAccessTokens(
            string refreshToken,
            string targetPrincipalName,
            string targetHost,
            string targetRealm,
            string clientId,
            string clientSecret)
        {
            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string formattedPrincipal = GetFormattedPrincipal(clientId, null, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(formattedPrincipal, clientSecret, refreshToken, resource);

            // Get token
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                if (wex.Response == null) throw;
                var stream = wex.Response.GetResponseStream();
                if (stream == null) throw;
                using (StreamReader sr = new StreamReader(stream))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// Retrieves an app-only access token from ACS to call the specified principal 
        /// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is 
        /// null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <param name="clientId">ACS client id</param>
        /// <param name="clientSecret">ACS client secret</param>
        /// <returns>An access token with an audience of the target principal</returns>
        public static OAuth2AccessTokenResponse GetAppOnlyAccessToken(
            string targetPrincipalName,
            string targetHost,
            string targetRealm,
            string clientId,
            string clientSecret)
        {

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string formattedPrincipal = GetFormattedPrincipal(clientId, "", targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(formattedPrincipal, clientSecret, resource);
            oauth2Request.Resource = resource;

            // Get token
            OAuth2S2SClient client = new OAuth2S2SClient();

            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex)
            {
                if (wex.Response == null) throw;
                var stream = wex.Response.GetResponseStream();
                if (stream == null) throw;
                using (StreamReader sr = new StreamReader(stream))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        private static string GetStsUrl(string realm)
        {
            JsonMetadataDocument document = GetMetadataDocument(realm);

            JsonEndpoint s2SEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

            if (null != s2SEndpoint)
            {
                return s2SEndpoint.location;
            }
            else
            {
                throw new Exception("Metadata document does not contain STS endpoint url");
            }
        }

        private static JsonMetadataDocument GetMetadataDocument(string realm)
        {
            string acsMetadataEndpointUrlWithRealm = String.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
                                                                    GetAcsMetadataEndpointUrl(),
                                                                    realm);
            byte[] acsMetadata;
            using (WebClient webClient = new WebClient())
            {

                acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
            }
            string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            JsonMetadataDocument document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

            if (null == document)
            {
                throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
            }

            return document;
        }

        private static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        private class JsonMetadataDocument
        {
            public string serviceName { get; set; }
            public List<JsonEndpoint> endpoints { get; set; }
            public List<JsonKey> keys { get; set; }
        }

        private class JsonEndpoint
        {
            public string location { get; set; }
            public string protocol { get; set; }
            public string usage { get; set; }
        }

        private class JsonKey
        {
            public string usage { get; set; }
            public JsonKeyValue keyValue { get; set; }
        }

        private class JsonKeyValue
        {
            public string type { get; set; }
            public string value { get; set; }
        }

        public static string Base64UrlEncode(string input)
        {
            var inputBytes = System.Text.Encoding.UTF8.GetBytes(input);
            // Special "url-safe" base64 encode.
            return Convert.ToBase64String(inputBytes)
              .Replace('+', '-')
              .Replace('/', '_')
              .Replace("=", "");
        }

        static readonly Encoding TextEncoding = Encoding.UTF8;

        static readonly char Base64PadCharacter = '=';
        static readonly char Base64Character62 = '+';
        static readonly char Base64Character63 = '/';
        static readonly char Base64UrlCharacter62 = '-';
        static readonly char Base64UrlCharacter63 = '_';

        private static byte[] DecodeBytes(string arg)
        {
            if (String.IsNullOrEmpty(arg))
            {
                throw new ApplicationException("String to decode cannot be null or empty.");
            }

            StringBuilder s = new StringBuilder(arg);
            s.Replace(Base64UrlCharacter62, Base64Character62);
            s.Replace(Base64UrlCharacter63, Base64Character63);

            int pad = s.Length % 4;
            s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

            return Convert.FromBase64String(s.ToString());
        }

        public static string Base64DecodeJwtToken(string arg)
        {
            return TextEncoding.GetString(DecodeBytes(arg));
        }

        /// <summary>
        /// Get authentication realm from SharePoint
        /// </summary>
        /// <param name="targetApplicationUri">Url of the target SharePoint site</param>
        /// <returns>String representation of the realm GUID</returns>
        public static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }
    }

    /// <summary>
    /// A JsonWebSecurityToken generated by SharePoint to authenticate to a 3rd party application and allow callbacks using a refresh token
    /// </summary>
    public class SharePointContextToken : JsonWebSecurityToken
    {
        public static SharePointContextToken Create(JsonWebSecurityToken contextToken)
        {
            return new SharePointContextToken(contextToken.Issuer, contextToken.Audience, contextToken.ValidFrom, contextToken.ValidTo, contextToken.Claims);
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims)
            : base(issuer, audience, validFrom, validTo, claims)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims, SecurityToken issuerToken, JsonWebSecurityToken actorToken)
            : base(issuer, audience, validFrom, validTo, claims, issuerToken, actorToken)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims, SigningCredentials signingCredentials)
            : base(issuer, audience, validFrom, validTo, claims, signingCredentials)
        {
        }

        public string NameId
        {
            get
            {
                return GetClaimValue(this, "nameid");
            }
        }

        /// <summary>
        /// The principal name portion of the context token's "appctxsender" claim
        /// </summary>
        public string TargetPrincipalName
        {
            get
            {
                string appctxsender = GetClaimValue(this, "appctxsender");

                if (appctxsender == null)
                {
                    return null;
                }

                return appctxsender.Split('@')[0];
            }
        }

        /// <summary>
        /// The context token's "refreshtoken" claim
        /// </summary>
        public string RefreshToken
        {
            get
            {
                return GetClaimValue(this, "refreshtoken");
            }
        }

        /// <summary>
        /// The context token's "CacheKey" claim
        /// </summary>
        public string CacheKey
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string cacheKey = (string)dict["CacheKey"];

                return cacheKey;
            }
        }

        /// <summary>
        /// The context token's "SecurityTokenServiceUri" claim
        /// </summary>
        public string SecurityTokenServiceUri
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string cacheKey = (string)dict["SecurityTokenServiceUri"];

                return cacheKey;
            }
        }

        /// <summary>
        /// The realm portion of the context token's "audience" claim
        /// </summary>
        public string Realm
        {
            get
            {
                string aud = Audience;
                if (aud == null)
                {
                    return null;
                }

                string tokenRealm = aud.Substring(aud.IndexOf('@') + 1);

                return tokenRealm;
            }
        }

        private static string GetClaimValue(JsonWebSecurityToken token, string claimType)
        {
            if (token == null)
            {
                throw new ArgumentNullException("token");
            }

            foreach (JsonWebTokenClaim claim in token.Claims)
            {
                if (StringComparer.Ordinal.Equals(claim.ClaimType, claimType))
                {
                    return claim.Value;
                }
            }

            return null;
        }
    }
    /// <summary>
    /// Represents a security token which contains multiple security keys that are generated using symmetric algorithms.
    /// </summary>
    public class MultipleSymmetricKeySecurityToken : SecurityToken
    {
        /// <summary>
        /// Initializes a new instance of the MultipleSymmetricKeySecurityToken class.
        /// </summary>
        /// <param name="keys">An enumeration of Byte arrays that contain the symmetric keys.</param>
        public MultipleSymmetricKeySecurityToken(IEnumerable<byte[]> keys)
            : this(UniqueId.CreateUniqueId(), keys)
        {
        }

        /// <summary>
        /// Initializes a new instance of the MultipleSymmetricKeySecurityToken class.
        /// </summary>
        /// <param name="id">The unique identifier of the security token.</param>
        /// <param name="keys">An enumeration of Byte arrays that contain the symmetric keys.</param>
        public MultipleSymmetricKeySecurityToken(string id, IEnumerable<byte[]> keys)
        {
            if (keys == null)
            {
                throw new ArgumentNullException("keys");
            }

            if (String.IsNullOrEmpty(id))
            {
                throw new ArgumentException("Value cannot be a null or empty string.", "id");
            }

            var enumerable = keys as IList<byte[]> ?? keys.ToList();
            foreach (byte[] key in enumerable)
            {
                if (key.Length == 0)
                {
                    throw new ArgumentException("The key length must be greater then zero.", "keys");
                }
            }

            this.id = id;
            effectiveTime = DateTime.UtcNow;
            securityKeys = CreateSymmetricSecurityKeys(enumerable);
        }

        /// <summary>
        /// Gets the unique identifier of the security token.
        /// </summary>
        public override string Id
        {
            get
            {
                return id;
            }
        }

        /// <summary>
        /// Gets the cryptographic keys associated with the security token.
        /// </summary>
        public override ReadOnlyCollection<SecurityKey> SecurityKeys
        {
            get
            {
                return securityKeys.AsReadOnly();
            }
        }

        /// <summary>
        /// Gets the first instant in time at which this security token is valid.
        /// </summary>
        public override DateTime ValidFrom
        {
            get
            {
                return effectiveTime;
            }
        }

        /// <summary>
        /// Gets the last instant in time at which this security token is valid.
        /// </summary>
        public override DateTime ValidTo
        {
            get
            {
                // Never expire
                return DateTime.MaxValue;
            }
        }

        /// <summary>
        /// Returns a value that indicates whether the key identifier for this instance can be resolved to the specified key identifier.
        /// </summary>
        /// <param name="keyIdentifierClause">A SecurityKeyIdentifierClause to compare to this instance</param>
        /// <returns>true if keyIdentifierClause is a SecurityKeyIdentifierClause and it has the same unique identifier as the Id property; otherwise, false.</returns>
        public override bool MatchesKeyIdentifierClause(SecurityKeyIdentifierClause keyIdentifierClause)
        {
            if (keyIdentifierClause == null)
            {
                throw new ArgumentNullException("keyIdentifierClause");
            }

            // Since this is a symmetric token and we do not have IDs to distinguish tokens, we just check for the
            // presence of a SymmetricIssuerKeyIdentifier. The actual mapping to the issuer takes place later
            // when the key is matched to the issuer.
            if (keyIdentifierClause is SymmetricIssuerKeyIdentifierClause)
            {
                return true;
            }
            return base.MatchesKeyIdentifierClause(keyIdentifierClause);
        }

        private List<SecurityKey> CreateSymmetricSecurityKeys(IEnumerable<byte[]> keys)
        {
            List<SecurityKey> symmetricKeys = new List<SecurityKey>();
            foreach (byte[] key in keys)
            {
                symmetricKeys.Add(new InMemorySymmetricSecurityKey(key));
            }
            return symmetricKeys;
        }

        private string id;
        private DateTime effectiveTime;
        private List<SecurityKey> securityKeys;
    }

    
}
