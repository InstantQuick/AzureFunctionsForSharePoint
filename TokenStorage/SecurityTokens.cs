using System;

namespace TokenStorage
{
    public class SecurityTokens
    {
        public string RefreshToken;
        public string AccessToken;
        public DateTime AccessTokenExpires;
        public string AppWebUrl;
        public string Realm;
        public string ClientId;
    }
}
