using System.Security;
using System.Web.Script.Serialization;
using Microsoft.SharePoint.Client;

namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Information used to connect to SharePoint
    /// This class is converted to/from JSON and the resulting JSON is encrypted to create the token
    /// </summary>
    public class CredentialedSharePointConnectionInfo
    {
        /// <summary>
        /// Identifies the container where the config file with the encryption password and salt
        /// </summary>
        public string ClientId { get; set; } = string.Empty;
        /// <summary>
        /// The SharePoint site to which this connects
        /// </summary>
        public string SiteUrl { get; set; } = string.Empty;
        /// <summary>
        /// The username login name
        /// </summary>
        public string UserName { get; set; } = string.Empty;
        /// <summary>
        /// The password
        /// </summary>
        public string Password { get; set; } = string.Empty;

        public string GetEncryptedToken(string encryptionPassword, string salt)
        {
            return Encryption.Encrypt((new JavaScriptSerializer()).Serialize(this), encryptionPassword, salt);
        }

        public ClientContext GetSharePointClientContext()
        {
            var clientContext = new ClientContext(SiteUrl);

            var securePassword = new SecureString();
            foreach (char c in Password.ToCharArray()) securePassword.AppendChar(c);
            var credentials = new SharePointOnlineCredentials(UserName, securePassword);

            clientContext.Credentials = credentials;
            clientContext.Load(clientContext.Web, w => w.Title);

            //This will throw if the connection info is no good
            clientContext.ExecuteQuery();
            return clientContext;
        }

        public static ClientContext GetSharePointClientContext(string encryptedToken, string encryptionPassword,
            string salt)
        {
            var connectionInfo = GetFromEncryptedToken(encryptedToken, encryptionPassword, salt);
            var clientContext = new ClientContext(connectionInfo.SiteUrl);

            var securePassword = new SecureString();
            foreach (char c in connectionInfo.Password) securePassword.AppendChar(c);
            var credentials = new SharePointOnlineCredentials(connectionInfo.UserName, securePassword);

            clientContext.Credentials = credentials;
            clientContext.Load(clientContext.Web, w => w.Title);

            //This will throw if the connection info is no good
            clientContext.ExecuteQuery();
            return clientContext;
        }

        private static CredentialedSharePointConnectionInfo GetFromEncryptedToken(string encryptedToken, string encryptionPassword, string salt)
        {
            var json = Encryption.Decrypt(encryptedToken, encryptionPassword, salt);
            return (CredentialedSharePointConnectionInfo)(new JavaScriptSerializer()).Deserialize(json, typeof(CredentialedSharePointConnectionInfo));
        }
    }
}
