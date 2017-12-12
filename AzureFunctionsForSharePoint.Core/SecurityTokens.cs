using System;
using System.Configuration;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Web.Script.Serialization;

namespace AzureFunctionsForSharePoint.Core
{
    /// <summary>
    /// The security tokens for a user of an app and methods for reading and writing them in the client's container in Azure storage
    /// </summary>
    public class SecurityTokens
    {
        private static readonly string SecurityTokensBlobName = "tokens.json";

        /// <summary>
        /// The web to which the tokens apply
        /// </summary>
        public string AppWebUrl { get; set; }
        /// <summary>
        /// The host name of the SharePoint tenancy, this is used to access host webs for hybrid solutions with App Webs
        /// </summary>
        public string SPHostName { get; set; }
        /// <summary>
        /// The host part of the app web URL
        /// </summary>
        public string Realm { get; set; }
        /// <summary>
        /// The client to which the tokens apply
        /// </summary>
        public string ClientId { get; set; }
        /// <summary>
        /// Refresh token used to get a new access token
        /// </summary>
        public string RefreshToken { get; set; }
        /// <summary>
        /// Access token for connecting to SP
        /// </summary>
        public string AccessToken { get; set; }
        /// <summary>
        /// Historically unreliable indicator of when the token expires
        /// </summary>
        public DateTime AccessTokenExpires { get; set; }
        
        /// <summary>
        /// Gets the tokens from storage based on the app's config file
        /// </summary>
        /// <param name="cacheKey">The user's cache key</param>
        /// <param name="clientId">The client id</param>
        /// <returns>The tokens</returns>
        public static SecurityTokens GetSecurityTokens(string cacheKey, string clientId)
        {
            return GetSecurityTokens(cacheKey, clientId, ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]);
        }
        /// <summary>
        /// Gets the tokens from storage
        /// </summary>
        /// <param name="cacheKey">The user's cache key</param>
        /// <param name="clientId">The client id</param>
        /// <param name="storageAccount">The storage account</param>
        /// <param name="storageKey">The storage account key</param>
        /// <returns>The tokens</returns>
        public static SecurityTokens GetSecurityTokens(string cacheKey, string clientId, string storageAccount, string storageKey)
        {
            var containerName = clientId.ToLowerInvariant();

            var container = GetContainer(storageAccount, storageKey, containerName, false);
            if (container == null) return null;
            var securityTokenJson = GetSecurityTokens(container, cacheKey);

            return (new JavaScriptSerializer()).Deserialize<SecurityTokens>(securityTokenJson);
        }
        /// <summary>
        /// Saves the tokens to storage based on the app's config file
        /// </summary>
        /// <param name="tokens">The tokens to save</param>
        /// <param name="cacheKey">The user's cache key</param>
        public static void StoreSecurityTokens(SecurityTokens tokens, string cacheKey)
        {
            StoreSecurityTokens(tokens, cacheKey, ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]);
        }
        /// <summary>
        /// Saves the tokens to storage
        /// </summary>
        /// <param name="tokens">The tokens</param>
        /// <param name="cacheKey">The user's cache key</param>
        /// <param name="storageAccount">The storage account</param>
        /// <param name="storageKey">The storage account key</param>
        public static void StoreSecurityTokens(SecurityTokens tokens, string cacheKey, string storageAccount, string storageKey)
        {
            var containerName = tokens.ClientId.ToLowerInvariant();
            var container = GetContainer(storageAccount, storageKey, containerName, true);

            if (container == null) throw new Exception("Unable to get or create storage container.");

            var blob = container.GetBlockBlobReference($"{cacheKey}/{SecurityTokensBlobName}");
            blob.UploadText((new JavaScriptSerializer()).Serialize(tokens));
        }


        private static CloudBlobContainer GetContainer(string storageAccountName, string storageAccountKey,
            string containerName, bool createIfNotExists)
        {
            var connectionString =
                $@"DefaultEndpointsProtocol=https;AccountName={storageAccountName};AccountKey={storageAccountKey}";

            //get a reference to the container where you want to put the files
            var cloudStorageAccount = CloudStorageAccount.Parse(connectionString);
            var cloudBlobClient = cloudStorageAccount.CreateCloudBlobClient();
            var cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
            if (createIfNotExists) cloudBlobContainer.CreateIfNotExists();
            if (!cloudBlobContainer.Exists()) return null;
            return cloudBlobContainer;
        }

        private static string GetSecurityTokens(CloudBlobContainer container, string cacheKey)
        {
            var blob = container.GetBlockBlobReference($"{cacheKey}/{SecurityTokensBlobName}");
            return blob.DownloadText();
        }
    }
}
