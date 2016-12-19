using System;
using System.Configuration;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Web.Script.Serialization;

namespace AzureFunctionsForSharePoint.Core
{
    public class SecurityTokens
    {
        private static readonly string SecurityTokensBlobName = "tokens.json";

        public string RefreshToken;
        public string AccessToken;
        public DateTime AccessTokenExpires;
        public string AppWebUrl;
        public string Realm;
        public string ClientId;

        public static SecurityTokens GetSecurityTokens(string cacheKey, string clientId)
        {
            return GetSecurityTokens(cacheKey, clientId, ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]);
        }
        public static SecurityTokens GetSecurityTokens(string cacheKey, string clientId, string storageAccount, string storageKey)
        {
            var containerName = clientId.ToLowerInvariant();

            var container = GetContainer(storageAccount, storageKey, containerName, false);
            if (container == null) return null;
            var securityTokenJson = GetSecurityTokens(container, cacheKey);

            return (new JavaScriptSerializer()).Deserialize<SecurityTokens>(securityTokenJson);
        }

        public static void StoreSecurityTokens(SecurityTokens tokens, string cacheKey)
        {
            StoreSecurityTokens(tokens, cacheKey, ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]);
        }
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
