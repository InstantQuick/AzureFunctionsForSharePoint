using System;
using System.Configuration;
using System.Web.Script.Serialization;
using IQAppProvisioningBaseClasses.Provisioning;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;

namespace ClientConfiguration
{
    public class Configuration
    {
        private static readonly string ConfigBlobName = "config.json";

        public string ClientId { get; set; }
        public string ProductId { get; set; }
        public string ClientSecret { get; set; }
        public bool AllowAppOnly { get; set; }
        public string ServiceBusConnectionString { get; set; }
        public string NotificationQueueName { get; set; }

        //This is to avoid serialization
        private string _storageAccount;
        private string _storageAccountKey;

        public string GetStorageAccount()
        {
            return _storageAccount;
        }
        public string GetStorageAccountKey()
        {
            return _storageAccountKey;
        }

        public static Configuration GetConfiguration(string clientId)
        {
            return GetConfiguration(clientId, ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]);
        }
        public static Configuration GetConfiguration(string clientId, string storageAccount, string storageAccountKey)
        {
            var containerName = clientId.ToLowerInvariant();

            var container = GetContainer(storageAccount, storageAccountKey, containerName, false);
            if (container == null) return null;
            var configJson = GetConfigJson(container);
            if (string.IsNullOrEmpty(configJson)) return null;
            try
            {
                var configuration = (new JavaScriptSerializer()).Deserialize<Configuration>(configJson);
                configuration._storageAccount = storageAccount;
                configuration._storageAccountKey = storageAccountKey;
                return configuration;
            }
            catch
            {
                return null;
            }
        }

        public static AppManifestBase GetBaseManifest(string clientId, string storageAccount, string storageAccountKey)
        {
            return AppManifestBase.GetManifestFromAzureStorage(storageAccount, storageAccountKey, clientId, "bootstrapmanifest.json");
        }

        public static void SetConfiguration(Configuration config, string storageAccount, string storageKey)
        {
            var containerName = config.ClientId.ToLowerInvariant();
            var container = GetContainer(storageAccount, storageKey, containerName, true);

            if (container == null) throw new Exception("Unable to get or create storage container.");

            var configJson = (new JavaScriptSerializer()).Serialize(config);
            var blob = container.GetBlockBlobReference(ConfigBlobName);
            blob.UploadText(configJson);
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

        private static string GetConfigJson(CloudBlobContainer container)
        {
            var blob = container.GetBlockBlobReference(ConfigBlobName);
            return blob.DownloadText();
        }
    }
}
