using System;
using System.Configuration;
using System.Web.Script.Serialization;
using IQAppProvisioningBaseClasses.Provisioning;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;

namespace AzureFunctionsForSharePoint.Core
{
    /// <summary>
    /// This is the configuration for a client that is registered as an app and uses ACS for authorization
    /// </summary>
    /// <remarks>
    /// These properties ware originally part of ClientConfiguration and this change breaks existing config files  
    /// </remarks>
    public class ACSClientConfig
    {
        /// <summary>
        /// The id of the product. This must match the product id from a SharePoint add-in manifest.
        /// </summary>
        public string ProductId { get; set; } = default(Guid).ToString();

        /// <summary>
        /// The registered client secret of the client
        /// </summary>
        public string ClientSecret { get; set; } = string.Empty;
    }

    /// <summary>
    /// This is the configuration for clients that use real credentials provided by users
    /// at runtime. In these configurations the client and function app share the information
    /// required to encrypt and decrypt the connection information.
    /// </summary>
    /// <seealso cref="System.Security.Cryptography.Rfc2898DeriveBytes"/>
    public class CredentialedClientConfig
    {
        /// <summary>
        /// The password used to derive the encryption key.
        /// </summary>
        public string Password { get; set; } = string.Empty;

        /// <summary>
        /// The key salt used to derive the encryption key.
        /// </summary>
        public string Salt { get; set; } = string.Empty;
    }

    /// <summary>
    /// The configuration of a client and methods to read and store the configuration in Azure storage as JSON
    /// </summary>
    public class ClientConfiguration
    {
        private static readonly string ConfigBlobName = "config.json";

        /// <summary>
        /// The id of the client. If you are using ACS, this should match the client id from a SharePoint add-in manifest.
        /// </summary>
        public string ClientId { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// This is the configuration for a client that is registered as an app and uses ACS for authorization
        /// </summary>
        /// <remarks>
        /// These properties ware originally part of ClientConfiguration and this change breaks existing config files  
        /// </remarks>
        public ACSClientConfig AcsClientConfig { get; set; } = new ACSClientConfig();

        /// <summary>
        /// This is the configuration for clients that use real credentials provided by users
        /// at runtime. In these configurations the client and function app share the information
        /// required to encrypt and decrypt the connection information.
        /// </summary>
        /// <seealso cref="System.Security.Cryptography.Rfc2898DeriveBytes"/>
        public CredentialedClientConfig CredentialedClientConfig { get; set; } = new CredentialedClientConfig();

        /// <summary>
        /// Connection string to the service bus queue the client will use to receive event notifications
        /// </summary>
        public string ServiceBusConnectionString { get; set; } = string.Empty;
        /// <summary>
        /// Name of the queue to which notifications are set
        /// </summary>
        public string NotificationQueueName { get; set; } = string.Empty;

        //This is to avoid serialization
        private string _storageAccount;
        private string _storageAccountKey;

        /// <summary>
        /// Gets the storage account where the config is stored
        /// </summary>
        /// <returns>The Azure storage account name</returns>
        public string GetStorageAccount()
        {
            return _storageAccount;
        }
        /// <summary>
        /// Gets the key to the storage account where the config is stored
        /// </summary>
        /// <returns>The Azure storage account key</returns>
        public string GetStorageAccountKey()
        {
            return _storageAccountKey;
        }

        /// <summary>
        /// Reads the config from Azure storage using the hosting app's ConfigurationManager.AppSettings
        /// </summary>
        /// <param name="clientId">The id of the client. This must match the client id from a SharePoint add-in manifest.</param>
        /// <returns>A ClientConfiguration object</returns>
        public static ClientConfiguration GetConfiguration(string clientId)
        {
            return GetConfiguration(clientId, ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]);
        }
        /// <summary>
        /// Reads the config from Azure storage using the given storage account and account key
        /// </summary>
        /// <param name="clientId">The id of the client. This must match the client id from a SharePoint add-in manifest.</param>
        /// <param name="storageAccount">Azure storage account name</param>
        /// <param name="storageAccountKey">Azure storage account key</param>
        /// <returns>A ClientConfiguration object</returns>
        public static ClientConfiguration GetConfiguration(string clientId, string storageAccount, string storageAccountKey)
        {
            var containerName = clientId.ToLowerInvariant();

            var container = GetContainer(storageAccount, storageAccountKey, containerName, false);
            if (container == null) return null;
            var configJson = GetConfigJson(container);
            if (string.IsNullOrEmpty(configJson)) return null;
            try
            {
                var configuration = (new JavaScriptSerializer()).Deserialize<ClientConfiguration>(configJson);
                configuration._storageAccount = storageAccount;
                configuration._storageAccountKey = storageAccountKey;
                return configuration;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Loads an app manifest named bootstrapmanifest.json from the client config container. If none exists, returns an empty app manifest
        /// </summary>
        /// <param name="clientId">The id of the client. This must match the client id from a SharePoint add-in manifest.</param>
        /// <param name="storageAccount">Azure storage account name</param>
        /// <param name="storageAccountKey">Azure storage account key</param>
        /// <returns>An app manifest</returns>
        public static AppManifestBase GetBootstrapManifest(string clientId, string storageAccount, string storageAccountKey)
        {
            return AppManifestBase.GetManifestFromAzureStorage(storageAccount, storageAccountKey, clientId, "bootstrapmanifest.json");
        }

        /// <summary>
        /// Saves a client config to Azure storage in a container named the same as the client id.
        /// </summary>
        /// <param name="config">A ClientConfiguration instance</param>
        /// <param name="storageAccount">Azure storage account name</param>
        /// <param name="storageKey">Azure storage account key</param>
        public static void SetConfiguration(ClientConfiguration config, string storageAccount, string storageKey)
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
