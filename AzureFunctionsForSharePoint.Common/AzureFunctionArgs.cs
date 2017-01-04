namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// This is the base class for function input. It includes Azure storage connection info used to access client configurations and security tokens. 
    /// </summary>
    /// <remarks>
    /// This class is virtual because using it directly from the csx files creates assembly load conflicts. Therefore you are required to implement by sub-classing it for each function.
    /// </remarks>
    public abstract class AzureFunctionArgs
    {
        /// <summary>
        /// Name of the Azure storage account
        /// </summary>
        public virtual string StorageAccount { get; set; }
        /// <summary>
        /// Access key to the Azure storage account
        /// </summary>
        public virtual string StorageAccountKey { get; set; }
    }
}
