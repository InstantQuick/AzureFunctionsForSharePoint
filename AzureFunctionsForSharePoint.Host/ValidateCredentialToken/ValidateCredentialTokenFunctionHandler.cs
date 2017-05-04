using System.Configuration;
using System.Threading.Tasks;
using System.Net.Http;
using Microsoft.Azure.WebJobs.Host;
using AzureFunctionsForSharePoint.Functions;

namespace AzureFunctionsForSharePoint.Host
{
    /// <summary>
    /// Entry point to the actual functionality for an Azure function host
    /// This class passes the input from the trigger to the function and logs notification events
    /// There are two assemblies to allow hosting the function code in another container without a dependency of the WebJobs SDK
    /// </summary>
    /// <remarks>
    /// If you want to test a locally hosted version, follow the instructions here: https://github.com/lindydonna/FunctionsAsWebProject
    /// </remarks>
    public class ValidateCredentialTokenFunctionHandler
    {
        public static Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
        {
            Log(log, $"C# HTTP trigger function processed a request! RequestUri={req.RequestUri}");
            var func = new ValidateCredentialTokenHandler(req);
            func.FunctionNotify += (sender, args) => Log(log, args.Message);

            var ValidateCredentialTokenFunctionArgs = new ValidateCredentialTokenFunctionArgs()
            {
                StorageAccount = ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
                StorageAccountKey = ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]
            };

            return Task.FromResult(func.Execute(ValidateCredentialTokenFunctionArgs));
        }

        public static void Log(TraceWriter log, string message)
        {
            log.Info(message);
        }
    }
}