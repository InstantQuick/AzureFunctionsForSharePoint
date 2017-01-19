#r "CreateCredentialToken.dll"
#r "AzureFunctionsForSharePoint.Common.dll"
using System.Net;
using System.Configuration;
using System.Net.Http.Formatting;
using CreateCredentialToken;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    Log(log, $"C# HTTP trigger function processed a request! RequestUri={req.RequestUri}");
    var func = new CreateCredentialTokenHandler(req);
    func.FunctionNotify += (sender, args) => Log(log, args.Message);

    var CreateCredentialTokenerFunctionArgs = new CreateCredentialTokenerFunctionArgs()
    {
        StorageAccount = ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
        StorageAccountKey = ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]
    };

    return func.Execute(CreateCredentialTokenerFunctionArgs);
}

public static void Log(TraceWriter log, string message)
{
    log.Info(message);
}
