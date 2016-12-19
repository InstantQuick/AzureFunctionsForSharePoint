#r "AppLaunch.dll"
#r "AzureFunctionsForSharePoint.Common.dll"
using System.Net;
using System.Configuration;
using System.Net.Http.Formatting;
using AppLaunch;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    Log(log, $"C# HTTP trigger function processed a request! RequestUri={req.RequestUri}");
    var func = new AppLaunchHandler(req);
    func.FunctionNotify += (sender, args) => Log(log, args.Message);
    
    var functionArgs = new AzureFunctionArgs()
    {
        StorageAccount = ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
        StorageAccountKey = ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]
    };

    return func.Execute(functionArgs);
}

public static void Log(TraceWriter log, string message)
{
    log.Info(message);
}
