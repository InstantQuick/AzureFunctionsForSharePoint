#r "AppLaunch.dll"
#r "FunctionsCore.dll"
using System.Net;
using System.Configuration;
using System.Net.Http.Formatting;
using AppLaunch;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    Log(log, $"C# HTTP trigger function processed a request! RequestUri={req.RequestUri}");
    var func = new AppLaunchHandler(req);
    func.FunctionNotify += (sender, args) => Log(log, args.Message);
    
    var appLauncherFunctionArgs = new AppLauncherFunctionArgs()
    {
        StorageAccount = ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
        StorageAccountKey = ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]
    };

    return func.Execute(appLauncherFunctionArgs);
}

public static void Log(TraceWriter log, string message)
{
    log.Info(message);
}
