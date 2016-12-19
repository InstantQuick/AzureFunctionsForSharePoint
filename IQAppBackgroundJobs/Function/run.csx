#r "AzureFunctionsForSharePoint.Common.dll"
#r "IQAppBackgroundJobs.dll"
#r "Microsoft.ServiceBus.dll"
using System;
using System.Configuration;
using Microsoft.ServiceBus.Messaging;
using IQAppBackgroundJobs;

public static void Run(BrokeredMessage receivedMessage, TraceWriter log)
{
    log.Info($"C# ServiceBus queue trigger function processed message: {receivedMessage.ContentType}");
    var func = new BackgroundJobHandler();
    func.FunctionNotify += (sender, args) => Log(log, args.Message);

    var appEventFunctionArgs = new BackgroundJobHandlerArgs()
    {
        StorageAccount = ConfigurationManager.AppSettings["ConfigurationStorageAccount"],
        StorageAccountKey = ConfigurationManager.AppSettings["ConfigurationStorageAccountKey"]
    };

    func.Execute(receivedMessage, appEventFunctionArgs);
}

public static void Log(TraceWriter log, string message)
{
    log.Info(message);
}