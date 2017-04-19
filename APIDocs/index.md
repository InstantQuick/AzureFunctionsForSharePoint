# Azure Functions for SharePoint
AzureFunctionsForSharePoint is a multi-tenant, multi-addin back-end for SharePoint add-ins built on [Azure Functions](https://azure.microsoft.com/en-us/services/functions/). 
The goal of this project is to provide the minimal set of functions necessary to support the common scenarios shared by most SharePoint provider hosted add-ins cheaply and reliably.

Features include:
* Centralized Identity and ACS token management 
* Installation and provisioning of add-in components to SharePoint
* Remote event dispatching to add-in specific back-end services via message queues including
  * App installation
  * App launch
  * SharePoint Remote Events

## Navigating the Documentation
These documents consist of [articles](articles/intro.html) that explain what the functions do, how to set up the hosting environment, and how to use the functions in your add-ins and [API documentation for .NET developers](api/index.html) linked to the source code in [GitHub](https://github.com/InstantQuick/AzureFunctionsForSharePoint).

## A Note on Terminology
These documents use the term **client** to refer to a given SharePoint add-in. A client is identified using its **client ID** which is the GUID that identifies the add-in's ACS client ID in the [SharePoint add-in's AppManifest.xml](https://msdn.microsoft.com/en-us/library/office/fp179918.aspx#AppManifest).

## Functions
There are five functions in this function app.
  
1. [AppLaunch](articles/AppLaunch.html)
2. [EventDispatch](articles/EventDispatch.html)
3. [GetACSAccessTokens](articles/GetACSAccessTokens.html)
4. [CreateCredentialToken](articles/CreateCredentialToken.html)
5. [ValidateCredentialToken](articles/ValidateCredentialToken.html)

## Setup Guide
The Visual Studio Solution includes a PowerShell script you can use with Task Runner Explorer and [Command Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.CommandTaskRunner).
To set up a new environment from scratch automatically, see [Deploy an Azure Function app using Azure ARM Templates](https://peteskelly.com/deploy-an-azure-function-app-using-azure-arm-templates/)

### Configuring the Function App
AzureFunctionsForSharePoint requires an Azure storage account which stores the configuration of each client as well as the associated tokens. This requires the presence of two app settings and their corresponding values.
* **ConfigurationStorageAccount**
* **ConfigurationStorageAccountKey**

![App Settings](images/Appsettings.png)

### Configuring SharePoint Add-ins to use the Function App
Azure Functions for SharePoint is multi-tenant in that it can service add-ins installed broadly across SharePoint Online 
and also because the back-end processes that respond to client specific events in SharePoint 
or rely on Azure Functions for SharePoint for security token management can be located anywhere with a connection to the Internet. 

See the [Client Configuration Guide](articles/ClientConfiguration.html) for more information. 

### Using the Function App to Support Custom Back-ends
It is possible to use Azure Functions for SharePoint to deliver pure client-side solutions, i.e. HTML/JS. 
However, many add-ins must support scenarios that are difficult or impossible to achieve through pure JavaScript.
Azure Functions for SharePoint supports custom back-ends in two ways:
1. Notification of add-in and SharePoint events via Azure Service Bus queues via the [EventDispatch Function](articles/EventDispatch.html)
2. A REST service that provides security access tokens for registered clients via the [GetACSAccessTokens Function](articles/GetACSAccessTokens.html)

In both cases the client back-end receives all the information it needs to connect to SharePoint 
as either the user or as an app-only identity with full control. 
The function app does the actual authorization flow and its client configuration is the only place where the client secret is stored.

Your custom back-ends can live anywhere 
from the same Function App where you deployed Azure Functions for SharePoint to completely different Azure tenancies or on-premises servers. 
All that is required is that the back-end can read Azure Service Bus Queues and access the REST services via the Internet. 
Aside from these requirements, the back-end can run on any platform and be written in any language.

That said, if you are using .NET, this project included an assembly named [AzureFunctionsForSharePoint.Common](api/AzureFunctionsForSharePoint.Common.html) that you can use to make things even easier!

## API Docs
Complete documentation of the Azure Functions for SharePoint API see the [API Guide](api/index.md).

## Recent Changes (Apr 19, 2017)
* Added the ValidateCredentialToken function
* Updated docs with link to [Deploy an Azure Function app using Azure ARM Templates](https://peteskelly.com/deploy-an-azure-function-app-using-azure-arm-templates/)