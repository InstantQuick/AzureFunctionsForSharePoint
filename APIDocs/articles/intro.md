# Azure Functions for SharePoint
AzureFunctionsForSharePoint is a multi-tenant, multi-addin back-end for SharePoint add-ins built on [Azure Functions](https://azure.microsoft.com/en-us/services/functions/). The goal of this project is to provide the minimal set of functions necessary to support the common scenarios shared by most SharePoint provider hosted add-ins cheaply and reliably.

Features include:
* Centralized and Secure Identity and ACS token management
* Provisioning
* Remote event dispatching to add-in specific back-end services via message queues including
  * App installation
  * App launch
  * SharePoint Remote Events

## Navigating the Documentation
These documents consist of [articles](articles/intro.html) that explain what the functions do, how to set up the hosting environment, and how to use the functions in your add-ins and [API documentation for .NET developers](api/index.html) linked to the source code in [GitHub](https://github.com/InstantQuick/AzureFunctionsForSharePoint).

## Functions
There are three functions in this function app.
  
1. [AppLaunch](articles/AppLaunch.html)
2. [EventDispatch](articles/EventDispatch.html)
3. [GetAccessToken](articles/GetAccessToken.html)

## Setup Guide
### Configuring the Function App
### Configuring SharePoint Add-ins to use the Function App
### Using the Function App to Support Custom Back-ends

## API Docs
* [API Guide](api/index.md)
