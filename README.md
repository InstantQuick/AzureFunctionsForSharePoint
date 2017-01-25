## Recent Changes 
###(Jan 25, 2017)
* Removed from AzureFunctionsForSharePoint.sln the AzureFunctionsForSharePoint.Common.csproj and AzureFunctionsForSharePoint.Core.csproj and gave each its own solution file
* Projects in AzureFunctionsForSharePoint.sln now get AzureFunctionsForSharePoint.Common and AzureFunctionsForSharePoint.Core as packages via nuget
* Fixed issues with Microsoft.IdentityModel references. These are now fulfilled via nuget as dependencies of AzureFunctionsForSharePoint.Common

###(Jan 19, 2017)
* Added the CreateCredentialToken function
* Renamed GetAccessToken to GetACSAccessTokens
* Breaking changes in config file format to support credential clients

**The docs are here: [Azure Functions for SharePoint](https://afspdocs.blob.core.windows.net/docs/index.html)**

