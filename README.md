**The docs are here: [Azure Functions for SharePoint](https://afspdocs.blob.core.windows.net/docs/index.html)**

## Recent Changes 
**(May 4, 2017)**
* Major changes to the project's organization. The csx files are gone and there is now one assembly for the function host.
* Startup time is substantially improved by moving to precompiled functions
* New function entry point is a new project AzureFunctionsForSharePoint.Host
* The deploy.ps1 file is updated to support the new project structure
* New beta nuget packages for AzureFunctionsForSharePoint.Core and AzureFunctionsForSharePoint.Common
* Individual function implementations consolidated into a new project AzureFunctionsForSharePoint.Functions
* Doc site regenerated and updated

My appologies for dropping this into the main branch. I clearly need to get better at github.

*(Jan 25, 2017)*
* Removed from AzureFunctionsForSharePoint.sln the AzureFunctionsForSharePoint.Common.csproj and AzureFunctionsForSharePoint.Core.csproj and gave each its own solution file
* Projects in AzureFunctionsForSharePoint.sln now get AzureFunctionsForSharePoint.Common and AzureFunctionsForSharePoint.Core as packages via nuget
* Fixed issues with Microsoft.IdentityModel references. These are now fulfilled via nuget as dependencies of AzureFunctionsForSharePoint.Common

*(Jan 19, 2017)*
* Added the CreateCredentialToken function
* Renamed GetAccessToken to GetACSAccessTokens
* Breaking changes in config file format to support credential clients

**The docs are here: [Azure Functions for SharePoint](https://afspdocs.blob.core.windows.net/docs/index.html)**

