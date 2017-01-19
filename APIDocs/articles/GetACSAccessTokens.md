# Understanding the GetACSAccessTokens Function
The GetACSAccessTokens function is an HttpTrigger function that receives 
**clientId** and **cacheKey** query string parameters and returns a JSON object containing 
userAccessToken and appOnlyAccessToken properties for a valid request or 404 for an invalid request.

The AppLaunch function includes these values on the query string during its final redirect to SharePoint. 

It is recommended, but not required, that at a minimum you augment this function's security 
using [Function Keys](https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-http-webhook#working-with-keys).