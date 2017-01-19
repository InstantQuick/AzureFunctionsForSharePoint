# Understanding the CreateCredentialToken Function
This function is intended to assist in applications where a connection is made 
to SharePoint using real user credentials instead of ACS or Azure AD. The function receives a POST
of application/json as follows:
~~~
{
	"ClientId":"[YOUR_CLIENT_ID]",
	"UserName":"[YOUR_USER_ID]",
	"Password":"[YOUR_PASSWORD]",
	"SiteUrl":"[YOUR_SP_SITE_URL]"
}
~~~

The JSON is deserialized at runtime into [CredentialedSharePointConnectionInfo](../api/AzureFunctionsForSharePoint.Common.CredentialedSharePointConnectionInfo.html) and
tested for validity by attempting to connect to the indicated SharePoint site. If the validation succeeds,
the function then encrypts the JSON using the client's [CredentialedClientConfig](../api/AzureFunctionsForSharePoint.Core.CredentialedClientConfig.html) in config.json
and returns the result as application/json as follows:
~~~
{
    "credentialToken":"[RESULT]"
}
~~~

Keep the encryption password and salt safe, and share it only with the client that will 
consume the connection information. If the client application is based on .NET, the client can install
the [AzureFunctionsForSharePoint.Common nuget package](https://www.nuget.org/packages/AzureFunctionsForSharePoint.Common/)
and use the **GetSharePointClientContext** method in 
[AzureFunctionsForSharePoint.Common.CredentialedSharePointConnectionInfo](../api/AzureFunctionsForSharePoint.Common.CredentialedSharePointConnectionInfo.html).




