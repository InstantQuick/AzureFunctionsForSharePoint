# Understanding the ValidateCredentialToken Function
This function validates a credential token created previously by the [CreateCredentialToken function](./CreateCredentialToken.html).

Invoke this function via HTTP GET with query string values:
* **cId**: The client id of a valid [Client Configuration Guide](ClientConfiguration.md)
* **token**: The token to validate

The token is decrypted and tested for validity by attempting to connect to the SharePoint site indicated by the token. If the validation succeeds or fails,
the function returns the result as application/json as follows:
~~~
{
    "valid":true | false
}
~~~
