using System;
using System.Collections.Specialized;
using System.Net;
using System.Net.Http;
using AzureFunctionsForSharePoint.Common;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using System.Net.Http.Headers;
using System.Web.Script.Serialization;

namespace AzureFunctionsForSharePoint.Functions
{
    /// <summary>
    /// Function specific configuration elements should be added as properties here to extend the <see cref="AzureFunctionArgs" /> class.
    /// </summary>
    public class CreateCredentialTokenFunctionArgs : AzureFunctionArgs { }

    /// <summary>
    /// This function is intended to assist in applications where a connection is made to SharePoint using real user credentials instead of ACS or Azure AD.
    /// The JSON is deserialized at runtime into <see cref="CredentialedSharePointConnectionInfo"/> and tested for validity by attempting to connect to the indicated SharePoint site. If the validation succeeds, the function then encrypts the JSON using the client's CredentialedClientConfig in config.json and returns the result as application/json.
    /// 
    /// A valid client configuration is required.
    /// </summary>
    /// <remarks>
    /// This class inherits <see cref="FunctionBase"/> for its simple logging notification event. 
    /// </remarks>
    public class CreateCredentialTokenHandler : FunctionBase
    {
        private readonly CredentialedSharePointConnectionInfo _credentialedSharePointConnectionInfo = null;
        private readonly HttpResponseMessage _response;

        /// <summary>
        /// Initializes the handler for a given HttpRequestMessage received from the function trigger
        /// </summary>
        /// <param name="request">The current request</param>
        public CreateCredentialTokenHandler(HttpRequestMessage request)
        {
            try
            {
                //Inexplicably this test fails when receiving from JavaScript (angular) instead
                //I can clearly see if I debug that the header is application/json, 
                //but the Equals test fails
                //So, just try the operation, the end result is the same either way
                //TODO: Figure this out!

                //if (request.Content.Headers.ContentType.Equals(new MediaTypeHeaderValue("application/json")))
                //{
                var body = request.Content.ReadAsStringAsync().Result;
                _credentialedSharePointConnectionInfo =
                    (new JavaScriptSerializer()).Deserialize<CredentialedSharePointConnectionInfo>(body);
                //}
            }
            catch
            {
                //ignored
            }
            _response = request.CreateResponse();
        }

        /// <summary>
        /// Attempts to connect to SharePoint and creates a <see cref="CredentialedSharePointConnectionInfo"/> returned as application/json
        /// </summary>
        /// <param name="args">An <see cref="CreateCredentialTokenFunctionArgs"/> instance specifying the location of the client configuration in Azure storage.</param>
        public HttpResponseMessage Execute(CreateCredentialTokenFunctionArgs args)
        {
            //If the input is bad in any way there will be an error and the response shall be Unauthorized
            try
            {
                if (_credentialedSharePointConnectionInfo == null)
                {
                    throw new InvalidOperationException("The input didn't have connection info in json format");
                }
                var clientConfig = GetConfiguration(_credentialedSharePointConnectionInfo.ClientId);

                //This will throw if the connection info is no good
                _credentialedSharePointConnectionInfo.GetSharePointClientContext();

                var token = _credentialedSharePointConnectionInfo.GetEncryptedToken(clientConfig.CredentialedClientConfig.Password, clientConfig.CredentialedClientConfig.Salt);

                _response.StatusCode = HttpStatusCode.OK;
                _response.Content = new StringContent($"{{\"credentialToken\":\"{token}\"}}");
                _response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                return _response;
            }
            catch (Exception ex)
            {
                Log($"Error connecting {ex}");
                _response.StatusCode = HttpStatusCode.Unauthorized;
                return _response;
            }
        }
    }
}
