using System;
using System.Collections.Specialized;
using System.Net;
using System.Net.Http;
using AzureFunctionsForSharePoint.Common;
using static AzureFunctionsForSharePoint.Core.ClientConfiguration;
using System.Net.Http.Headers;
using System.Web.Script.Serialization;

namespace CreateCredentialToken
{
    /// <summary>
    /// Function specific configuration elements should be added as properties here to extend the <see cref="AzureFunctionArgs" /> class.
    /// </summary>
    public class CreateCredentialTokenerFunctionArgs : AzureFunctionArgs { }

    /// <summary>
    /// This function is called when SharePoint POSTs an ACS token for a SharePoint add-in.
    /// The SharePoint add-in's manifest XML must specify the function URL as the value of the <see href="https://msdn.microsoft.com/en-us/library/office/jj583318.aspx">StartPage element</see>. 
    /// A valid client configuration is required.
    /// 
    /// Once connected to a SharePoint site, the function checks the add-in's install status and provisions as indicated by the bootstrapmanifest.json located in the client's configuration storage container. If provisioning occurs a message is sent to the service bus queue specified in the client configuration to notify the client for additional processing as desired.
    /// Finally,  a message is sent to the service bus queue specified in the client configuration to notify the client of the add-in's launch.
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
                if (request.Content.Headers.ContentType.Equals(new MediaTypeHeaderValue("application/json")))
                {
                    var body = request.Content.ReadAsStringAsync().Result;
                    _credentialedSharePointConnectionInfo =
                        (new JavaScriptSerializer()).Deserialize<CredentialedSharePointConnectionInfo>(body);
                }
            }
            catch
            {
                //ignored
            }
            _response = request.CreateResponse();
        }

        /// <summary>
        /// Performs the app launch flow for the current request
        /// </summary>
        /// <param name="args">An <see cref="CreateCredentialTokenerFunctionArgs"/> instance specifying the location of the client configuration in Azure storage.</param>
        /// <returns>If launch succeeds the response is a 302 redirect back to the SharePoint site's home page.</returns>
        public HttpResponseMessage Execute(CreateCredentialTokenerFunctionArgs args)
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
                _response.Content = new StringContent($"{{'credentialToken':'{token}'}}");
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
