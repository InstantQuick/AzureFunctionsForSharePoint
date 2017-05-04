using System;
using System.Collections.Generic;
using System.Linq;
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
    public class ValidateCredentialTokenFunctionArgs : AzureFunctionArgs { }

    /// <summary>
    /// This function takes a client ID (cId) and credential token (token) from the query string and validates the token,
    /// first by ensuring it is encrypted properly and then by using the decrypted token to connect to the indicated site.
    /// A valid client configuration is required.
    /// </summary>
    /// <remarks>
    /// This class inherits <see cref="FunctionBase"/> for its simple logging notification event. 
    /// </remarks>
    public class ValidateCredentialTokenHandler : FunctionBase
    {
        private readonly Dictionary<string, string> _queryParams;
        private readonly HttpResponseMessage _response;

        /// <summary>
        /// Initializes the handler for a given HttpRequestMessage received from the function trigger
        /// </summary>
        /// <param name="request">The current request</param>
        public ValidateCredentialTokenHandler(HttpRequestMessage request)
        {
            try
            {
                _queryParams = request.GetQueryNameValuePairs()?
               .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

                _response = request.CreateResponse();
            }
            catch
            {
                //ignored
            }
            _response = request.CreateResponse();
        }

        /// <summary>
        /// Validates the current request
        /// </summary>
        /// <param name="args">An <see cref="ValidateCredentialTokenFunctionArgs"/> instance specifying the location of the client configuration in Azure storage.</param>
        /// <returns>OK with json '{"valid": true || false}'</returns>
        public HttpResponseMessage Execute(ValidateCredentialTokenFunctionArgs args)
        {
            var responseObject = new { valid = false };

            //If the input is bad in any way there will be an error and the response shall be OK with json {"valid":false}
            try
            {
                if (!_queryParams.ContainsKey("token") || !_queryParams.ContainsKey("cId"))
                {
                    throw new Exception();
                }
                var clientConfig = GetConfiguration(_queryParams["cId"]);

                //Will throw if the token can't be decrypted
                var ctx = CredentialedSharePointConnectionInfo.GetSharePointClientContext(_queryParams["token"],
                    clientConfig.CredentialedClientConfig.Password, clientConfig.CredentialedClientConfig.Salt);

                ctx.Load(ctx.Web);

                //Will throw if the credentials are no good
                ctx.ExecuteQuery();
                responseObject = new { valid = true };
            }
            catch
            {
                //ignored - if there is an exception the request is simply not valid
            }
            _response.StatusCode = HttpStatusCode.OK;
            _response.Content = new StringContent((new JavaScriptSerializer()).Serialize(responseObject));
            _response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            return _response;
        }
    }
}
