using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.CustomerEngagement.Samples
{
    /// <summary>
    /// Provides the D365 web service and Azure app registration information read
    /// from the App.config file in this project.
    /// </summary>
    /// <remarks>You must provide your own values for the app settings in the App.config
    /// file before running this sample.</remarks>
    class WebApiConfiguration
    {
        public string ClientId { get; set; }
        public string Secret { get; set; }
        public string TenantId { get; set; }
        public string ResourceUri { get; set; }
        public string ServiceRoot { get; set; }

        public WebApiConfiguration()
        {
            var appSettings = ConfigurationManager.AppSettings;

            ClientId = appSettings["Client-ID"];
            Secret = appSettings["Client-Secret"];
            TenantId = appSettings["Tenant-ID"];
            ResourceUri = appSettings["Resource-URL"];
            ServiceRoot = appSettings["Service-Root"];
        }
    }

    /// <summary>
    /// Single tenant service-to-service (S2S) sample. This sample makes use of an
    /// app registration in Azure to access a D365 server using WebAPI calls without
    /// requiring a user's logon credentials.
    /// </summary>
    class SingleTenantS2S
    {
        static void Main(string[] args)
        {
            // Obtain the app registration and service configuration values from the App.config file.
            var webConfig = new WebApiConfiguration();

            var account1 = new JObject
                    {
                          {"name", "WEB MARKETING"},
                          {"crm_pravnioblik", 400000000},
                          {"crm_oib", "25065980939"},
                          {"crm_mb", "02085429"},
                          {"crm_sifradjelatnosti", "6201"},
                          {"address1_city", "Zagreb"},
                          {"address1_line1", "Medarska 56b"},
                          {"address1_postofficebox", "10090"},
                          {"telephone1", "3817-757"},
                          {"fax", "3864-105"},
                          {"websiteurl", "www.wem.hr"},
                          {"crm_udruga_sifra", "033"},
                          {"crm_brojdjelatnika", 10},
                          {"crm_regija@odata.bind", "/teams(353FC4D2-9EE6-E911-A829-000D3ABA5555)"}};

            var accUri = PostAsync("accounts", account1, webConfig).GetAwaiter().GetResult();

            var getAcc = GetAsync(accUri.ToString(), webConfig).GetAwaiter().GetResult();

            //2 (success 204, checked in crm it was created by app )
            var contact1 = new JObject
                        {
                         {"firstname", "Hrvoje"},
                          { "lastname", "Nekić"},
                          { "gendercode", 1},
                          { "jobtitle", "direktor"},
                          { "telephone1", "013817757"},
                          { "fax", "013864105"},
                          { "mobilephone", "0912501660"},
                          { "emailaddress1", "hrvoje@wem.hr"},
                          { "address1_city", "Zagreb"},
                          { "address1_line1", "Medarska 56b"},
                          { "address1_postofficebox", "10090"},
                          { "donotemail", false}};
            var contactUri = PostAsync("contacts", contact1, webConfig).GetAwaiter().GetResult();

            var contact = GetAsync(contactUri.ToString(), webConfig).GetAwaiter().GetResult();

            //SingleTenantS2S demo code

            // Send a WebAPI message request for the top 3 account names.
            //var response = SendMessageAsync(webConfig, HttpMethod.Get,
            //    webConfig.ServiceRoot + "accounts?$select=name&$top=3").Result;

            // Format and then output the JSON response to the console.
            //if (response.IsSuccessStatusCode)
            //{
            //    JObject body = JObject.Parse(response.Content.ReadAsStringAsync().Result);
            //    Console.WriteLine(body.ToString());
            //}
            //else
            //{
            //    Console.WriteLine("The request failed with a status of '{0}'",
            //           response.ReasonPhrase);
            //}
        }

        /// <summary>
        /// Send a message via HTTP.
        /// </summary>
        /// <param name="webConfig">A WebAPI configuration.</param>
        /// <param name="httpMethod">The HTTP method to use with the message.</param>
        /// <param name="messageUri">The URI of the WebAPI endpoint plus ODATA parameters.</param>
        /// <param name="body">The message body; otherwise, null.</param>
        /// <returns></returns>
        public static async Task<HttpResponseMessage> SendMessageAsync(WebApiConfiguration webConfig,
            HttpMethod httpMethod, string messageUri, string body = null)
        {
            // Get the access token that is required for authentication.
            var accessToken = await GetAccessToken(webConfig);

            // Create an HTTP message with the required WebAPI headers populated.
            var client = new HttpClient();
            var message = new HttpRequestMessage(httpMethod, messageUri);

            message.Headers.Add("OData-MaxVersion", "4.0");
            message.Headers.Add("OData-Version", "4.0");
            message.Headers.Add("Prefer", "odata.include-annotations=*");
            message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Add any body content specified in the passed parameter.
            if (body != null)
                message.Content = new StringContent(body, UnicodeEncoding.UTF8, "application/json");

            // Send the message to the WebAPI.
            return await client.SendAsync(message);
        }

        /// <summary>
        /// Posts a payload to the specified resource asynchronously.
        /// </summary>
        /// <param name="path">The path to the resource.</param>
        /// <param name="body">The payload to send.</param>
        /// <param name="headers">Any headers to control optional behaviors.</param>
        /// <returns>The response from the request.</returns>
        public static async Task<Uri> PostAsync(string path, object body, WebApiConfiguration webConfig)
        {
            try
            {
                HttpClient httpClient = new HttpClient();
                httpClient.BaseAddress = new Uri(webConfig.ServiceRoot);

                httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
                httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
                httpClient.DefaultRequestHeaders.Add("Prefer", "odata.include-annotations=*");
                httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));

                // Get the access token that is required for authentication.
                var accessToken = await GetAccessToken(webConfig);

                using (var message = new HttpRequestMessage(HttpMethod.Post, path))
                {
                    message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    if (body != null)
                        message.Content = new StringContent(JObject.FromObject(body).ToString(), UnicodeEncoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await httpClient.SendAsync(message, HttpCompletionOption.ResponseHeadersRead))
                    {
                        return new Uri(response.Headers.GetValues("OData-EntityId").FirstOrDefault());
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Retrieves data from a specified resource asychronously.
        /// </summary>
        /// <param name="path">The path to the resource.</param>
        /// <param name="headers">Any custom headers to control optional behaviors.</param>
        /// <returns>The response to the request.</returns>
        public static async Task<JToken> GetAsync(string path, WebApiConfiguration webConfig, Dictionary<string, List<string>> headers = null)
        {
            try
            {
                HttpClient httpClient = new HttpClient();
                httpClient.BaseAddress = new Uri(webConfig.ServiceRoot);

                httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
                httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
                httpClient.DefaultRequestHeaders.Add("Prefer", "odata.include-annotations=*");
                httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));

                // Get the access token that is required for authentication.
                var accessToken = await GetAccessToken(webConfig);

                using (var message = new HttpRequestMessage(HttpMethod.Get, path))
                {
                    message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    if (headers != null)
                    {
                        foreach (KeyValuePair<string, List<string>> header in headers)
                        {
                            message.Headers.Add(header.Key, header.Value);
                        }
                    }

                    using (HttpResponseMessage response = await httpClient.SendAsync(message, HttpCompletionOption.ResponseContentRead))
                    {
                        if (response.StatusCode != HttpStatusCode.NotModified)
                        {
                            return JToken.Parse(await response.Content.ReadAsStringAsync());
                        }
                        return null;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Updates a property of an entity asychronously
        /// </summary>
        /// <param name="path">The path to the entity.</param>
        /// <param name="property">The name of the property to update.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>Task</returns>
        public async Task PutAsync(string path, string property, string value, WebApiConfiguration webConfig)
        {
            try
            {
                HttpClient httpClient = new HttpClient();
                httpClient.BaseAddress = new Uri(webConfig.ServiceRoot);

                httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
                httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
                httpClient.DefaultRequestHeaders.Add("Prefer", "odata.include-annotations=*");
                httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));

                // Get the access token that is required for authentication.
                var accessToken = await GetAccessToken(webConfig);

                using (var message = new HttpRequestMessage(HttpMethod.Put, new Uri($"{path}/{property}")))
                {
                    message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var body = new JObject
                    {
                        ["value"] = value
                    };
                    message.Content = new StringContent(body.ToString());
                    message.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                    HttpResponseMessage response = await httpClient.SendAsync(message);
                    response.Dispose();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get the authentication access token.
        /// </summary>
        /// <param name="webConfig">The WebAPI configuration.</param>
        /// <returns></returns>
        public static async Task<string> GetAccessToken(WebApiConfiguration webConfig)
        {
            var credentials = new ClientCredential(webConfig.ClientId, webConfig.Secret);
            var authContext = new AuthenticationContext(
                "https://login.microsoftonline.com/" + webConfig.TenantId);
            var result = await authContext.AcquireTokenAsync(webConfig.ResourceUri, credentials);

            return result.AccessToken;
        }
    }
}