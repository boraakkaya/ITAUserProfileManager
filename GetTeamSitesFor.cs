using System;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;

namespace ITAUserProfileManager
{
    public static class GetTeamSitesFor
    {
        [FunctionName("GetTeamSitesFor")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string account = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "account", true) == 0).Value;            
            log.Info("User  : " + account);
            var graphAuthenticationContext = new AuthenticationContext("https://login.microsoftonline.com/"+ ConfigurationManager.AppSettings["tenantID"] +"/oauth2/authorize", false);
            //Pass ClientID and ClientSecret for app credentials
            ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["clientID"], ConfigurationManager.AppSettings["clientSecret"]);
            AuthenticationResult graphAuthenticationResult = await graphAuthenticationContext.AcquireTokenAsync("https://graph.microsoft.com", clientCred);
            string graphToken = graphAuthenticationResult.AccessToken;
            log.Info("Graph Token : " + graphToken);
            var graphResponse = await postToGraph(graphToken, account, log);
            var graphResponseResult = await graphResponse.Content.ReadAsStringAsync();
            log.Info("Graph Response Result : " + graphResponseResult);
            log.Info("Graph Response Status Code  : " + graphResponse.StatusCode.ToString());

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Content = new StringContent(graphResponseResult, Encoding.UTF8, "application/json");
            return response;
        }

        public static async Task<HttpResponseMessage> postToGraph(string graphToken, string account, TraceWriter log)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphToken);
            string resource = "https://graph.microsoft.com/beta/users/" + account + "/memberof";
            var requestUrl = resource;
            var requestMethod = new HttpMethod("GET");
            var request = new HttpRequestMessage(requestMethod, requestUrl);           
            HttpResponseMessage hrm = await client.SendAsync(request);
            var reqResponseResult = await hrm.Content.ReadAsStringAsync();
            log.Info("In Async : " + reqResponseResult);
            return hrm;
        }

    }
}
