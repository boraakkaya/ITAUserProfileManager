using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ITAUserProfileManager
{
    class helper
    {
        static string clientID = ConfigurationManager.AppSettings["clientID"];
        static string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
        static string tenantURL = ConfigurationManager.AppSettings["tenantURL"];
        static string tenantID = ConfigurationManager.AppSettings["tenantID"];
        static string spPrinciple = ConfigurationManager.AppSettings["spPrinciple"];
        static string spAuthUrl = ConfigurationManager.AppSettings["spAuthUrl"];
        public string TenantURL { get => tenantURL; }
                
        public static ClientContext GetClientContext(string siteUrl, string accessToken)
        {   
            var ctx = new ClientContext(siteUrl);
            ctx.ExecutingWebRequest += (s, e) =>
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };
            return ctx;
        }
        public static async Task<string> getSharePointToken(TraceWriter log)
        {
            HttpClient client = new HttpClient();
            KeyValuePair<string, string>[] body = new KeyValuePair<string, string>[]
            {
        new KeyValuePair<string, string>("grant_type", "client_credentials"),
        new KeyValuePair<string, string>("client_id", $"{clientID}@{tenantID}"),
        new KeyValuePair<string, string>("resource", $"{spPrinciple}/{tenantURL}@{tenantID}".Replace("https://", "")),
        new KeyValuePair<string, string>("client_secret", clientSecret)
            };
            var content = new FormUrlEncodedContent(body);
            var contentLength = content.ToString().Length;
            string token = "";
            using (HttpResponseMessage response = await client.PostAsync(spAuthUrl, content))
            {
                if (response.Content != null)
                {
                    string responseString = await response.Content.ReadAsStringAsync();                    
                    JObject data = JObject.Parse(responseString);             
                    token = data.Value<string>("access_token");
                }
            }           
            return token;
        }

        public static async Task<User> getUser(string account, ClientContext ctx, TraceWriter log)
        {
            string userEmail = account.Split('|')[2];
            Microsoft.SharePoint.Client.User user = ctx.Web.SiteUsers.GetByEmail(userEmail);
            ctx.Load(user);
            ctx.ExecuteQuery();
            return new User() { displayName = user.Title, email = user.Email };            
        }
        public static async Task<Microsoft.SharePoint.Client.User> getSPUser(string account, ClientContext ctx, TraceWriter log)
        {
            string userEmail = account.Split('|')[2];
            Microsoft.SharePoint.Client.User user = ctx.Web.SiteUsers.GetByEmail(userEmail);
            ctx.Load(user);
            ctx.ExecuteQuery();
            return user;
        }

    }
}
