using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Configuration;
using System.Text;
using System.Net.Http.Headers;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Collections.Specialized;
using Microsoft.SharePoint.Client.UserProfiles;
using Newtonsoft.Json;
using System;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client;
using System.Net.Http.Formatting;
using Microsoft.SharePoint.Client.Publishing.Navigation;

namespace ITAUserProfileManager
{
    public static class GetNavigation
    {
        [FunctionName("GetNavigation")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string authenticationToken = await helper.getSharePointToken(log, ConfigurationManager.AppSettings["tenantURL"]);
            try
            {
                ITANavigation itaNavigation = new ITANavigation();
                using (var clientContext = helper.GetClientContext(ConfigurationManager.AppSettings["tenantURL"], authenticationToken))
                {
                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);                    
                    TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                    NavigationTermSet navigationTermset;
                    clientContext.Load(termStore, store => store.Groups.Where(g => g.IsSystemGroup == false && g.IsSiteCollectionGroup == false).Include(group => group.TermSets));
                    clientContext.ExecuteQuery();
                    foreach (var group in termStore.Groups)
                    {
                        foreach(var termSet in group.TermSets)
                        {
                            if(termSet.Name == "Offices")
                            {                                
                                List<ITANavigationItem> navigation = new List<ITANavigationItem>();                                
                                navigationTermset = NavigationTermSet.GetAsResolvedByWeb(clientContext, termSet, clientContext.Web, "GlobalNavigationTaxonomyProvider");
                                clientContext.Load(navigationTermset);
                                clientContext.ExecuteQuery();
                                var termCollection = navigationTermset.Terms;
                                clientContext.Load(termCollection,tc=>tc.Include(t=>t.Title,t=>t.SimpleLinkUrl));
                                clientContext.ExecuteQuery();
                                string accessToken = await helper.getSharePointToken(log, "https://itadev.sharepoint.com");
                                foreach (var navTerm in termCollection)
                                {
                                    navigation.Add(new ITANavigationItem()
                                    {
                                        title = navTerm.Title.Value,
                                        link = navTerm.SimpleLinkUrl,
                                        //visible = true,
                                        //target = "_blank",
                                        children = getChildTerms(navTerm, clientContext)
                                        
                                    });                                    
                                }
                                //getStaticLinks(itaNavigation, accessToken);
                                itaNavigation.navigation = navigation;
                            }
                        }
                    }
                }
                var output = JsonConvert.SerializeObject(itaNavigation);
                var response = req.CreateResponse();
                response.StatusCode = HttpStatusCode.OK;
                response.Content = new StringContent(output, Encoding.UTF8, "application/json");
                return response;
            }
            catch(Exception ex)
            {
                return req.CreateResponse(HttpStatusCode.OK, "Error occured : "+ex.Message);
            }
                    
        }
        public static List<ITANavigationItem> getChildTerms(NavigationTerm parent, ClientContext clientContext)
        {
            clientContext.Load(parent, p => p.Terms.Include(t => t.Title, t => t.SimpleLinkUrl));
            clientContext.ExecuteQuery();
            List<ITANavigationItem> children = new List<ITANavigationItem>();
            foreach(var navTerm in parent.Terms)
            {
                children.Add(new ITANavigationItem()
                {
                    title = navTerm.Title.Value,
                    link = navTerm.SimpleLinkUrl,
                    //visible = true,
                    //target = "_blank",
                    children = getChildTerms(navTerm,clientContext)
                });
            }
            return children;
        }
        //public static void getStaticLinks(ITANavigation itaNavigation, string accessToken)
        //{
            
        //    using (var ctx = helper.GetClientContext("https://itadev.sharepoint.com/", accessToken))
        //    {
        //        var web = ctx.Web;
        //        ctx.Load(web);
        //        ctx.ExecuteQuery();
        //        string webTitle = web.Title;
        //        var list = web.Lists.GetByTitle("Mega Menu");
        //        ctx.Load(list);
        //        ctx.ExecuteQuery();
        //        string listTitle = list.Title;
        //        ListItemCollection collection = list.GetItems(new CamlQuery()
        //        {
        //            ViewXml = @"<View><OrderBy><FieldRef Name='Title'/></OrderBy></View>"
        //        });
        //        ctx.Load(collection);
        //        ctx.ExecuteQuery();
        //        List<ITAStaticLink> staticLinks = new List<ITAStaticLink>();
        //        foreach(ListItem listItem in collection)
        //        {

        //            var exists  = staticLinks.Find(a => a.owner == listItem["Owner"].ToString());

        //            if (exists != null)
        //            {
        //                var headingexists = exists.headings.Find(x => x.title == listItem["Heading"].ToString());
        //                if(headingexists != null)
        //                {
        //                    headingexists.links.Add(new ITANavigationItem()
        //                    {
        //                        title = listItem["Title"].ToString(),
        //                        link = listItem["Link"].ToString()
        //                    });
        //                }
        //                else
        //                {
        //                    exists.headings.Add( new ITAStaticLinkHeading() {
        //                        title = listItem["Heading"].ToString(),
        //                        links = new List<ITANavigationItem> { new ITANavigationItem() {
        //                            title =  listItem["Title"].ToString(),
        //                            link = listItem["Link"].ToString()
        //                            }}});
        //                }
        //            }
        //            else
        //            {
        //                staticLinks.Add(new ITAStaticLink()
        //                {
        //                    owner = listItem["Owner"].ToString(),
        //                    headings = new List<ITAStaticLinkHeading> { new ITAStaticLinkHeading() {
        //                        title = listItem["Heading"].ToString(),
        //                        links = new List<ITANavigationItem> { new ITANavigationItem() {
        //                            title =  listItem["Title"].ToString(),
        //                            link = listItem["Link"].ToString()
        //                            }
        //                        }
        //                    } }
        //                });
        //            }
        //        }
        //        itaNavigation.staticLinks = staticLinks;
        //    }
        //}
    }
}
