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

namespace ITAUserProfileManager
{
    public static class GetTermStore
    {
        [FunctionName("GetTermStore")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");
            log.Info(ConfigurationManager.AppSettings["clientID"]);
            string authenticationToken = await helper.getSharePointToken(log);

            try
            {
                using (var clientContext = helper.GetClientContext(ConfigurationManager.AppSettings["tenantURL"], authenticationToken))
                {
                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                    //taxonomySession.UpdateCache();
                    TermStore termStore = taxonomySession.GetDefaultKeywordsTermStore();
                    clientContext.Load(termStore, store => store.Name, store => store.Groups.Where(g => g.IsSystemGroup == false && g.IsSiteCollectionGroup == false).Include(group => group.Id, group => group.Name, group => group.Description, group => group.IsSiteCollectionGroup, group => group.IsSystemGroup, group => group.TermSets.Include(termSet => termSet.Id, termSet => termSet.Name, termSet => termSet.Description, termSet => termSet.CustomProperties, termSet => termSet.Terms.Include(t => t.Id, t => t.Description, t => t.Name, t => t.IsDeprecated, t => t.Parent, t => t.Labels, t => t.LocalCustomProperties, t => t.IsSourceTerm, t => t.IsRoot, t => t.IsKeyword))));

                    clientContext.ExecuteQuery();


                    ITATermStore itaTermStore = new ITATermStore();
                    itaTermStore.ITATermGroupList = new List<ITATermGroup>();
                                        
                    TermGroupCollection allGroups = termStore.Groups;
                    log.Info("Total Groups "+allGroups.Count.ToString());
                    
                    
                    foreach (var termGroup in allGroups)
                    {
                        log.Info("Group Name " + termGroup.Name);
                        
                        ITATermGroup grp = new ITATermGroup();
                        grp.Id = termGroup.Id;
                        grp.Name = termGroup.Name != null ? termGroup.Name : "";
                        //TermGroup specific logic goes here
                        grp.TermSets = new List<ITATermSets>();
                        foreach (var termSet in termGroup.TermSets)
                        {
                            log.Info("Termset Name : " + termSet.Name);
                            ITATermSets itaTermSet = new ITATermSets();
                            itaTermSet.Id = termSet.Id != null ? termSet.Id : new Guid();
                            itaTermSet.Name = termSet.Name != null ? termSet.Name : "";

                            itaTermSet.Terms = new List<ITATerms>();
                            //TermSet specific logic goes here
                            foreach (var term in termSet.Terms)
                            {
                                log.Info("Term Name " + term.Name);
                                ITATerms itaTerm = new ITATerms();
                                itaTerm.Name = term.Name != null ? term.Name : "";
                                itaTerm.Id = term.Id !=null ? term.Id : new Guid();
                                itaTermSet.Terms.Add(itaTerm);
                            }
                            grp.TermSets.Add(itaTermSet);
                        }
                        itaTermStore.ITATermGroupList.Add(grp);
                    }
                    
                    
                    string output = JsonConvert.SerializeObject(itaTermStore);
                    var response = req.CreateResponse(HttpStatusCode.OK);
                    response.Content = new StringContent(output, Encoding.UTF8, "application/json");
                    return response;                    
                }
            }
            catch(Exception ex)
            {
                return req.CreateResponse(HttpStatusCode.OK, "Errorx : "+ ex.Message);
            }
        }
    }
}
