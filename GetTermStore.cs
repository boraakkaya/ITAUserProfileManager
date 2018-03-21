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
                    TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                    clientContext.Load(termStore, store => store.Name, store => store.Groups.Where(g => g.IsSystemGroup == false && g.IsSiteCollectionGroup == false).Include(group => group.Id, group => group.Name, group => group.Description, group => group.IsSiteCollectionGroup, group => group.IsSystemGroup, group => group.TermSets.Include(termSet => termSet.Id, termSet => termSet.Name, termSet => termSet.Description, termSet => termSet.CustomProperties, termSet => termSet.Terms.Include(t => t.Id, t => t.Description, t => t.Name, t => t.IsDeprecated,t => t.Parent, t => t.Labels, t => t.LocalCustomProperties, t => t.IsSourceTerm, t => t.IsRoot, t => t.IsKeyword, t=>t.TermsCount, t=>t.CustomProperties,
                        t => t.Terms.Include(t1 => t1.Id, t1 => t1.Description, t1 => t1.Name, t1 => t1.IsDeprecated, t1 => t1.Parent, t1 => t1.Labels, t1 => t1.LocalCustomProperties, t1 => t1.IsSourceTerm, t1 => t1.IsRoot, t1 => t1.IsKeyword, t1 => t1.TermsCount, t1 => t1.CustomProperties,
                        t1 => t1.Terms.Include(t2 => t2.Id, t2 => t2.Description, t2 => t2.Name, t2 => t2.IsDeprecated, t2 => t2.Parent, t2 => t2.Labels, t2 => t2.LocalCustomProperties, t2 => t2.IsSourceTerm, t2 => t2.IsRoot, t2 => t2.IsKeyword, t2 => t2.TermsCount, t2 => t2.CustomProperties,
                        t2 => t2.Terms.Include(t3 => t3.Id, t3 => t3.Description, t3 => t3.Name, t3 => t3.IsDeprecated, t3 => t3.Parent, t3 => t3.Labels, t3 => t3.LocalCustomProperties, t3 => t3.IsSourceTerm, t3 => t3.IsRoot, t3 => t3.IsKeyword, t3 => t3.TermsCount, t3 => t3.CustomProperties)))
                    ))));

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

                                if (term.Labels.Count > 0)
                                {
                                    itaTerm.Labels = new List<string>();
                                    foreach (var label in term.Labels)
                                    {
                                        itaTerm.Labels.Add(label.Value);
                                    }
                                }
                                if (term.CustomProperties.Count > 0)
                                {
                                    itaTerm.CustomProps = new List<ITATermProperty>();
                                    foreach (var prop in term.CustomProperties)
                                    {
                                        itaTerm.CustomProps.Add(new ITATermProperty() { key = prop.Key, value = prop.Value });
                                    }
                                }

                                getSubTerms(term,itaTerm,log);
                                
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
        public static void getSubTerms(Term parentTerm, ITATerms parentITATerm, TraceWriter log)
        {
            if(parentTerm.TermsCount >0)
            {
                parentITATerm.Terms = new List<ITATerms>();
                

                foreach(var term in parentTerm.Terms)
                {
                    LabelCollection col = term.Labels;
                    
                    ITATerms itaTerm = new ITATerms();
                    itaTerm.Name = term.Name != null ? term.Name : "";
                    itaTerm.Id = term.Id != null ? term.Id : new Guid();
                    
                    if(term.Labels.Count >0)
                    {
                        itaTerm.Labels = new List<string>();
                        foreach(var label in term.Labels)
                        {
                            itaTerm.Labels.Add(label.Value);
                        }
                    }
                    if(term.CustomProperties.Count > 0)
                    {
                        itaTerm.CustomProps = new List<ITATermProperty>();
                        foreach(var prop in term.CustomProperties)
                        {
                            itaTerm.CustomProps.Add(new ITATermProperty() { key = prop.Key, value = prop.Value });
                        }
                    }
                    log.Info(parentITATerm.Name + " ----- " + term.Name);
                    getSubTerms(term,itaTerm,log);
                    parentITATerm.Terms.Add(itaTerm);
                }
            }
        }
    }
    
}
