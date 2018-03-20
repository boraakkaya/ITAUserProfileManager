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

namespace ITAUserProfileManager
{
    public static class UpdateUserProfile
    {
        [FunctionName("UpdateUserProfile")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");
            string account = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "account", true) == 0)
                .Value;
            // Get request body
            UserProfile data = await req.Content.ReadAsAsync<UserProfile>();
            string authenticationToken = await helper.getSharePointToken(log);
            if (account != null)
            {
                try
                {
                    using (var clientContext = helper.GetClientContext(ConfigurationManager.AppSettings["tenantURL"], authenticationToken))
                    {
                        PeopleManager peopleManager = new PeopleManager(clientContext);
                        clientContext.Load(peopleManager);
                        if(data.jobTitle != null && data.jobTitle!="")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "Title", data.jobTitle);
                        }
                        if(data.employeeID != null && data.employeeID != "")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "EmployeeID", data.employeeID);
                        }
                        if(data.employeeType != null && data.employeeType != "")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "EmployeeType", data.employeeType);
                        }
                        if(data.workPhone != null && data.workPhone != "")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "WorkPhone", data.workPhone);
                        }
                        if(data.department != null && data.department != "")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "Department", data.department);
                        }

                        if(!string.IsNullOrEmpty(data.officeNumber))
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "OfficeNumber", data.officeNumber);
                        }

                        if(data.manager.email != null && data.manager.email != "")
                        {
                            string managerAccount = "i:0#.f|membership|" + data.manager.email;
                            Microsoft.SharePoint.Client.User mngr = await helper.getSPUser(managerAccount, clientContext, log);
                            peopleManager.SetSingleValueProfileProperty(account, "Manager", mngr.LoginName);
                        }
                        if(data.cellPhone != null && data.cellPhone != "")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "CellPhone", data.cellPhone);
                        }
                        if(data.officeRegion != null && data.officeRegion !="")
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "Office", data.officeRegion);
                        }                        
                        if(!string.IsNullOrEmpty(data.accountExpiration))
                        {
                            peopleManager.SetSingleValueProfileProperty(account, "AccountExpiration",data.accountExpiration.ToString());
                        }
                        if(data.countrySpecialities != null)
                        {
                            List<string> countries = new List<string>();
                            foreach(var country in data.countrySpecialities)
                            {
                                countries.Add(country);
                            }                            
                            peopleManager.SetMultiValuedProfileProperty(account, "CountrySpecialities", countries);
                        }
                        if(data.industrySpecialities != null)
                        {
                            List<string> industries = new List<string>();
                            foreach(var industry in data.industrySpecialities)
                            {
                                industries.Add(industry);
                            }
                            
                            peopleManager.SetMultiValuedProfileProperty(account, "IndustrySpecialities", industries);
                        }
                        
                        if(data.education != null)
                        {
                            string strEducation = "";
                            foreach(Education edu in data.education)
                            {
                                strEducation += "{schoolName:\"" + edu.schoolName + "\",degree:\""+ edu.degree + "\",year:\"" + edu.year +"\"}#;";
                            }
                            peopleManager.SetSingleValueProfileProperty(account, "Education", strEducation);

                        }
                        
                        if(data.certifications != null)
                        {
                            string strCerification = "";
                            foreach(Certification cert in data.certifications)
                            {
                                strCerification += "{organization:\"" + cert.organization + "\",title:\"" + cert.title + "\",year:\"" + cert.year + "\"}#;";
                            }
                            peopleManager.SetSingleValueProfileProperty(account, "Certifications", strCerification);
                        }

                        if(data.emergencyContactInformation != null)
                        {
                            string strContacts = "";
                            foreach(Contact contact in data.emergencyContactInformation)
                            {
                                strContacts += "{firstName:\"" + contact.firstName + "\",lastName:\"" + contact.lastName + "\",phoneNumber:\"" + contact.phoneNumber + "\",emailAddress:\"" + contact.emailAddress +"\"}#;";
                            }
                            peopleManager.SetSingleValueProfileProperty(account, "EmergencyContacts", strContacts);
                        }

                        if(data.mailingAddress != null)
                        {
                            string mailingAddress = data.mailingAddress.addressLine1 + "#;" + data.mailingAddress.addressLine2 + "#;" + data.mailingAddress.city + "#;" + data.mailingAddress.state + "#;" + data.mailingAddress.zipCode + "#;" + data.mailingAddress.country;
                            peopleManager.SetSingleValueProfileProperty(account, "MailingAddress", mailingAddress);
                        }                        
                        clientContext.ExecuteQuery();
                        log.Info("First Name : " + data.firstName + data.lastName);
                        var response = req.CreateResponse(HttpStatusCode.OK);
                        response.Content = new StringContent("{\"status\":\"Success\"}", Encoding.UTF8, "application/json");
                        return response;
                    }
                }
                catch (Exception ex)
                {
                    var response =  req.CreateResponse(HttpStatusCode.OK);
                    response.Content = new StringContent("{\"status\":\"Error\",\"message\":\""+ ex.Message +"\"}", Encoding.UTF8, "application/json");
                    return response;
                }
            }
            else
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, "Please provide valid ITA account");
            }
        }
    }
}
