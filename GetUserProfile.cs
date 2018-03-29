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
    public static class GetUserProfile
    {
        /// ******** Required App Permissions in SharePoint ****
        /// <AppPermissionRequests AllowAppOnlyPolicy="true">
        /// <AppPermissionRequest Scope = "http://sharepoint/content/tenant" Right="FullControl" />
        /// <AppPermissionRequest Scope = "http://sharepoint/social/tenant" Right="FullControl" />
        /// <AppPermissionRequest Scope = "http://sharepoint/content/sitecollection" Right="FullControl" />
        /// </AppPermissionRequests>
        ///  *****************************************************
        /// sample account i:0#.f|membership|bora@tenant.onmicrosoft.com
        [FunctionName("GetUserProfile")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            log.Info("Message is " , ConfigurationManager.AppSettings["Message"]);
            string message = ConfigurationManager.AppSettings["Message"];
            // parse query parameter
            string account = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "account", true) == 0)
                .Value;
            dynamic data = await req.Content.ReadAsAsync<object>();
            
            account = account ?? data?.account;
            string authenticationToken = await helper.getSharePointToken(log, ConfigurationManager.AppSettings["tenantURL"]);            
            if (account != null)
            {
                try
                {
                    using (var clientContext = helper.GetClientContext(ConfigurationManager.AppSettings["tenantURL"], authenticationToken))
                    {
                        PeopleManager peopleManager = new PeopleManager(clientContext);
                        PersonProperties personProperties = peopleManager.GetPropertiesFor(account);
                        clientContext.Load(personProperties);
                        clientContext.ExecuteQuery();

                        UserProfile profile = new UserProfile();
                        profile.firstName = personProperties.UserProfileProperties["FirstName"].ToString();
                        profile.middleName = "";
                        profile.lastName = personProperties.UserProfileProperties["LastName"].ToString();
                        profile.suffix = personProperties.UserProfileProperties["Suffix"].ToString();
                        profile.workPhone = personProperties.UserProfileProperties["WorkPhone"].ToString();

                        profile.jobTitle = personProperties.UserProfileProperties["Title"].ToString();
                        profile.department = personProperties.UserProfileProperties["Department"].ToString();

                        if(personProperties.UserProfileProperties["TaxonomyDepartment"] != null)
                        {
                            profile.taxonomyDepartment = personProperties.UserProfileProperties["TaxonomyDepartment"].ToString();
                        }
                        profile.cellPhone = personProperties.UserProfileProperties["CellPhone"].ToString();
                        profile.emailAddress = personProperties.UserProfileProperties["WorkEmail"].ToString();
                        profile.officeRegion = "OfficeRegion";//personProperties.UserProfileProperties["OfficeRegion"].ToString();
                        profile.officeCountry = "OfficeCountry"; //personProperties.UserProfileProperties[""].ToString();
                        profile.officeState = "OfficeState"; //personProperties.UserProfileProperties[""].ToString();

                        MailingAddress userAddress = new MailingAddress() { addressLine1 = "", addressLine2 = "", city = "", state = "", zipCode = "", country = "" };
                        if (personProperties.UserProfileProperties["MailingAddress"].ToString() != "")
                        {
                            string[] mailingAddressSections = personProperties.UserProfileProperties["MailingAddress"].ToString().Split(new string[] { "#;" }, StringSplitOptions.None);
                            log.Info("Mailing Address Sections " + mailingAddressSections[0].ToString());
                            if (mailingAddressSections.Length > 0)
                            {
                                userAddress = new MailingAddress() { addressLine1 = mailingAddressSections[0], addressLine2 = mailingAddressSections[1], city = mailingAddressSections[2], state = mailingAddressSections[3], zipCode = mailingAddressSections[4], country = mailingAddressSections[5] };
                            }
                        }
                        profile.mailingAddress = userAddress;
                        profile.officeNumber = personProperties.UserProfileProperties["OfficeNumber"].ToString();
                        User managerObject = new User() { displayName = "", email = "" };
                        if (personProperties.UserProfileProperties["Manager"].ToString() != "")
                        {
                            managerObject = await helper.getUser(personProperties.UserProfileProperties["Manager"].ToString(), clientContext, log);
                        }
                        profile.manager = managerObject;
                        profile.employeeID = personProperties.UserProfileProperties["EmployeeID"].ToString();
                        profile.employeeType = personProperties.UserProfileProperties["EmployeeType"].ToString();

                        if (!string.IsNullOrEmpty(personProperties.UserProfileProperties["AccountExpiration"].ToString()))
                        {
                            profile.accountExpiration = personProperties.UserProfileProperties["AccountExpiration"].ToString();
                        }
                        string[] countrySpecialitiesArray = personProperties.UserProfileProperties["CountrySpecialities"].ToString().Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        profile.countrySpecialities = countrySpecialitiesArray;
                        string[] industrySpecialtiesArray = personProperties.UserProfileProperties["IndustrySpecialities"].ToString().Split(new char[] {'|'},StringSplitOptions.RemoveEmptyEntries);
                        profile.industrySpecialities = industrySpecialtiesArray;
                        if (personProperties.UserProfileProperties["CSATCompletion"] != null && personProperties.UserProfileProperties["CSATCompletion"].ToString() != "")
                        {
                            profile.CSATCompletion = Convert.ToDateTime(personProperties.UserProfileProperties["CSATCompletion"]);
                        }
                        else
                        {
                            profile.CSATCompletion = null;
                        }
                        List<Contact> emergencyContacts = new List<Contact> { };
                        string contactsProfileServiceValue = personProperties.UserProfileProperties["EmergencyContacts"].ToString();
                        if (contactsProfileServiceValue.Length > 0)
                        {
                            string[] contactsArray = contactsProfileServiceValue.Split(new string[] { "#;" }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var obj in contactsArray)
                            {
                                JObject jobj = JObject.Parse(obj);
                                emergencyContacts.Add(new Contact() { firstName = jobj.Value<string>("firstName"), lastName = jobj.Value<string>("lastName"), emailAddress = jobj.Value<string>("emailAddress"), phoneNumber = jobj.Value<string>("phoneNumber") });
                            }
                        }
                        profile.emergencyContactInformation = emergencyContacts;

                        List<Certification> certifications = new List<Certification> { };
                        string certificationsProfileServiceValue = personProperties.UserProfileProperties["Certifications"].ToString();
                        if (certificationsProfileServiceValue.Length > 0)
                        {
                            string[] certificationsArray = certificationsProfileServiceValue.Split(new string[] { "#;" }, StringSplitOptions.RemoveEmptyEntries);
                            foreach(var obj in certificationsArray)
                            {
                                JObject jobj = JObject.Parse(obj);
                                certifications.Add(new Certification() { title = jobj.Value<string>("title"), organization = jobj.Value<string>("organization"), year = jobj.Value<string>("year")});
                            }
                        }
                        profile.certifications = certifications;

                        List<Education> education = new List<Education> { };
                        string educationsProfileServiceValue = personProperties.UserProfileProperties["Education"].ToString();
                        if (educationsProfileServiceValue.Length > 0)
                        {
                            string[] educationsArray = educationsProfileServiceValue.Split(new string[] { "#;" }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var obj in educationsArray)
                            {
                                log.Info(educationsArray + obj);
                                JObject jobj = JObject.Parse(obj);
                                education.Add(new Education() { schoolName = jobj.Value<string>("schoolName"), degree = jobj.Value<string>("degree"), year = jobj.Value<string>("year") });
                            }
                        }
                        profile.education = education;

                        List<User> directReports = new List<User> { };
                        foreach(var obj in personProperties.DirectReports)
                        {
                            User user = await helper.getUser(obj,clientContext,log);
                            directReports.Add(user);
                        }
                        profile.directReports = directReports;

                        string output = JsonConvert.SerializeObject(profile);                        
                        var response = req.CreateResponse(HttpStatusCode.OK);
                        response.Content = new StringContent(output, Encoding.UTF8, "application/json");
                        return response;                       
                    }
                }
                catch(Exception ex)
                {
                    return req.CreateResponse(HttpStatusCode.OK, ex.Message);
                   
                }
            }
            else
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, "Please provide valid ITA account");
            }
            
        }
        


    }
}
