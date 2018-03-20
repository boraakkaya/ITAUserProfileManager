using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITAUserProfileManager
{
    /// <summary>
    /// 
    /// 
//    User Profile : 
//UserProfile_GUID - c4661c80-04c8-42f1-a6d2-c4aae2f9ed13
//SID - i:0h.f|membership|10037ffea914261a @live.com
//ADGuid - System.Byte[]
//AccountName - i:0#.f|membership|bora@itadev.onmicrosoft.com
//FirstName - Bora
//SPS-PhoneticFirstName - 
//LastName - Akkaya
//SPS-PhoneticLastName - 
//PreferredName - Bora Akkaya
//SPS-PhoneticDisplayName - 
//WorkPhone - 123-456-7788
//Department - OCIO-PD-PS
//Title - SharePoint Consultant
//SPS-Department - 
//Manager - i:0#.f|membership|michael@itadev.onmicrosoft.com
//AboutMe - 
//PersonalSpace - 
//PictureURL - https://itadev-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/bora_itadev_onmicrosoft_com_MThumb.jpg
//UserName - bora @itadev.onmicrosoft.com
//QuickLinks - 
//WebSite - 
//PublicSiteRedirect - 
//SPS-JobTitle - 
//SPS-Dotted-line - 
//SPS-Peers - 
//SPS-Responsibility - 
//SPS-SipAddress - bora @itadev.onmicrosoft.com
//SPS-MySiteUpgrade - 
//SPS-ProxyAddresses - 
//SPS-HireDate - 
//SPS-DisplayOrder - 
//SPS-ClaimID - bora @itadev.onmicrosoft.com
//SPS-ClaimProviderID - membership
//SPS-ResourceSID - 
//SPS-ResourceAccountName - 
//SPS-MasterAccountName - 
//SPS-UserPrincipalName - bora @itadev.onmicrosoft.com
//SPS-O15FirstRunExperience - 
//SPS-PersonalSiteInstantiationState - 
//SPS-DistinguishedName - CN= 5a90cea2-f71c-4c35-a42c-e32ef3c0e0d8, OU= db77fb27 - 2425 - 4d42-81c6-e7e3228d9f46, OU= Tenants, OU= MSOnline, DC= spods18016244, DC= msoprd, DC= msft, DC= netSPS - SourceObjectDN -
//SPS - ClaimProviderType - Forms
//SPS-SavedAccountName - i:0#.f|membership|bora@itadev.onmicrosoft.com
//SPS-SavedSID - System.Byte[]
//SPS-ObjectExists - 
//SPS-PersonalSiteCapabilities - 0
//SPS-PersonalSiteFirstCreationTime - 
//SPS-PersonalSiteLastCreationTime - 
//SPS-PersonalSiteNumberOfRetries - 
//SPS-PersonalSiteFirstCreationError - 
//SPS-FeedIdentifier - 
//WorkEmail - bora @itadev.onmicrosoft.com
//CellPhone - 111-222-3344Fax - 
//HomePhone - 
//Office - Office of Chief Information Technology
//SPS-Location - 
//Assistant - 
//SPS-PastProjects - 
//SPS-Skills - 
//SPS-School - 
//SPS-Birthday - 
//SPS-StatusNotes - 
//SPS-Interests - 
//SPS-HashTags - 
//SPS-EmailOptin - 0
//SPS-PrivacyPeople - True
//SPS-PrivacyActivity - 4095SPS-PictureTimestamp - 63656652524
//SPS-PicturePlaceholderState - 1
//SPS-PictureExchangeSyncState - 1
//SPS-TimeZone - 
//OfficeGraphEnabled - 
//SPS-UserType - 0
//SPS-HideFromAddressLists - False
//SPS-RecipientTypeDetails - 
//DelveFlags - 
//msOnline-ObjectId - 5a90cea2-f71c-4c35-a42c-e32ef3c0e0d8SPS-PointPublishingUrl - 
//SPS-TenantInstanceId - 
//SPS-SharePointHomeExperienceState - 
//SPS-MultiGeoFlags - 
//PreferredDataLocation - 
//EmployeeID - 1005
//EmployeeType - Contractor
//AccountExpiration - 12/31/2019 12:00:00 AM
//CountrySpecialities - France|Italy|Spain
//IndustrySpecialities - Automobiles|Chemicals|Clean Coal Technology
//Education
//Certifications
//EmergencyContacts
    /// 
    /// 
    /// 
    /// 
    /// 
    /// 
    /// </summary>
    class UserProfile
    {
        public string firstName { get; set; }
        public string middleName { get; set; }
        public string lastName { get; set; }
        public string suffix { get; set; }
        public string workPhone { get; set; }
        public string cellPhone { get; set; }
        public string jobTitle { get; set; }
        public string department { get; set; }
        public string emailAddress { get; set; }
        public string officeRegion { get; set; }
        public string officeCountry { get; set; }
        public string officeState { get; set; }
        public MailingAddress mailingAddress { get; set; }        
        public string officeNumber { get; set; }
        public User manager { get; set; }
        public string employeeID { get; set; }
        public string employeeType { get; set; }
        public string accountExpiration { get; set; }
        public string[] countrySpecialities { get; set; }
        public string[] industrySpecialities { get; set; }
        public DateTime? CSATCompletion { get; set; }
        public List<Contact> emergencyContactInformation { get; set; }
        public List<Education> education { get; set; }
        public List<Certification> certifications { get; set; }
        public List<User> directReports { get; set; }
    }
    class MailingAddress
    {
        public string addressLine1 { get; set; }
        public string addressLine2 { get; set; }
        public string state { get; set; }
        public string city { get; set; }
        public string zipCode { get; set; }
        public string country { get; set; }
    }
    class User
    {
        public string displayName { get; set; }
        public string email { get; set; }
    }
    class Contact
    {
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string phoneNumber { get; set; }
        public string emailAddress { get; set; }
    }
    class Education
    {
        public string schoolName { get; set; }
        public string degree { get; set; }
        public string year { get; set; }
    }
    class Certification
    {
        public string title { get; set; }
        public string organization { get; set; }
        public string year { get; set; }
    }
}
