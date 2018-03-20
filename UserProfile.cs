using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITAUserProfileManager
{   
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
