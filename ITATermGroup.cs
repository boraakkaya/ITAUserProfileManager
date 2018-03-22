using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITAUserProfileManager
{
    public class ITATermStore
    {
        public List<ITATermGroup> ITATermGroupList { get; set; }
    }
    public class ITATermGroup
    {
        public Guid Id { get; set; }
        public String Name { get; set; }
        public List<ITATermSets> TermSets { get; set; }
    }
    public class ITATermSets
    {
        public Guid Id { get; set; }
        public String Name { get; set; }
        public String Type { get; set; }
        public List<ITATermProperty> CustomProps { get; set; }
        public List<ITATerms> Terms { get; set; }
    }
    public class ITATerms
    {
        public Guid Id { get; set; }
        public String Name { get; set; }
        public String Type { get; set; }
        public List<ITATerms> Terms { get; set; }
        public List<string> Labels { get; set; }
        public List<ITATermProperty> CustomProps { get; set; }
    }
    public class ITATermProperty
    {
        public string key { get; set; }
        public string value { get; set; }
    }
    
}
