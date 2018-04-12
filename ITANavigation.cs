using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ITAUserProfileManager
{
    public class ITANavigation
    {
        public List<ITANavigationItem> navigation { get; set; }
        //public List<ITAStaticLink> staticLinks { get; set; }
    }
    public class ITANavigationItem
    {
        public string title { get; set; }
        public string link { get; set; }
        //public string description { get; set; }
        //public bool visible { get; set; }
        //public string target { get; set; }
        public List<ITANavigationItem> children { get; set; }
    }
    //public class ITAStaticLink
    //{
    //    public string owner { get; set; }
    //    public List<ITAStaticLinkHeading> headings { get; set; }
        
    //}
    //public class ITAStaticLinkHeading
    //{
    //    public string title { get; set; }
    //    public List<ITANavigationItem> links { get; set; }
    //}
}
