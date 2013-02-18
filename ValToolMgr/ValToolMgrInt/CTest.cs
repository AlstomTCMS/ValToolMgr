using System;
using System.Collections;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CTest : ArrayList
    {
        public string Title;
        public string Description;

        public CTest(string title, string description)
        {
            // TODO: Complete member initialization
            this.Title = title;
            this.Description = description;
        }
    }
}
