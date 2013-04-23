using System;
using System.Collections;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CStep
    {
        public string DescCheck;
        public string title;
        public string DescAction;
        public ArrayList actions = new ArrayList();
        public ArrayList checks = new ArrayList();

        public CStep(string title, string descAction, string descCheck)
        {
            this.title = title;
            this.DescAction = descAction;
            this.DescCheck = descCheck;
        }
    }
}
