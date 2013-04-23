using System;
using System.Collections;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.Interface
{
    public class CStep
    {
        public string DescCheck;
        public string title;
        public string DescAction;
        public ArrayList actions = new ArrayList();
        public ArrayList checks = new ArrayList();
    }
}
