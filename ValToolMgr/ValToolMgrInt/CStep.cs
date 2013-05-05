using System.Collections.Generic;

namespace ValToolMgrInt
{
    public class CStep
    {
        public string DescCheck;
        public string title;
        public string DescAction;
        public List<CInstruction> actions = new List<CInstruction>();
        public List<CInstruction> checks = new List<CInstruction>();

        public CStep(string title, string descAction, string descCheck)
        {
            this.title = title;
            this.DescAction = descAction;
            this.DescCheck = descCheck;
        }
    }
}
