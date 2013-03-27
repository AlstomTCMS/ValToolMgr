using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsPopup : CTsGenericInstr
    {


        public CTsPopup(string text)
        {
            this.Text = TestStandAdapter.protectText(text);
        }

        public override string InstructionName
        {
            get { return "MessagePopup"; }
            protected set { }
        }
        

    }
}
