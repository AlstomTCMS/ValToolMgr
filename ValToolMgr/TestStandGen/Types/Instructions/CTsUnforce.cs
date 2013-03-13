using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsUnforce : CTsCbVariable
    {
        public override string InstructionName
        {
            get { return "CB_UnForce"; }
            protected set { }
        }

        public CTsUnforce(CVariable var)
            : base(var)
        {
        }
    }
}
