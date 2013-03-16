using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTtForce : CTsBasedVarInstr
    {
        public override string InstructionName
        {
            get { return "Variable_Force"; }
            protected set { }
        }

        public CTsTtForce(CTsVariable var)
            : base(var)
        {
        }
    }
}
