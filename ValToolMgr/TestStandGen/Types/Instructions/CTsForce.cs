using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsForce : CTsCbVariable
    {
        public override string InstructionName
        {
            get { return "CB_Force"; }
            protected set { }
        }

        public CTsForce(CVariable var)
            : base(var)
        {
        }
    }
}
