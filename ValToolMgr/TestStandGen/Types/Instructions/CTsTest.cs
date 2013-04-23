using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen.Types.Instructions
{
    class CTsTest : CTsBasedVarInstr
    {
        public override string InstructionName
        {
            get { return "CB_Test"; }
            protected set { }
        }

        public CTsTest(CTsVariable var)
            : base(var)
        {
        }
    }
}
