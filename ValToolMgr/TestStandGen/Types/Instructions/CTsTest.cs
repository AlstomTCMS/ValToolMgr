using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTest : CTsCbVariable
    {
        public override string InstructionName
        {
            get { return "CB_Test"; }
            protected set { }
        }

        public CTsTest(CVariable var)
            : base(var)
        {
        }
    }
}
