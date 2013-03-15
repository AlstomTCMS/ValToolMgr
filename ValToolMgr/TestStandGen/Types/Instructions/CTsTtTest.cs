using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTtTest : CTsTtVariable
    {
        public override string InstructionName
        {
            get { return "Variable_Read_Test"; }
            protected set { }
        }

        public CTsTtTest(CVariable var)
            : base(var)
        {
        }
    }
}
