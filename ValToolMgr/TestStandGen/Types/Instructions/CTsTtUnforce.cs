using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTtUnforce : CTsTtVariable
    {
        public override string InstructionName
        {
            get { return "Variable_Release"; }
            protected set { }
        }

        public CTsTtUnforce(CVariable var)
            : base(var)
        {
        }
    }
}
