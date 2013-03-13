using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTestAna : CTsCbVariable
    {

        public override string InstructionName
        {
            get { return "CB_TestAna"; }
            protected set { }
        }

        public CTsTestAna(CVariableDouble var) : base(var)
        {
        }
    }
}
