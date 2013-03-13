using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsForceArray : CTsCbVariable
    {
        public uint Index;

        public override string InstructionName
        {
            get { return "CB_ForceArrayElement"; }
            protected set { }
        }

        public CTsForceArray(CVariableArray var) : base(var)
        {
            Index = var.Index;
        }
    }
}
