using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTestArray : CTsBasedVarInstr
    {
        public uint Index;

        public override string InstructionName
        {
            get { return "CB_TestArrayElement"; }
            protected set { }
        }

        public CTsTestArray(CTsVariable var) : base(var)
        {
            Index = var.Index;
        }
    }
}
