using System;

namespace TestStandGen.Types.Instructions
{
    class CTsTtUnforce : CTsBasedVarInstr
    {
        public override string InstructionName
        {
            get { return "Variable_Release"; }
            protected set { }
        }

        public CTsTtUnforce(CTsVariable var)
            : base(var)
        {
        }
    }
}
