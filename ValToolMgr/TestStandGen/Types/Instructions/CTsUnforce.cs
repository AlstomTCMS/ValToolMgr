using System;

namespace TestStandGen.Types.Instructions
{
    class CTsUnforce : CTsBasedVarInstr
    {
        public override string InstructionName
        {
            get { return "CB_UnForce"; }
            protected set { }
        }

        public CTsUnforce(CTsVariable var)
            : base(var)
        {
        }
    }
}
