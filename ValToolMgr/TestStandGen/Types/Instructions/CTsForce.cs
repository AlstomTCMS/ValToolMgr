using System;

namespace TestStandGen.Types.Instructions
{
    class CTsForce : CTsBasedVarInstr
    {
        public override string InstructionName
        {
            get { return "CB_Force"; }
            protected set { }
        }

        public CTsForce(CTsVariable var)
            : base(var)
        {
        }
    }
}
