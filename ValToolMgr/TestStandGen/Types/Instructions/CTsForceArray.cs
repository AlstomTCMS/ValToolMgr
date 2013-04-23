using System;

namespace TestStandGen.Types.Instructions
{
    class CTsForceArray : CTsBasedVarInstr
    {
        public uint Index;

        public override string InstructionName
        {
            get { return "CB_ForceArrayElement"; }
            protected set { }
        }

        public CTsForceArray(CTsVariable var) : base(var)
        {
            Index = var.Index;
        }
    }
}
