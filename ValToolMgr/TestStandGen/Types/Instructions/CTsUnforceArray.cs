using System;

namespace TestStandGen.Types.Instructions
{
    class CTsUnforceArray : CTsBasedVarInstr
    {
        public uint Index;

        public override string InstructionName
        {
            get { return "CB_UnForceArrayElement"; }
            protected set { }
        }

        public CTsUnforceArray(CTsVariable var)
            : base(var)
        {
            Index = var.Index;
        }
    }
}
