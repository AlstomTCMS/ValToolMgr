using System;

namespace TestStandGen.Types.Instructions
{
    class CTsTestAna : CTsBasedVarInstr
    {

        public override string InstructionName
        {
            get { return "CB_TestAna"; }
            protected set { }
        }

        public CTsTestAna(CTsVariable var) : base(var)
        {
        }
    }
}
