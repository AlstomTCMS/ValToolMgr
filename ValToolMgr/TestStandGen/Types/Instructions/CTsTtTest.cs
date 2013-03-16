using System;

namespace TestStandGen.Types.Instructions
{
    class CTsTtTest : CTsBasedVarInstr
    {
        public override string InstructionName
        {
            get { return "Variable_Read_Test"; }
            protected set { }
        }

        public CTsTtTest(CTsVariable var)
            : base(var)
        {
        }
    }
}
