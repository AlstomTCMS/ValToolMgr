using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen.Types.Instructions
{
    abstract class CTsBasedVarInstr : CTsGenericInstr
    {
        public CTsVariable Variable;

        public CTsBasedVarInstr(CTsVariable var)
        {
            this.Variable = var;
        }
    }
}
