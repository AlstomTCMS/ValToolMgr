using System;
using System.Xml;
using System.Collections;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsCbVariable : CTsVariable
    {

        public CTsCbVariable(CVariable var) : base(var)
        {
        }

        public CTsCbVariable(CTsInstrFactory.CbTarget cbTarget, CVariable variable) : base(variable)
        {
            Location = cbTarget.Identifier;
        }
    }
}
