using System;
using System.Xml;
using System.Collections;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTtVariable : CTsVariable
    {
        public CTsTtVariable(CVariable var)
            : base(var)
        {
        }

        public CTsTtVariable(CTsInstrFactory.TtTarget ttTarget, CVariable variable)
            : base(variable)
        {
            Location = ttTarget.Identifier;
            Path = ttTarget.prefix + Path;
        }
    }
}
