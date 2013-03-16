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
            if (String.Equals(variable.GetType().FullName, typeof(CVariableBool).FullName))
                Value = Convert.ToInt32(Convert.ToBoolean(variable.value)).ToString();
            Location = ttTarget.Identifier;
            Path = ttTarget.prefix + Path;
        }
    }
}
