using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTestArray : CTsGenericInstr
    {
        public string Name;
        public string Value;
        public string Path;
        public uint Index;

        public override string InstructionName
        {
            get { return "CB_TestArrayElement"; }
            protected set { }
        }

        public CTsTestArray(CVariableArray var)
        {
            CVariable variable = (CVariable)var.value;
            Name = variable.name;
            Value = variable.value.ToString();
            Path = variable.path;
            Index = var.Index;
        }
    }
}
