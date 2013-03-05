using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsForce : CTsGenericInstr
    {
        public string Name;
        public string Value;
        public string Path;

        public override string InstructionName
        {
            get { return "CB_Force"; }
            protected set { }
        }

        public CTsForce(CVariable var)
        {
            Name = var.name;
            if (String.Equals(var.GetType().FullName, typeof(CVariableDouble).FullName))
                Value = var.value.ToString().Replace(',', '.');
            else
                Value = var.value.ToString();
            Path = var.path;
        }
    }
}
