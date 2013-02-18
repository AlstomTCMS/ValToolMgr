using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsTestAna : CTsGenericInstr
    {
        public string Name;
        public string Value;
        public string Path;

        public override string InstructionName
        {
            get { return "CB_TestAna"; }
            protected set { }
        }

        public CTsTestAna(CVariable var)
        {
            Name = var.name;
            Value = var.value.ToString().Replace(',', '.');
            Path = var.path;
        }
    }
}
