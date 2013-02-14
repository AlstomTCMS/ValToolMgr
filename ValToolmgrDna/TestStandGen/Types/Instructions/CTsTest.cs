using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrDna.Interface;

namespace TestStandGen.Types.Instructions
{
    class CTsTest : CTsGenericInstr
    {
        public string Name;
        public string Value;
        public string Path;

        public override string InstructionName
        {
            get { return "CB_Test"; }
            protected set { }
        }

        public CTsTest(CVariable var)
        {
            Name = var.name;
            Value = var.value.ToString();
            Path = var.path;
            this.Text = "Force " + Name + " at " + Value;
        }
    }
}
