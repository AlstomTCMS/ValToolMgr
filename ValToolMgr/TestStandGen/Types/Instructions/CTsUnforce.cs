using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    class CTsUnforce : CTsGenericInstr
    {
        public string Name;
        public string Path;

        public override string InstructionName
        {
            get { return "CB_UnForce"; }
            protected set { }
        }

        public CTsUnforce(CVariable var)
        {
            Name = var.name;
            Path = var.path;
            this.Text = "Unforce " + Name;
        }
    }
}
