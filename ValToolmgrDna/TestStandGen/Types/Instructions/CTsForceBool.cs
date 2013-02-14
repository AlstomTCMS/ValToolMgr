using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrDna.Interface;

namespace TestStandGen.Types.Instructions
{
    class CTsForceBool : CTsForce
    {

        public CTsForceBool(CVariable var)
            : base(var)
        {
            this.Text = "Force " + Name + " at " + Value + "(Bool)";
        }
    }
}
