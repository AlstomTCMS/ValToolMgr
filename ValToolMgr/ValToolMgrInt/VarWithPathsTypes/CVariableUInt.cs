using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableUInt : CVariable
    {
        public CVariableUInt(string VariableName, string Location, string Path)
        {
            this.name = VariableName;
            this.path = Path;
            this.Location = Location;
        }

        public override string convValToValidStr(string value)
        {
            uint Value = 0;
            if (value != null) Value = Convert.ToUInt32(value);
            return Value.ToString();
        }
    }
}
