using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableInt : CVariable
    {

        public CVariableInt(string VariableName, string Location, string Path)
        {
            this.name = VariableName;
            this.path = Path;
            this.Location = Location;
        }

        public override string convValToValidStr(string value)
        {
            int Value = 0;
            if (value != null) Value = Convert.ToInt32(value);
            return Value.ToString();
        }
    }
}
