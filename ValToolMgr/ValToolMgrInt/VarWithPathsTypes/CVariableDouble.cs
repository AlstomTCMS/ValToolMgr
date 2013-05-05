using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableDouble : CVariable
    {
        public CVariableDouble(string VariableName, string Location, string Path)
        {
            this.name = VariableName;
            this.path = Path;
            this.Location = Location;
        }

        public override string convValToValidStr(string value)
        {
            double Value = 0.0;
            if (value != null)
            {
                if (value is string) value = ((string)value).Replace('.', ',');
                Value = Convert.ToDouble(value);
            }
            return Value.ToString();
        }
    }
}
