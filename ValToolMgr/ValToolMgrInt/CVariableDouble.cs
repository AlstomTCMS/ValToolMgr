using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableDouble : CVariable
    {
        private double Value = 0.0;

        public CVariableDouble(string VariableName, string Location, string Path, string Value)
        {
            this.name = VariableName;
            this.path = Path;
            this.value = Value;
            this.Location = Location;
        }

        public override object value   // the property
        {
            get
            {
                return Value;
            }

            set
            {
                if (value == null)
                {
                    Value = 0.0;
                }
                else
                {
                    if (value is string) value = ((string)value).Replace('.', ',');
                    Value = Convert.ToDouble(value);
                }
            }
        }
    }
}
