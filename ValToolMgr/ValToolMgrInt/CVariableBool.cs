using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableBool : CVariable
    {
        private bool Value;

        public CVariableBool(string VariableName, string Location, string Path, string Value)
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
                if (value == null) value = false;
                if (value.ToString() == "0") value = false;
                else if(value.ToString() == "1") value = true;
                Value = Convert.ToBoolean(value);
            }
        }
    }
}
