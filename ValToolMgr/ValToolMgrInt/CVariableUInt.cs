using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableUInt : CVariable
    {
        private UInt32 Value = 0;

        public CVariableUInt(string VariableName, string Location, string Path, string Value)
        {
            this.name = VariableName;
            this.path = Path;
            this.value = Value;
            this.Location = Location;
        }

        public override object value   // the Name property
        {
            get
            {
                return Value;
            }

            set
            {
                if (value == null)
                {
                    Value = 0;
                }
                else
                {
                    Value = Convert.ToUInt32(value);
                }
            }
        }
    }
}
