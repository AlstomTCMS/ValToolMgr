using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableInt : CVariable
    {
        private int Value = 0;

        public CVariableInt(string VariableName, string Path, string Value)
        {
            this.name = VariableName;
            this.path = Path;
            this.value = Value;
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
                    Value = Convert.ToInt32(value);
                }
            }
        }
    }
}
