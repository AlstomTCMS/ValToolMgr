using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.Interface
{
    class CVariableDouble : CVariable
    {
        private double Value = 0.0;

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
                    Value = Convert.ToDouble(value);
                }
            }
        }
    }
}
