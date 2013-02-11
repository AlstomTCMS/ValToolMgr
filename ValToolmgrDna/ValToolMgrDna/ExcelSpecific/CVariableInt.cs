using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.ExcelSpecific
{
    class CVariableInt : CVariable
    {
        private int Value = 0;

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
