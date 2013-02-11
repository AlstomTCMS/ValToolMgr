using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.ExcelSpecific
{
    class CVariableDateTime : CVariable
    {
        private DateTime Value;

        public override object value   // the property
        {
            get
            {
                return Value;
            }

            set
            {
                Value = Convert.ToDateTime(value);
            }
        }
    }
}
