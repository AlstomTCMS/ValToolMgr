using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.Interface
{
    public class CVariableBool : CVariable
    {
        private bool Value;

        public override object value   // the property
        {
            get
            {
                return Value;
            }

            set
            {
                Value = Convert.ToBoolean(value);
            }
        }
    }
}
