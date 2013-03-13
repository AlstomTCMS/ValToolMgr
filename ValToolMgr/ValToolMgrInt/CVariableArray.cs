using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableArray : CVariable
    {
        private CVariable Variable;
        public uint Index;

        public CVariableArray(CVariable Var, uint Index)
        {
            // TODO: Complete member initialization
            this.Variable = Var;
            this.path = Var.path;
            this.Location = Var.Location;
            this.Index = Index;
        }

        public override object value   // the property
        {
            get
            {
                return Variable;
            }

            set
            {
                if (value == null)
                {
                    Variable = null;
                }
                else
                {
                    Variable = (CVariable)value;
                }
            }
        }
    }
}
