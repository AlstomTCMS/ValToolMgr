using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public class CVariableBool : CVariable
    {

        public CVariableBool(string VariableName, string Location, string Path)
        {
            this.name = VariableName;
            this.path = Path;
            this.Location = Location;
        }



        public override string convValToValidStr(string value)
        {
            if (value == null) value = "False";
            if (value.ToString() == "0") value = "False";
            else if(value.ToString() == "1") value = "True";
            return Convert.ToBoolean(value).ToString();
        }
    }
}
