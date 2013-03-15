using System;
using System.Xml;
using System.Collections;
using ValToolMgrInt;

namespace TestStandGen.Types.Instructions
{
    abstract class CTsTtVariable : CTsGenericInstr
    {
        public string Name;
        public string Value;
        public string Path;
        public string Location;


        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);



        public CTsTtVariable(CVariable var)
        {
            if (String.Equals(var.GetType().FullName, typeof(CVariableArray).FullName))
            {
                CVariableArray v = (CVariableArray)var;
                parseVariable((CVariable)v.value);
            }
            else
            {
                parseVariable(var);
            }


            logger.Debug("OK");
        }

        private void parseVariable(CVariable var)
        {

            if (String.Equals(var.GetType().FullName, typeof(CVariableDouble).FullName))
                Value = var.value.ToString().Replace(',', '.');
            else
                Value = var.value.ToString();

            Name = var.name;
            Path = var.path;
            Location = var.Location;
        }
    }
}
