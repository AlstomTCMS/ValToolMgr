using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolMgrInt;

namespace ValToolMgrDna.ExcelSpecific
{
    public class VariableParser
    {
        public enum SupportedTypes
        {
            REAL,
            INTEGER,
            UNSIGNED_INTEGER,
            BOOLEAN
        }

        public static CVariable parseAsVariable(string VariableName, string Path, string Value, SupportedTypes expectedType)
        {
            switch (expectedType)
            {
                case SupportedTypes.BOOLEAN:
                    return new CVariableBool(VariableName, Path, Value);
                case SupportedTypes.INTEGER:
                    return new CVariableInt(VariableName, Path, Value);
                case SupportedTypes.REAL:
                    return new CVariableDouble(VariableName, Path, Value);
                case SupportedTypes.UNSIGNED_INTEGER:
                    return new CVariableUInt(VariableName, Path, Value);
                default:
                    return null;
            }
        }

        public static CVariable parseAsVariable(string VariableName, string Path, string Value)
        {

            string[] explodedValue = VariableName.Split(':');

            if(explodedValue.Length > 2)
                throw new FormatException(String.Format("\"{0}\" contains too much \":\"", VariableName));

            if (explodedValue.Length == 1) 
            {
                return parseAsVariable(VariableName, Path, Value, SupportedTypes.BOOLEAN);
            }
            else
            {
                switch(explodedValue[0])
                {
                    case "I":
                        return parseAsVariable(explodedValue[1], Path, Value, SupportedTypes.INTEGER);
                    case "R":
                        return parseAsVariable(explodedValue[1], Path, Value, SupportedTypes.REAL);
                    default:
                        throw new FormatException(String.Format("\"{0}\" is not valid type specifier.", explodedValue[0]));
                }
            }
        }

        //private CVariable buildVariable(string Target, string Location, object CellValue)
        //{
        //    CVariable buildVariable;
        //    string typeExpected = "UNDEF";

        //    try
        //    {
        //        if (Target.IndexOf("I:") == 0)
        //        {
        //            typeExpected = "INT";
        //            buildVariable = new CVariableInt();
        //            Target = Target.Substring(2);
        //            buildVariable.value = Convert.ToInt32(CellValue);

        //        }
        //        else if (Target.IndexOf("R:") == 0)
        //        {
        //            typeExpected = "REAL";
        //            buildVariable = new CVariableDouble();
        //            Target = Target.Substring(2);
        //            buildVariable.value = Convert.ToDouble(CellValue);
        //        }
        //        else if (Target.IndexOf("DT:") == 0)
        //        {
        //            throw new NotImplementedException();
        //        }
        //        else
        //        {
        //            typeExpected = "BOOL";
        //            buildVariable = new CVariableBool();
        //            buildVariable.value = Convert.ToBoolean(Convert.ToInt32(CellValue));
        //        }

        //        buildVariable.name = Target;
        //        buildVariable.path = Location;
        //    }
        //    catch (Exception)
        //    {
        //        throw new InvalidCastException("Invalid value, expected : " + typeExpected + ", have : " + CellValue.ToString());
        //    }

        //    return buildVariable;
        //}

    }
}
