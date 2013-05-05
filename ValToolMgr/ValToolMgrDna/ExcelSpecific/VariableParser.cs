using System;
using System.Text.RegularExpressions;
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

        public static CVariable parseAsVariable(string VariableName, string Location, string Path, SupportedTypes expectedType)
        {

	        // Here we call Regex.Match.
	        Match match = Regex.Match(VariableName, @"^(.*)\[(\d)*\]$");

	        // Here we check the Match instance.
	        if (match.Success)
	        {
	            // Finally, we get the Group value and display it.
	            VariableName = match.Groups[1].Value;
                uint Index = Convert.ToUInt32(match.Groups[2].Value);

                CVariable Var = parseAsVariable(VariableName, Location, Path, expectedType);
                CVariableArray array = new CVariableArray(Var, Index);
                return array;
	        }

            switch (expectedType)
            {
                case SupportedTypes.BOOLEAN:
                    return new CVariableBool(VariableName, Location, Path);
                case SupportedTypes.INTEGER:
                    return new CVariableInt(VariableName, Location, Path);
                case SupportedTypes.REAL:
                    return new CVariableDouble(VariableName, Location, Path);
                case SupportedTypes.UNSIGNED_INTEGER:
                    return new CVariableUInt(VariableName, Location, Path);
                default:
                    return null;
            }
        }

        public static CVariable parseAsVariable(string VariableName, string Location, string Path)
        {
            string[] explodedValue = VariableName.Split(':');

            if(explodedValue.Length > 2)
                throw new FormatException(String.Format("\"{0}\" contains too much \":\"", VariableName));

            if (explodedValue.Length == 1) 
            {
                return parseAsVariable(VariableName, Location, Path, SupportedTypes.BOOLEAN);
            }
            else
            {
                switch(explodedValue[0])
                {
                    case "I":
                        return parseAsVariable(explodedValue[1], Location, Path, SupportedTypes.INTEGER);
                    case "R":
                        return parseAsVariable(explodedValue[1], Location, Path, SupportedTypes.REAL);
                    default:
                        throw new FormatException(String.Format("\"{0}\" is not valid type specifier.", explodedValue[0]));
                }
            }
        }
    }
}
