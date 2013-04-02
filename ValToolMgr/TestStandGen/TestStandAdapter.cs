using System;
using System.Text.RegularExpressions;

namespace TestStandGen
{
    public class TestStandAdapter
    {
        public static string protectText(string value)
        {
            value = value.Replace("\\\\", "\\");
            value = value.Replace("\\", "\\\\");
            string output = Regex.Replace(value, "(\r\n|\n)", "\\n");
            value = value.Replace("\"", "\\\\\\\"");
            return value;
        }
    }
}
