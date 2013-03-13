using System;

namespace TestStandGen
{
    class TestStandAdapter
    {
        public static string protectBackslashes(string value)
        {
            return value.Replace("\\", "\\\\");
        }
    }
}
