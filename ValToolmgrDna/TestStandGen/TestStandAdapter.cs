using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
