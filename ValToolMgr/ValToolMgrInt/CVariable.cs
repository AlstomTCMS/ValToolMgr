using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ValToolMgrInt
{
    public abstract class CVariable
    {
        private string Name;
        private string location;


        public string name
        {
            get
            {
                return Name;
            }

            set
            {
                if (Regex.IsMatch(value, "^[a-zA-Z]+[a-zA-Z0-9_]*$"))
                {
                    Name = value;
                }
                else
                {
                    throw new FormatException(String.Format("\"{0}\" is invalid for variable name.", value));
                 }
            }
        }

        public string path { get; set; }

        public string Location
        {
            get
            {
                return location;
            }

            set
            {
                if (Regex.IsMatch(value, "^[a-zA-Z]+[a-zA-Z0-9_]*/[a-zA-Z]+[a-zA-Z0-9_]*$"))
                {
                    location = value;
                }
                else
                {
                    throw new FormatException(String.Format("\"{0}\" is invalid for Location name.", value));
                }
            }
        }

        public abstract object value { get; set; }
    }
}
