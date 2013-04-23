using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ValToolFunctionsStub
{
    public class StringValueAttribute : System.Attribute
    {

        private string _value;

        public StringValueAttribute(string value)
        {
            _value = value;
        }

        public string Value
        {
            get { return _value; }
        }

        public override string ToString()
        {
            return _value;
        }
    }

    public class StringEnum
    {
        public static string GetStringValue(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            StringValueAttribute[] attributes =
                (StringValueAttribute[])fi.GetCustomAttributes(typeof(StringValueAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Value;
            else
                return value.ToString();
        }
    }
}
