using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolFunctions_2013
{
    public class ExcelApplicationMissingException : Exception
    {
        public ExcelApplicationMissingException() { }
        public ExcelApplicationMissingException(string message) : base(message) { }
        public ExcelApplicationMissingException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
