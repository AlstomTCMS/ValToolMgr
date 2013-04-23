using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolFunctionsStub
{
    public class ExcelApplicationNotAvailableException : Exception
    {
        public ExcelApplicationNotAvailableException() { }
        public ExcelApplicationNotAvailableException(string message) : base(message) { }
        public ExcelApplicationNotAvailableException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
