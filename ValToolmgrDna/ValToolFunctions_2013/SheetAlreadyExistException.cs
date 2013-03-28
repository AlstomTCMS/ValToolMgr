using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolFunctions_2013
{
    public class SheetAlreadyExistException : Exception
    {
        public SheetAlreadyExistException() { }
        public SheetAlreadyExistException(string message) : base(message) { }
        public SheetAlreadyExistException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
