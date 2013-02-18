using System;
using System.Collections;
using System.Linq;
using System.Text;

namespace TestStandGen.Types
{
    class TestStandArray : ArrayList
    {
        public int Max
        {
            get
            {
                return this.Count - 1;
            }
            protected set
            {
            }
        }
    }
}
