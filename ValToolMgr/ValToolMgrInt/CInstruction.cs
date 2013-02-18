using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public abstract class CInstruction
    {
        public bool Skipped = false;
        public bool ForceFailed = false;
        public bool ForcePassed = false;

        public object data { get; set; }
    }
}
