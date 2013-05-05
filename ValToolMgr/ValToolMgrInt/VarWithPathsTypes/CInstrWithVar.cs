using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrInt
{
    public abstract class CInstrWithVar : CInstruction
    {
        public CVariable Variable { get; set; }
    }
}
