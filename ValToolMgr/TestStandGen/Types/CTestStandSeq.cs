using System;
using System.Collections;
using System.Linq;
using System.Text;
using TestStandGen.Types.Instructions;

namespace TestStandGen.Types
{
    class CTestStandSeq : TestStandArray
    {
        //public List<CTsGenericInstr> List = new List<CTsGenericInstr>();

        public string identifier;

        public string Title { get; set; }
    }
}
