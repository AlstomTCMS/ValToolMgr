using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen
{
    class TestStandFile
    {
        private string filename;
        public string Filename
        {
            get
            {
                return filename;
            }

            set
            {
                filename = TestStandAdapter.protectBackslashes(value);
            }
        }

        public CTestStandSeqContainer Sequences { get; set; }

        public TestStandFile(string Filename)
        {
            this.Filename = Filename;
            Sequences = new CTestStandSeqContainer();
        }
    }
}
