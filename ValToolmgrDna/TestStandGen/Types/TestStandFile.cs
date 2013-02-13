﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using TestStandGen.Types.Instructions;

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

        public List<CTsSequenceCall> HeaderList = new List<CTsSequenceCall>();

        public TestStandFile(string Filename)
        {
            this.Filename = Filename;
            Sequences = new CTestStandSeqContainer();
        }

        public void addSequence(string text, CTestStandSeq sequence)
        {
            this.HeaderList.Add(new CTsSequenceCall(sequence.identifier, text));

            this.Sequences.Add(sequence);
        }
    }
}
