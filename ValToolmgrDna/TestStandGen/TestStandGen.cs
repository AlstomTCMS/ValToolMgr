using System;
using TestStandGen.Types;
using TestStandGen.Types.Instructions;

using ValToolMgrInt;

namespace TestStandGen
{
    using Antlr.Runtime;
    using Antlr4.StringTemplate;
    using Antlr4.StringTemplate.Compiler;
    using Antlr4.StringTemplate.Misc;
    using System.IO;

    public class TestStandGen
    {
        private CTestContainer sequence;
        private string outFile;
        private string templatePath;


        private bool alreadyGenerated;
        private TestStandFile TSFile;

        public static void genSequence(CTestContainer sequence, string outFile, string templatePath)
        {
            TestStandGen test = new TestStandGen(sequence, outFile, templatePath);
            
            test.writeScenario();
        }

        private TestStandGen(CTestContainer sequence, string outFile, string templatePath)
        {
            this.sequence = sequence;
            this.outFile = outFile;
            this.templatePath = templatePath;
            initialize();
        }

        private void initialize()
        {
            TSFile = new TestStandFile(outFile);
            alreadyGenerated = false;
        }

        private void writeScenario()
        {
            if (alreadyGenerated) initialize();

            TemplateGroup group = new TemplateGroupDirectory(this.templatePath, '$', '$');

                ErrorBuffer errors = new ErrorBuffer();
                group.Listener = errors;
                group.Load();

                Template st = group.GetInstanceOf("MainTemplate");

                this.TSFile = genTsStructFromTestContainer(this.outFile, sequence);

                st.Add("TestStandFile", this.TSFile);

                string result = st.Render();

                if (errors.Errors.Count > 0)
                {
                    foreach (TemplateMessage m in errors.Errors)
                    {
                        throw new Exception(m.ToString());
                    }
                }

                StreamWriter output = new StreamWriter(this.outFile);

                output.Write(result);
                output.Close();

                CTsGenericInstr.resetIdCounter();
            this.alreadyGenerated = true;
        }


        /// <summary>
        /// Converts a potentially complex tree structure to a standardized, linear TestStand sequence list
        /// </summary>
        /// <param name="sequence">Sequence to convert</param
        /// <returns>Sequence, in a format understandable to generate</returns>
        private TestStandFile genTsStructFromTestContainer(string filename, CTestContainer sequence)
        {
            TestStandFile ts = new TestStandFile(filename);

            foreach(CTest test in sequence)
            {
                CTestStandSeq SubSeq = genInstrListFromTest(test);

                ts.addSequence("Call to subsequence " + SubSeq.identifier, SubSeq);

            }
            return ts;
        }

        private CTestStandSeq genInstrListFromTest(CTest TestContainer)
        {
            CTestStandSeq SubSeq = new CTestStandSeq();
            SubSeq.identifier = TestContainer.Title;
            SubSeq.Title = TestContainer.Title;

            foreach (CStep step in TestContainer)
            {
                SubSeq.Add(new CTsLabel("===================================="));
                SubSeq.Add(new CTsLabel("========  " + step.title));
                SubSeq.Add(new CTsLabel("===================================="));

                if (step.DescAction.Length > 0) SubSeq.Add(new CTsLabel("= Actions : " + step.DescAction));
                foreach(CInstruction instr in step.actions)
                {
                     SubSeq.Add(getTsEquivFromInstr(instr));
                }

                if (step.DescCheck.Length > 0) SubSeq.Add(new CTsLabel("= Checks : " + step.DescCheck));
                foreach (CInstruction instr in step.checks)
                {
                    SubSeq.Add(getTsEquivFromInstr(instr));
                }
            }
            return SubSeq;
        }

        private CTsGenericInstr getTsEquivFromInstr(CInstruction inst)
        {
            string typeOfStep = inst.GetType().ToString();
            string typeOfData = inst.data.GetType().ToString();

            CTsGenericInstr instr = null;

            if (String.Equals(typeOfStep, typeof(CInstrPopup).FullName))
                instr = new CTsPopup(inst.data.ToString());

            if (String.Equals(typeOfStep, typeof(CInstrWait).FullName))
                instr = new CTsWait(Convert.ToInt32(inst.data));

            if (String.Equals(typeOfStep, typeof(CInstrUnforce).FullName) && !String.Equals(typeOfData, typeof(CVariableArray).FullName))
                instr = new CTsUnforce((CVariable)inst.data);

            if (String.Equals(typeOfStep, typeof(CInstrForce).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName) || String.Equals(typeOfData, typeof(CVariableUInt).FullName) || String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    instr = new CTsForce((CVariable)inst.data);

                if (String.Equals(typeOfData, typeof(CVariableArray).FullName))
                    instr = new CTsForceArray((CVariableArray)inst.data);
            }

            if (String.Equals(typeOfStep, typeof(CInstrTest).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName)  || String.Equals(typeOfData, typeof(CVariableUInt).FullName))
                    instr = new CTsTest((CVariable)inst.data);

                if(String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    instr = new CTsTestAna((CVariable)inst.data);

                if (String.Equals(typeOfData, typeof(CVariableArray).FullName))
                    instr = new CTsTestArray((CVariableArray)inst.data);
            }

            if(instr != null)
            {
                instr.Skipped = inst.Skipped;
                instr.ForceFailed = inst.ForceFailed;
                instr.ForcePassed = inst.ForcePassed;
                return instr;
            }

            throw new NotImplementedException(String.Format("Data not handled : [{0}, {1}]", typeOfStep, typeOfData));
        }
    }
}
