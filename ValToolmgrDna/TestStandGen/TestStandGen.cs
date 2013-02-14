using System;
using TestStandGen.Types;
using TestStandGen.Types.Instructions;

using ValToolMgrDna.Interface;

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
            try
            {
                ErrorBuffer errors = new ErrorBuffer();
                group.Listener = errors;
                group.Load();

                Template st = group.GetInstanceOf("MainTemplate");
                Console.WriteLine("==========================");

                this.TSFile = genTsStructFromTestContainer(this.outFile, sequence);

                st.Add("TestStandFile", this.TSFile);

                string result = st.Render();

                if (errors.Errors.Count > 0)
                {
                    foreach (TemplateMessage m in errors.Errors)
                    {
                        Console.WriteLine(m.ToString());
                    }
                    Console.ReadLine();
                }

                StreamWriter output = new StreamWriter(this.outFile);

                output.Write(result);
                output.Close();
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                Console.ReadLine();
            }

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
            SubSeq.identifier = TestContainer.title;
            SubSeq.Title = TestContainer.title;

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

            if (String.Equals(typeOfStep, typeof(CInstrPopup).FullName))
                return new CTsPopup(inst.data.ToString());

            if (String.Equals(typeOfStep, typeof(CInstrWait).FullName))
                return new CTsWait(Convert.ToInt32(inst.data));

            if (String.Equals(typeOfStep, typeof(CInstrUnforce).FullName))
                return new CTsUnforce((CVariable)inst.data);

            if (String.Equals(typeOfStep, typeof(CInstrForce).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName) || String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    return new CTsForce((CVariable)inst.data);
            }

            if (String.Equals(typeOfStep, typeof(CInstrTest).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName))
                    return new CTsTest((CVariable)inst.data);
            }

            Console.WriteLine(typeof(CInstrForce));
            throw new NotImplementedException(String.Format("Data not handled : [{0}, {1}]", typeOfStep, typeOfData));
        }
    }
}
