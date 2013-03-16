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
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private bool alreadyGenerated;
        private TestStandFile TSFile;

        public static void genSequence(CTestContainer sequence, string outFile, string templatePath)
        {
            CTsInstrFactory.loadConfiguration("C:\\macros_alstom\\Configuration\\LocationConfiguration.xml");

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
                        logger.Error(m);
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
                logger.Info("Processing sequence");
                CTestStandSeq SubSeq = genInstrListFromTest(test);

                ts.addSequence("Call to subsequence " + SubSeq.identifier, SubSeq);
                logger.Debug("End of sequence Processing");
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
                logger.Debug("Processing step");
                SubSeq.Add(new CTsLabel("===================================="));
                SubSeq.Add(new CTsLabel("========  " + step.title));
                SubSeq.Add(new CTsLabel("===================================="));

                if (step.DescAction.Length > 0) SubSeq.Add(new CTsLabel("= Actions : " + step.DescAction));
                logger.Debug("Processing actions.");
                foreach(CInstruction instr in step.actions)
                {
                    logger.Debug("Processing action.");
                     SubSeq.Add(CTsInstrFactory.getTsEquivFromInstr(instr));
                     logger.Debug("End of action processing.");
                }

                if (step.DescCheck.Length > 0) SubSeq.Add(new CTsLabel("= Checks : " + step.DescCheck));
                foreach (CInstruction instr in step.checks)
                {
                    logger.Debug("Processing check.");
                    SubSeq.Add(CTsInstrFactory.getTsEquivFromInstr(instr));
                    logger.Debug("End of check processing.");
                }
            }
            return SubSeq;
        }

    }
}
