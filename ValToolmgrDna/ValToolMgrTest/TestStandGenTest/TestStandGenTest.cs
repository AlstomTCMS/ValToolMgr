using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using ValToolMgrDna.Interface;
using TestStandGen;



namespace ValToolMgrTest
{
        [TestFixture]
        public class TestStandGenTest
        {
            [Test]
            public void GenerateScenario()
            {
                CTestContainer container = new CTestContainer();
                container.description = "Test container";
                for (int testIndex = 1; testIndex < 3; testIndex++)
                {
                    CTest test = new CTest();
                    test.description = "Test descriptor #" + testIndex;
                    test.title = "TEST_" + testIndex;

                    for (int stepIndex = 1; stepIndex < 20; stepIndex++)
                    {
                        CStep step = new CStep();
                        step.title = "Step " + testIndex + "." + stepIndex;
                        step.DescAction = "Action description for " + step.title;
                        step.DescCheck = "Check description for " + step.title;

                        for (int actionIndex = 1; actionIndex < 20; actionIndex++)
                        {
                            CInstruction action = new CInstruction();
                            action.category = CInstruction.actionList.A_FORCE;
                            CVariableBool var = new CVariableBool();
                            var.value = "true";
                            var.name = "Var" + actionIndex;
                            var.path = "/path/to/application" + actionIndex;
                            action.data = var;
                            step.actions.Add(action);
                        }

                        for (int checkIndex = 1; checkIndex < 20; checkIndex++)
                        {
                            CInstruction action = new CInstruction();
                            action.category = CInstruction.actionList.A_TEST;
                            CVariableBool var = new CVariableBool();
                            var.value = "true";
                            var.name = "Var" + checkIndex;
                            var.path = "/path/to/application" + checkIndex;
                            action.data = var;
                            step.checks.Add(action);
                        }
                        test.Add(step);
                    }
                    container.Add(test);
                }

                TestStandGen.TestStandGen.genSequence(container, "C:\\macros_alstom\\test\\genTest.seq", "C:\\macros_alstom\\templates\\ST-TestStand3\\");

                Assert.IsTrue(true);
            }
        }
}
