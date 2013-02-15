﻿using System;
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
            public void GenerateAllSteps()
            {
                CTestContainer container = new CTestContainer();

                CTest test = new CTest("Test_1", "This is my description");

                CStep step = new CStep("Step 1", null, null);
            }

            [Test]
            public void GenerateScenario()
            {
                CTestContainer container = new CTestContainer();
                container.description = "Test container";
                for (int testIndex = 1; testIndex <= 3; testIndex++)
                {
                    CTest test = new CTest("Test_1." + testIndex, "Test descriptor #" + testIndex);

                    for (int stepIndex = 1; stepIndex < 2; stepIndex++)
                    {
                        string title = "Step " + testIndex + "." + stepIndex;
                        CStep step = new CStep(
                            title, 
                             "Action description for " + title,
                             "Check description for " + title
                            );

                        for (int actionIndex = 1; actionIndex < 10; actionIndex++)
                        {
                            CInstruction action = new CInstrForce();
                            CVariableBool var = new CVariableBool();
                            var.value = "true";
                            var.name = "Var" + actionIndex;
                            var.path = "/path/to/application" + actionIndex;
                            action.data = var;
                            step.actions.Add(action);
                        }

                        for (int checkIndex = 1; checkIndex < 10; checkIndex++)
                        {
                            CInstruction action = new CInstrTest();
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
