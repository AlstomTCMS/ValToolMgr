using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using ValToolMgrInt;
using ValToolMgrDna.ExcelSpecific;

namespace ValToolMgrTest
{
    [TestFixture]
    class VariableParserTest
    {
        [Test]
        public void testVariableAsBoolean()
        {
            string Path = "/this/is/my/path";
            string Variable = "mySimpleVariable";

            testIfVariableAsBoolean(Variable, Path, "1", true);
            testIfVariableAsBoolean(Variable, Path, "TRUE", true);
            testIfVariableAsBoolean(Variable, Path, "true", true);
            testIfVariableAsBoolean(Variable, Path, "True", true);

            testIfVariableAsBoolean(Variable, Path, "0", false);
            testIfVariableAsBoolean(Variable, Path, "false", false);
            testIfVariableAsBoolean(Variable, Path, "FALSE", false);
            testIfVariableAsBoolean(Variable, Path, "False", false);
        }

        [Test]
        public void testVariableAsInteger()
        {
            string Path = "/this/is/my/path";
            string Variable = "I:mySimpleVariable";

            testIfVariableAsInteger(Variable, Path, "10", 10);
            testIfVariableAsInteger(Variable, Path, "-255", -255);
            testIfVariableAsInteger(Variable, Path, "0", 0);
            testIfVariableAsInteger(Variable, Path, "-0", 0);
        }

        [Test]
        public void testVariableAsReal()
        {
            string Path = "/this/is/my/path";
            string Variable = "R:mySimpleVariable";

            testIfVariableAsReal(Variable, Path, "10", 10);
            testIfVariableAsReal(Variable, Path, "-255", -255);
            testIfVariableAsReal(Variable, Path, "0", 0);
            testIfVariableAsReal(Variable, Path, "-0,0", 0);
            testIfVariableAsReal(Variable, Path, "-0.0", 0);
            testIfVariableAsReal(Variable, Path, "-0.2", -0.2);
            testIfVariableAsReal(Variable, Path, "-0,2", -0.2);
            testIfVariableAsReal(Variable, Path, "0.2", 0.2);
            testIfVariableAsReal(Variable, Path, "1000.02", 1000.02);
            testIfVariableAsReal(Variable, Path, "1000.0", 1000);
        }

        #region Test details

        private void testIfVariableAsBoolean(string Variable, string Path, string Value, bool expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, Path, Value);

            Assert.IsNotNull(var);
            
            Assert.IsInstanceOf<CVariableBool>(var);
            Assert.IsTrue(var.value is bool);
            Assert.AreEqual(var.value, expectedValue);
        }

        private void testIfVariableAsInteger(string Variable, string Path, string Value, int expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, Path, Value);

            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableInt>(var);
            Assert.IsTrue(var.value is Int32);
            Assert.AreEqual(var.value, expectedValue);
        }

        private void testIfVariableAsReal(string Variable, string Path, string Value, double expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, Path, Value);

            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableDouble>(var);
            Assert.IsTrue(var.value is double);
            Assert.AreEqual(var.value, expectedValue);
        }

        #endregion
    }
        
}
