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

        [Test]
        public void testVariableAsBooleanArray()
        {
            string Path = "/this/is/my/path";
            string Variable = "mySimpleVariable[0]";

            testIfVariableAsBooleanArray(Variable, Path, "1", 0, true);
            testIfVariableAsBooleanArray(Variable, Path, "TRUE", 0, true);
            testIfVariableAsBooleanArray(Variable, Path, "true", 0, true);
            testIfVariableAsBooleanArray(Variable, Path, "True", 0, true);

            testIfVariableAsBooleanArray(Variable, Path, "0", 0, false);
            testIfVariableAsBooleanArray(Variable, Path, "false", 0, false);
            testIfVariableAsBooleanArray(Variable, Path, "FALSE", 0, false);
            testIfVariableAsBooleanArray(Variable, Path, "False", 0, false);
        }

        [Test]
        public void testVariableAsIntegerArray()
        {
            string Path = "/this/is/my/path";
            string Variable = "I:mySimpleVariable[9]";

            testIfVariableAsIntegerArray(Variable, Path, "10", 9, 10);
            testIfVariableAsIntegerArray(Variable, Path, "-255", 9, -255);
            testIfVariableAsIntegerArray(Variable, Path, "0", 9, 0);
            testIfVariableAsIntegerArray(Variable, Path, "-0", 9, 0);
        }

        [Test]
        public void testVariableAsRealArray()
        {
            string Path = "/this/is/my/path";
            string Variable = "R:mySimpleVariable[3]";

            testIfVariableAsRealArray(Variable, Path, "10", 3, 10);
            testIfVariableAsRealArray(Variable, Path, "-255", 3, -255);
            testIfVariableAsRealArray(Variable, Path, "0", 3, 0);
            testIfVariableAsRealArray(Variable, Path, "-0,0", 3, 0);
            testIfVariableAsRealArray(Variable, Path, "-0.0", 3, 0);
            testIfVariableAsRealArray(Variable, Path, "-0.2", 3, -0.2);
            testIfVariableAsRealArray(Variable, Path, "-0,2", 3, -0.2);
            testIfVariableAsRealArray(Variable, Path, "0.2", 3, 0.2);
            testIfVariableAsRealArray(Variable, Path, "1000.02", 3, 1000.02);
            testIfVariableAsRealArray(Variable, Path, "1000.0", 3, 1000);
        }

        #region Test details

        private void testIfVariableAsBoolean(string Variable, string Path, string Value, bool expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, "SECTION1/ENV", Path);
            testIfVariableAsBoolean(var, expectedValue);
        }

        private void testIfVariableAsBoolean(CVariable var, bool expectedValue)
        {
            Assert.IsNotNull(var);
            
            Assert.IsInstanceOf<CVariableBool>(var);
            //Assert.IsTrue(var.value is bool);
            //Assert.AreEqual(var.value, expectedValue);
        }

        private void testIfVariableAsInteger(string Variable, string Path, string Value, int expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, "SECTION1/ENV", Path);
            testIfVariableAsInteger(var, expectedValue);

        }

        private void testIfVariableAsInteger(CVariable var, int expectedValue)
        {
            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableInt>(var);
            //Assert.IsTrue(var.value is Int32);
            //Assert.AreEqual(var.value, expectedValue);
        }

        private void testIfVariableAsReal(string Variable, string Path, string Value, double expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, "SECTION1/ENV", Path);
            testIfVariableAsReal(var, expectedValue);
        }

        private void testIfVariableAsReal(CVariable var, double expectedValue)
        {
            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableDouble>(var);
            //Assert.IsTrue(var.value is double);
            //Assert.AreEqual(var.value, expectedValue);
        }

        private void testIfVariableAsBooleanArray(string Variable, string Path, string Value, uint index, bool expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, "SECTION1/ENV", Path);

            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableArray>(var);
            //Assert.IsInstanceOf<CVariableBool>(var.value);
 
            CVariableArray array = (CVariableArray)var;
            Assert.AreEqual(index, array.Index);

            //testIfVariableAsBoolean((CVariableBool)var.value, expectedValue);
        }

        private void testIfVariableAsIntegerArray(string Variable, string Path, string Value, uint index, int expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, "SECTION1/ENV", Path);

            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableArray>(var);
            //Assert.IsInstanceOf<CVariableInt>(var.value);

            CVariableArray array = (CVariableArray)var;
            Assert.AreEqual(index, array.Index);

            //testIfVariableAsInteger((CVariableInt)var.value, expectedValue);
        }

        private void testIfVariableAsRealArray(string Variable, string Path, string Value, uint index, double expectedValue)
        {
            CVariable var = VariableParser.parseAsVariable(Variable, "SECTION1/ENV", Path);

            Assert.IsNotNull(var);
            Assert.IsInstanceOf<CVariableArray>(var);

            CVariableArray array = (CVariableArray)var;
            Assert.AreEqual(index, array.Index);

            //Assert.IsInstanceOf<CVariableDouble>(var.value);
            //testIfVariableAsReal((CVariableDouble)var.value, expectedValue);
        }

        #endregion
    }
        
}
