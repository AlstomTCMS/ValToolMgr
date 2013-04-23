using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using ValToolMgrInt;
using TestStandGen;
using TestStandGen.Types;

namespace ValToolMgrTest
{
    [TestFixture]
    class TestStandLocationTest
    {
        [Test]
        public void loadConfiguration()
        {
            CTsInstrFactory.loadConfiguration("C:\\macros_alstom\\Configuration\\LocationConfiguration.xml");
        }
    }
}
