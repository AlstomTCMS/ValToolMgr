using System;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Antlr4.StringTemplate;
using System.Collections.Generic;
using Antlr4.StringTemplate.Misc;
using TestStandGen;

namespace TestStandGen
{

    abstract class User
    {
        public User(string name, string phone)
        {
            this.Name = name;
            this.Phone = phone;
        }

        public string Name;
        public string Phone;
    }

    class Man : User
    {
        public Man(string name, string phone)
            : base(name, phone)
        {
        }

        public string InstructionName = typeof(Man).Name;
    }

    class Woman : User
    {
        public Woman(string name, string phone)
            : base(name, phone)
        {
        }

        public string InstructionName = typeof(Woman).Name;
    }

    [TestFixture]
    class TestDictionary
    {
        string newline = Environment.NewLine;

        [Test]
        public void GenerateScenario()
        {
            
            IDictionary<User, string> m = new Dictionary<User, string>();
            TemplateGroup group = new TemplateGroupFile( "C:\\macros_alstom\\ValToolmgrDna\\ValToolMgrTest\\TestDictionnary.stg");
            Template st = group.GetInstanceOf("MainSequence");
            st.impl.Dump();
            m.Add(new Woman("Toto 3", "1234"), "value1");
            m.Add(new Man("Toto 1", "1234"), "value2");
            m.Add(new Woman("Toto 2", "1234"), "value3");
            st.Add("items", m);
            string expecting = "int x = 0L;";
            string result = st.Render();

            Console.WriteLine("==========================");
            Console.Write(result);
            Console.WriteLine("==========================");
            Console.ReadKey();


            Assert.AreSame(expecting, expecting);
        }
    }
}
