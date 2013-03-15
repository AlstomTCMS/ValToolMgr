using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using TestStandGen.Types;
using TestStandGen.Types.Instructions;

using ValToolMgrInt;

namespace TestStandGen.Types
{
    public class CTsInstrFactory
    {
        private static Hashtable dictionnary;
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static CTsGenericInstr getTsEquivFromInstr(CInstruction inst)
        {
            string typeOfStep = inst.GetType().ToString();
            logger.Debug(String.Format("Found instruction [type={0}]", typeOfStep));

            CTsGenericInstr instr = null;

            if (String.Equals(typeOfStep, typeof(CInstrPopup).FullName))
                instr = new CTsPopup(inst.data.ToString());

            if (String.Equals(typeOfStep, typeof(CInstrWait).FullName))
                instr = new CTsWait(Convert.ToInt32(inst.data));

            if (instr == null)
                instr = basicOperatorAnalyser(inst);

            if (instr != null)
            {
                instr.Skipped = inst.Skipped;
                instr.ForceFailed = inst.ForceFailed;
                instr.ForcePassed = inst.ForcePassed;
                return instr;
            }

            throw new NotImplementedException(String.Format("Data not handled : [{0}, {1}]", typeOfStep, inst.data.GetType().ToString()));
        }

        private static CTsGenericInstr basicOperatorAnalyser(CInstruction inst)
        {
            if (!isInitialized())
                throw new NullReferenceException("CTsInstrFactory needs to be configured before use.");

            CVariable variable = (CVariable)inst.data;
            Target t = translateLocation(ref variable);
            inst.data = variable;

            if (String.Equals(t.GetType().FullName, typeof(CbTarget).FullName))
            {
                return buildCbStep(inst);
            }
            else if (String.Equals(t.GetType().FullName, typeof(TtTarget).FullName))
            {
                return buildTtStep(inst);
            }
            else
            {
                return null;
            }
        }



        public static bool isInitialized()
        {
            return dictionnary != null;
        }

        public static void loadConfiguration(string path)
        {
            logger.Info(String.Format("Loading Configuration \"{0}\"", path));
            logger.Warn("Function implementation is not complete");
            dictionnary = new Hashtable();

            XmlDocument doc = new XmlDocument();
            doc.Load(path);

            XmlNodeList TargetDefinitions = doc.SelectSingleNode("/Configuration/TargetDefinitions").SelectNodes("Target");
            Hashtable targetTable = new Hashtable();
            foreach (XmlNode node in TargetDefinitions)
            {
                Target target;

                string value = node.Attributes.GetNamedItem("testStandIdentifier").Value;
                string value2 = node.Attributes.GetNamedItem("trainTracerIdentifier").Value;
                if (value != null)
                    target = new CbTarget(node.Attributes.GetNamedItem("name").Value, value);
                else if (value2 != null)
                    target = new TtTarget(node.Attributes.GetNamedItem("name").Value, value2);
                else throw new FormatException("No field found valid for configuration file");

                targetTable.Add(target.name, target);
            }

            XmlNodeList LocationDefinitions = doc.SelectSingleNode("/Configuration/LocationDefinitions").SelectNodes("Location");

            foreach (XmlNode node in LocationDefinitions)
            {
                string name = node.Attributes.GetNamedItem("name").Value;
                string targetConfig = node.Attributes.GetNamedItem("targetConfig").Value;
                dictionnary.Add(name, targetTable[targetConfig]);
            }
        }

        private static Target translateLocation(ref CVariable variable)
        {
            if (dictionnary.Contains(variable.Location))
            {
                Target t = (Target)dictionnary[variable.Location];
                variable.Location = t.Identifier;
                return t;
            }
            else
            {
                string message = String.Format("Requested Location \"{0}\" is not defined inside configuration file", variable.Location);
                logger.Error(message);
                throw new FormatException(message);
            }


        }

        private static CTsGenericInstr buildTtStep(CInstruction inst)
        {

            CTsGenericInstr instr = null;
            string typeOfStep = inst.GetType().ToString();
            string typeOfData = inst.data.GetType().ToString();
            CVariable variable = (CVariable)inst.data;

            if (String.Equals(typeOfStep, typeof(CInstrUnforce).FullName) && !String.Equals(typeOfData, typeof(CVariableArray).FullName))
            {
                return new CTsTtUnforce(variable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrForce).FullName))
            {
                return new CTsTtForce(variable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrTest).FullName))
            {
                return new CTsTtTest(variable);
            }

            return instr;
        }

        abstract class Target
        {
            public string name;
            public string Identifier { get; set; }

            public Target(String Name, string Identifier)
            {
                name = Name;
                this.Identifier = Identifier;
            }
        }

        class CbTarget : Target
        {
            public CbTarget(string p, string value)
                : base(p, value)
            {
            }
        }

        class TtTarget : Target
        {
            public TtTarget(string p, string value)
                : base(p, value)
            {
            }
        }

        class Location
        {
            public string name { get; set; }

            public string strategy { get; set; }

            public string targetConfig { get; set; }
        }

        private static CTsGenericInstr buildCbStep(CInstruction inst)
        {
            CTsGenericInstr instr = null;
            string typeOfStep = inst.GetType().ToString();
            string typeOfData = inst.data.GetType().ToString();
            CVariable variable = (CVariable)inst.data;

            if (String.Equals(typeOfStep, typeof(CInstrUnforce).FullName) && !String.Equals(typeOfData, typeof(CVariableArray).FullName))
            {
                instr = new CTsUnforce(variable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrForce).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName) || String.Equals(typeOfData, typeof(CVariableUInt).FullName) || String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    instr = new CTsForce(variable);

                if (String.Equals(typeOfData, typeof(CVariableArray).FullName))
                    instr = new CTsForceArray((CVariableArray)inst.data);
            }

            if (String.Equals(typeOfStep, typeof(CInstrTest).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName) || String.Equals(typeOfData, typeof(CVariableUInt).FullName))
                    instr = new CTsTest((CVariable)inst.data);

                if (String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    instr = new CTsTestAna((CVariableDouble)inst.data);

                if (String.Equals(typeOfData, typeof(CVariableArray).FullName))
                    instr = new CTsTestArray((CVariableArray)inst.data);
            }
            return instr;
        }
    }
}
