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

            CTsVariable variable = translateLocation((CVariable)inst.data);

            if (String.Equals(variable.GetType().FullName, typeof(CTsCbVariable).FullName))
            {
                return buildCbStep(inst, variable);
            }
            else if (String.Equals(variable.GetType().FullName, typeof(CTsTtVariable).FullName))
            {
                return buildTtStep(inst, variable);
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

                XmlNode testStandIdentifierNode = node.Attributes.GetNamedItem("testStandIdentifier");
                XmlNode trainTracerIdentifierNode = node.Attributes.GetNamedItem("traintracerIdentifier");
                if (testStandIdentifierNode != null)
                    target = new CbTarget(node.Attributes.GetNamedItem("name").Value, testStandIdentifierNode.Value);
                else if (trainTracerIdentifierNode != null)
                    target = new TtTarget(node.Attributes.GetNamedItem("name").Value, trainTracerIdentifierNode.Value);
                else throw new FormatException("No field found valid for configuration file");

                targetTable.Add(target.name, target);
            }

            XmlNodeList LocationDefinitions = doc.SelectSingleNode("/Configuration/LocationDefinitions").SelectNodes("Location");

            foreach (XmlNode node in LocationDefinitions)
            {
                string name = node.Attributes.GetNamedItem("name").Value;
                string targetConfig = node.Attributes.GetNamedItem("targetConfig").Value;

                Target t = (Target)targetTable[targetConfig];
                if(String.Equals(t.GetType().FullName, typeof(TtTarget).FullName))
                {
                    ((TtTarget)t).prefix = node.Attributes.GetNamedItem("pathPrefix").Value;
                }
                dictionnary.Add(name, t);
            }
        }

        private static CTsVariable translateLocation(CVariable variable)
        {
            if (dictionnary.Contains(variable.Location))
            {
                Target t = (Target)dictionnary[variable.Location];
                if (String.Equals(t.GetType().FullName, typeof(CbTarget).FullName))
                {
                    return new CTsCbVariable((CbTarget)t, variable);
                }
                else if (String.Equals(t.GetType().FullName, typeof(TtTarget).FullName))
                {
                    return new CTsTtVariable((TtTarget)t, variable);
                }
                else
                {
                    throw new FormatException("This type of target is not managed");
                }
            }
            else
            {
                string message = String.Format("Requested Location \"{0}\" is not defined inside configuration file", variable.Location);
                logger.Error(message);
                throw new FormatException(message);
            }


        }

        private static CTsGenericInstr buildTtStep(CInstruction inst, CTsVariable TsVariable)
        {

            CTsGenericInstr instr = null;
            string typeOfStep = inst.GetType().ToString();
            string typeOfData = inst.data.GetType().ToString();
            CVariable variable = (CVariable)inst.data;

            if (String.Equals(typeOfStep, typeof(CInstrUnforce).FullName) && !String.Equals(typeOfData, typeof(CVariableArray).FullName))
            {
                return new CTsTtUnforce(TsVariable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrForce).FullName))
            {
                return new CTsTtForce(TsVariable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrTest).FullName))
            {
                return new CTsTtTest(TsVariable);
            }

            return instr;
        }

        public abstract class Target
        {
            public string name;
            public string Identifier { get; set; }

            public Target(String Name, string Identifier)
            {
                name = Name;
                this.Identifier = Identifier;
            }
        }

        public class CbTarget : Target
        {
            public CbTarget(string p, string value)
                : base(p, value)
            {
            }
        }

        public class TtTarget : Target
        {
            public string prefix;

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

        private static CTsGenericInstr buildCbStep(CInstruction inst, CTsVariable TsVariable)
        {
            CTsGenericInstr instr = null;
            string typeOfStep = inst.GetType().ToString();
            string typeOfData = inst.data.GetType().ToString();
            CVariable variable = (CVariable)inst.data;

            if (String.Equals(typeOfStep, typeof(CInstrUnforce).FullName) && !String.Equals(typeOfData, typeof(CVariableArray).FullName))
            {
                instr = new CTsUnforce(TsVariable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrForce).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName) || String.Equals(typeOfData, typeof(CVariableUInt).FullName) || String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    instr = new CTsForce(TsVariable);

                if (String.Equals(typeOfData, typeof(CVariableArray).FullName))
                    instr = new CTsForceArray(TsVariable);
            }

            if (String.Equals(typeOfStep, typeof(CInstrTest).FullName))
            {
                if (String.Equals(typeOfData, typeof(CVariableBool).FullName) || String.Equals(typeOfData, typeof(CVariableInt).FullName) || String.Equals(typeOfData, typeof(CVariableUInt).FullName))
                    instr = new CTsTest(TsVariable);

                if (String.Equals(typeOfData, typeof(CVariableDouble).FullName))
                    instr = new CTsTestAna(TsVariable);

                if (String.Equals(typeOfData, typeof(CVariableArray).FullName))
                    instr = new CTsTestArray(TsVariable);
            }
            return instr;
        }
    }
}
