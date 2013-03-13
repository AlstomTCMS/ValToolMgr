using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace TestStandGen
{
    public class CTestStandLocatorAdapter
    {
        private static Hashtable dictionnary;
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

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
                CbTarget target = new CbTarget();
                target.name = node.Attributes.GetNamedItem("name").Value;
                target.testStandIdentifier = node.Attributes.GetNamedItem("testStandIdentifier").Value;
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

        public static void translateLocation(ref string Location, ref string Path)
        {
            if (dictionnary.Contains(Location))
            {
                logger.Warn("Function implementation is not complete");
                Location = ((CbTarget)dictionnary[Location]).testStandIdentifier;
                logger.Debug(String.Format("Found following Location : \"{0}\"", Location));
            }
            else
            {
                string message = String.Format("Requested Location \"{0}\" is not defined inside configuration file", Location);
                logger.Error(message);
                throw new FormatException(message);
            }
        }

        class CbTarget
        {
            public string name;
            public string testStandIdentifier { get; set; }
        }

        class Location
        {
            public string name { get; set; }

            public string strategy { get; set; }

            public string targetConfig { get; set; }
        }
    }
}
