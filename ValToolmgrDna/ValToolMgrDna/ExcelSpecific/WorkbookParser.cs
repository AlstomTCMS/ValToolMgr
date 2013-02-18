using System;
using ExcelDna.Integration;
using ValToolMgrInt;
using Excel = NetOffice.ExcelApi;

namespace ValToolMgrDna.ExcelSpecific
{
    class WorkbookParser
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static CTestContainer parseTestsOfWorkbook(Excel.Sheets sheets)
        {
            logger.Info("Begin Analysis of selected sheets");
            CTestContainer listOfTests = new CTestContainer();
            foreach (Excel.Worksheet wsCurrentTestSheet in sheets)
            {
                try
                {
                    logger.Debug(String.Format("Processing sheet \"{0}\".", wsCurrentTestSheet.Name));
                    string testNumber = getTestNumber(wsCurrentTestSheet.Name);

                    CTest result = TestSheetParser.parseTest(wsCurrentTestSheet.Name, wsCurrentTestSheet, null, TestSheetParser.PR_TEST_TABLE_ACTION_PREFIX + testNumber, TestSheetParser.PR_TEST_TABLE_CHECK_PREFIX + testNumber);

                    logger.Debug("Adding sheet to result list");
                    listOfTests.Add(result);
                }
                catch (Exception ex)
                {
                    logger.Error("Sheet cannot be parsed : ", ex);
                    XlCall.Excel(XlCall.xlcAlert, "Sheet \"" + wsCurrentTestSheet.Name + " was not analysed. Message is : "+ ex.Message); 
                }
            }
            return listOfTests;
        }

        private static string getTestNumber(string TestText)
        {
            try
            {
                logger.Debug(String.Format("Analysing \"{0}\".", TestText));
                string result = TestText.Split('_')[1];
                logger.Debug(String.Format("Found key \"{0}\".", result));
                return result;
            }
            catch (Exception)
            {
                throw new FormatException(String.Format("Sheet \"{0}\" has to begin with \"{1}\"", TestText, TestSheetParser.PR_TEST_PREFIX));
            }
            
        }
    }
}
