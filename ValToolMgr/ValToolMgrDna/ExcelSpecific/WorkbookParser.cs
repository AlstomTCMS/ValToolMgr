using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NetOffice;
using ExcelDna.Integration;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.ExcelApi.GlobalHelperModules;
using ValToolMgrDna.Interface;

namespace ValToolMgrDna.ExcelSpecific
{
    public class WorkbookParser
    {
        public static CTestContainer parseTestsOfWorkbook(Excel.Sheets sheets)
        {
            CTestContainer listOfTests = new CTestContainer();
            foreach (Excel.Worksheet wsCurrentTestSheet in sheets)
            {
                try
                {
                    string testNumber = getTestNumber(wsCurrentTestSheet.Name);
                    Excel.ListObject loActionsTable = wsCurrentTestSheet.ListObjects[TestSheetParser.PR_TEST_TABLE_ACTION_PREFIX + testNumber];
                    Excel.ListObject loChecksTable = wsCurrentTestSheet.ListObjects[TestSheetParser.PR_TEST_TABLE_CHECK_PREFIX + testNumber];

                    CTest result = TestSheetParser.parseTest(wsCurrentTestSheet.Name, wsCurrentTestSheet, null, loActionsTable, loChecksTable);
                    listOfTests.Add(result);
                }
                catch (Exception)
                {
                    XlCall.Excel(XlCall.xlcAlert, "Sheet \"" + wsCurrentTestSheet.Name + " was not analysed."); 
                }
            }
            return listOfTests;
        }

        private static string getTestNumber(string TestText)
        {
            return TestText.Split('_')[1];
        }

    }
}
