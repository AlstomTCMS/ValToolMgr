using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ValToolFunctionsStub;

namespace ValToolFunctions_2013
{
    internal static class TestsGenerator
    {
        /// <summary>
        /// Create tests sheets from the SwVTP sheet
        /// </summary>
        internal static void FromSwVTP2Tests()
        {
            //TODO: trouver un mecanisme de protection (bouton inutilisable si pas de page SwVTP avec un formalisme convaincant)
            Worksheet swVtpS = RibbonHandler.ExcelApplication.ActiveWorkbook.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            ListObject testsTable = swVtpS.ListObjects[1];

            ListColumn testCol = testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.TEST)];
            testCol.Range.EntireColumn.Hidden = false;

            int indent_test = 0;
            string testName;
            foreach (ListRow test in testsTable.ListRows)
            {
                testName = TEST.TABLE.PREFIX.TEST + ++indent_test;
                CreateTest.createWholeTestFormat(testName);

                //create link in SwVTP to this test sheet
                Range tRange = test.Range[1];
                swVtpS.Hyperlinks.Add(tRange, "", "'" + testName + "'!A1", "Go to test " + indent_test, indent_test + "");
            }
            //testCol.Range.EntireColumn.AutoFit();
            //SwVTP_Creation.formatColumnsSwVTP();
        }
    }
}
