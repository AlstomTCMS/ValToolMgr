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
            
            if (General.WsExist(StringEnum.GetStringValue(SheetsNames.SW_VTP)))
            {
                Worksheet ws = RibbonHandler.ExcelApplication.ActiveWorkbook.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];

                //unHide Test's Column
                ws.Columns["B:B"].EntireColumn.Hidden = false;

                // Workaround : unprotect all sheets before adding a table style : https://www.google.fr/url?sa=t&rct=j&q=&esrc=s&source=web&cd=1&cad=rja&ved=0CDQQFjAA&url=http%3A%2F%2Fpradeepgali.blogspot.com%2F2013%2F02%2Fsteps-to-format-excel-table-suitable-to.html&ei=ck9cUZrnM8XiPLH7gJgF&usg=AFQjCNFR-SwI8Ns4n_ZeOB6VW3HTEB1GVg&sig2=rcSAUk2KeYCDJvWi-0cHXA
                Worksheet endpaper = RibbonHandler.ExcelApplication.ActiveWorkbook.Sheets[StringEnum.GetStringValue(SheetsNames.ENDPAPER_PR)];
                endpaper.Unprotect();
                CreateTest.AddDescTableFormat();
                endpaper.Protect();

                int indent_cat = 0;
                foreach (ListObject testlist in ws.ListObjects)
                {
                    //ListObject testsTable = ws.ListObjects[1];
                    //ListColumn testCol = testlist.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.TEST)];
                    indent_cat++;
                    int indent_test = 0;
                    string testName;
                    foreach (ListRow test in testlist.ListRows)
                    {
                        testName = TEST.TABLE.PREFIX.TEST + indent_cat + "." + ++indent_test;
                        CreateTest.createWholeTestFormat(testName);

                        //create link in SwVTP to this test sheet
                        Range tRange = test.Range[1];
                        ws.Hyperlinks.Add(tRange, "", "'" + testName + "'!A1", "Go to test " + indent_cat + "." + indent_test, indent_cat + "." + indent_test);
                    }
                    //SwVTP_Creation.formatColumnsSwVTP();

                }
                ws.Columns["B:B"].EntireColumn.AutoFit();
            }
        }
    }
}
