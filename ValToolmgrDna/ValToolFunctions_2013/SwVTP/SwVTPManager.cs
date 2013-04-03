using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using System.Text.RegularExpressions;
using ValToolFunctionsStub;
using Excel = Microsoft.Office.Interop.Excel;

namespace ValToolFunctions_2013
{
    internal class SwVTPManager
    {
        internal static void AddCategory()
        {
            if (General.isActivesheet_a_SwVTPSheet())
            {
                Worksheet ws = RibbonHandler.Factory.GetVstoObject(RibbonHandler.ExcelApplication.ActiveSheet);
                //ListObject newCategoryTestsTableT = ws.Controls.AddListObject();

                // déterminer la fin
                int lastRow = 0;
                int tableIndex = 0;
                // If it is the good sheet
                if (Regex.IsMatch(ws.Name, StringEnum.GetStringValue(SheetsNames.SW_VTP)))
                {
                    foreach (Excel.ListObject list in ws.ListObjects)
                    {
                        int row = list.Range.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                        if (row > lastRow)
                        {
                            lastRow = row;
                        }
                        try
                        {
                            int indexName = int.Parse(list.Name.Replace("TestsList_", ""));
                            if (indexName > tableIndex)
                            {
                                tableIndex = indexName;
                            }
                        }
                        catch { }
                    }

                    if (tableIndex == 0)
                    {
                        lastRow += 1;
                    }
                    // Add test list at the end with his category
                    ListObject newTestList = ws.Controls.AddListObject(ws.Range["B" + lastRow + ":F" + lastRow], "TestsList_" + ++tableIndex);
                    //Excel.ListObject newTestList = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range["B" + lastRow + ":F" + lastRow], Excel.XlYesNoGuess.xlNo);
                    //newTestList.Name = "TestsList_" + ++tableIndex;
                    //remove titles
                    newTestList.ShowHeaders = false;
                    newTestList.Range.Cut(ws.Range["B" + lastRow + ":F" + lastRow]);
                    //ungrey
                    General.UnformatGrey(newTestList.Range.EntireRow);
                }
            }
        }

    }
}
