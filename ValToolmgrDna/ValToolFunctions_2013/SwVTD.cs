using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolFunctionsStub;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ValToolFunctions_2013
{
    internal static class SwVTD
    {
        internal static void GenerateSwVTD()
        {
            Excel.Application app = RibbonHandler.ExcelApplication;
            Workbook Pr_wb = app.ActiveWorkbook;

            //copy SwVTP in a new SwVTD sheet. Delete it if it already exist
            if (General.WsExist(StringEnum.GetStringValue(SheetsNames.SW_VTD)))
            {
                Pr_wb.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTD)].Delete();
            }
            Worksheet swVTP = Pr_wb.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            swVTP.Copy(After: Pr_wb.Sheets[swVTP.Index]);
            Worksheet swVtdS = Pr_wb.Sheets[swVTP.Index + 1];
            swVtdS.Name = StringEnum.GetStringValue(SheetsNames.SW_VTD);


            //unhide 2 columns
            swVtdS.Columns[SwVTx_Columns.STEP].EntireColumn.Hidden = false;
            swVtdS.Columns[SwVTx_Columns.CHECK].EntireColumn.Hidden = false;

            //Modify test title zone
            Range testTitle = swVtdS.Range["B1:F1"];
            testTitle.UnMerge();
            testTitle = swVtdS.Range["B1:D1"];
            testTitle.Merge();

            //Add Step Zone
            Range StepTitle = swVtdS.Range["E1:H1"];
            StepTitle.Merge();
            StepTitle.Value = "STEP";

            //Fill SwVTD with tests's sheets data
            //swVtdS.



            ////create a new wb
            //app.SheetsInNewWorkbook = 1;
            //Workbook wb = app.Workbooks.Add(Type.Missing);

            //wb.Sheets[1].name = StringEnum.GetStringValue(SheetsNames.SW_VTD);

            ////Save file and show it
            //app.DisplayAlerts = false;
            //string swVTD_ext = "-" + StringEnum.GetStringValue(SheetsNames.SW_VTD) + ".";
            //wb.SaveAs(@Pr_wb.FullName.Replace(".", swVTD_ext));
            //app.DisplayAlerts = true;
            //wb.Saved = true;
        }
    }
}
