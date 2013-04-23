using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ValToolMgrDna.Interface;

namespace ValToolFunctions_2013
{
    public class TestSheet // : Worksheet
    {
        public CTest test;

        public ListObject loHeader;
        public ListObject loActionsTable;
        public ListObject loChecksTable;

        public TestSheet(string sheetName)
        {
            try
            {
                if (General.WsExist(sheetName))
                {
                    RibbonHandler.ExcelApplication.DisplayAlerts = false;
                    RibbonHandler.ExcelApplication.ActiveWorkbook.Worksheets[sheetName].Delete();
                    RibbonHandler.ExcelApplication.DisplayAlerts = true;
                }
            }
            catch { }

            //Ajout TEMPORAIRE d'un workbook s'il n'en existe pas
            if (!General.HasActiveBook(false))
            {
                RibbonHandler.ExcelApplication.Workbooks.Add();
            }

            Sheets sheets = (Sheets)RibbonHandler.ExcelApplication.ActiveWorkbook.Worksheets;

            //Si la feuille n'existe pas, on l'ajoute
            if (!General.WsExist(sheetName))
            {
                sheets.Add(After: sheets[sheets.Count]).Name = sheetName;
            }
            else
            {
                throw new SheetAlreadyExistException();
            }

            Worksheet testSheet = General.InitSheet(sheetName);
            testSheet.Activate();
            testSheet.Tab.ThemeColor = XlThemeColor.xlThemeColorLight2;
            testSheet.Tab.TintAndShade = 0;
            General.SetGreySheetPattern(testSheet);
            testSheet.Cells.ColumnWidth = 25;

            //AddTableDescription(testSheet);
            //AddTableAction(testSheet);
            //AddTableCheck(testSheet);
            //AddActionLabel(testSheet);
            //AddCheckLabel(testSheet);
            //FormatTestSheet(testSheet);
            //AddTestTitle(testSheet);
        }

   
    }
}
