using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ValToolFunctions_2013
{
    internal class SwVTP_Creation
    {
        /// <summary>
        /// Ask user for an PR name and create a new PR file with an empty SwVTP, the Bench conf sheet and 
        /// </summary>
        internal static void NewPR()
        {
            try
            {
                string FILENAME_PATTERN = "B2_XXX_Y_A0";
                string PRname = FILENAME_PATTERN;
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                Stream myStream;

                saveFileDialog1.Filter = "Excel Worksheets (*.xlsx)|*.xlsx"; //"Fichier Excel (*.xls)|*.xls|(*.xlsx)|*.xlsx|All   files (*.*)|*.*"
                saveFileDialog1.Title = "Save an Excel File";
                saveFileDialog1.FileName = FILENAME_PATTERN;
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.InitialDirectory = "C:\\Files";
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filename = saveFileDialog1.FileName;
                    if (filename != "" && filename != FILENAME_PATTERN)
                    {
                        SaveExcelFile(PRname);
                    }
                }
            }
            catch { }
        }

        private static void SaveExcelFile(string fileName)
        {
            Excel.Application app = RibbonHandler.ExcelApplication;
            Workbook wb = app.Workbooks.Add(Type.Missing);

            //Add init sheets
            CreateEndpaperSheet(wb);
            CreateEvolSheet(wb);
            CreateBenchConfSheet(wb);
            CreateSwVTPSheet(wb);

            //Save file and show it
            app.DisplayAlerts = false;
            wb.SaveAs(fileName);
            app.DisplayAlerts = true;
            wb.Saved = true;
        }

        /// <summary>
        /// Create the "Page de garde" of the book
        /// </summary>
        /// <param name="wb"></param>
        private static void CreateEndpaperSheet(Workbook wb)
        {
            Worksheet eps = wb.Sheets[1];
            eps.Name = "Endpaper";
        }

        private static void CreateEvolSheet(Workbook wb)
        {
            Worksheet es = wb.Sheets[2];
            es.Name = "Evol";

            Interior int_all = es.Cells.Interior;
            int_all.Pattern = XlPattern.xlPatternSolid;
            int_all.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            int_all.ThemeColor = XlThemeColor.xlThemeColorDark1;
            int_all.TintAndShade = -0.349986266670736;
            int_all.PatternTintAndShade = 0;

            es.Activate();
            RibbonHandler.ExcelApplication.ActiveWindow.DisplayGridlines = false;

            ListObject evolTable = es.ListObjects.Add(XlListObjectSourceType.xlSrcRange, es.Range["A1:D1"], XlYesNoGuess.xlYes);

            es.Range["A1:D1"].Value = new String[] { "Version", "Date", "Name", "Modification" };
            evolTable.Name = "evolTable";
            evolTable.TableStyle = "TableStyleMedium2";

            Interior int_evol = evolTable.Range.Interior;
            int_evol.Pattern = XlPattern.xlPatternNone;
            int_evol.TintAndShade = 0;
            int_evol.PatternTintAndShade = 0;

            //Init filling
            es.Range["A2:D2"].Value = new String[] { "A0", DateTime.Now.ToString(), Environment.UserName, "Creation" };

            evolTable.Range.EntireColumn.AutoFit();
            evolTable.ListColumns[3].Range.ColumnWidth = 15;
            evolTable.ListColumns[4].Range.ColumnWidth = 60;
        }

        private static void CreateBenchConfSheet(Workbook wb)
        {
            Worksheet bcs = wb.Sheets[3];
            bcs.Name = "Bench Conf";
        }

        private static void CreateSwVTPSheet(Workbook wb)
        {
            Worksheet SwvtpS = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            SwvtpS.Name = "SwVTP";
        }
    }
}
