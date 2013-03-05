using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using ValToolFunctionsStub;
using System.Text.RegularExpressions;

namespace ValToolFunctions_2013
{
    internal class SwVTP_Creation
    {
        /// <summary>
        /// Ask user for an PR name and create a new PR file with an empty SwVTP, the Bench conf sheet and 
        /// </summary>
        internal static void NewPR()
        {
            string FILENAME_PATTERN = "S5_XXX_Y_A0";
            string PRname = FILENAME_PATTERN;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

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
            eps.Name = StringEnum.GetStringValue(SheetsNames.ENDPAPER);

            eps.Range["B3:B10"].Value = new String[] { "Function", "Num_PR", "Indice_PR", "Date_PR", 
                                    "Ref_FRScc", "Ind_FRScc", "Versions MPU", "Aim of the function"};


            eps.Range["B2"].Value = "";//D4&" "&D5&" - "&D3;
        }

        private static void CreateEvolSheet(Workbook wb)
        {
            Worksheet es = wb.Sheets[2];
            es.Name = StringEnum.GetStringValue(SheetsNames.EVOLUTION);

            SetSheetFormatPattern(es);

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
            evolTable.ListColumns[3].Range.ColumnWidth = 15; //Name
            evolTable.ListColumns[4].Range.ColumnWidth = 60; //Modif        
        }

        private static void CreateBenchConfSheet(Workbook wb)
        {
            Worksheet bcs = wb.Sheets[3];
            bcs.Name = StringEnum.GetStringValue(SheetsNames.BENCH_CONF);


        }

        private static void CreateSwVTPSheet(Workbook wb)
        {
            Worksheet SwvtpS = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            SwvtpS.Name = StringEnum.GetStringValue(SheetsNames.SW_VTP);
            SetSheetFormatPattern(SwvtpS);

            ListObject testsTable = SwvtpS.ListObjects.Add(XlListObjectSourceType.xlSrcRange, SwvtpS.Range["B1:E1"], XlYesNoGuess.xlYes);
            testsTable.Name = "TestsList_0"; 
            SwvtpS.Range["A1:E1"].Value = new String[] { StringEnum.GetStringValue(SwVTP_Columns.CATEGORY), 
                                                        StringEnum.GetStringValue(SwVTP_Columns.BENCH_CONF), 
                                                        StringEnum.GetStringValue(SwVTP_Columns.REQUIREMENT),
                                                        StringEnum.GetStringValue(SwVTP_Columns.DESC),
                                                        StringEnum.GetStringValue(SwVTP_Columns.COMMENT) };

            testsTable.TableStyle = "TableStyleMedium2";
            testsTable.Range.EntireColumn.AutoFit();
            formatColumnsSwVTP();

            Interior cat_int = SwvtpS.Range["A1:E2"].Interior;
            cat_int.Pattern = XlPattern.xlPatternNone;
            cat_int.TintAndShade = 0;
            cat_int.PatternTintAndShade = 0;

            Interior cat_title_int = SwvtpS.Range["A1"].Interior;
            cat_title_int.Pattern = XlPattern.xlPatternSolid;
            cat_title_int.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            cat_title_int.ThemeColor = XlThemeColor.xlThemeColorAccent1;
            cat_title_int.TintAndShade = 0;
            cat_title_int.PatternTintAndShade = 0;
            Font font = SwvtpS.Range["A1"].Font;
            font.Bold = true;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;

            SwvtpS.Rows["1:1"].VerticalAlignment = XlVAlign.xlVAlignCenter;

            //DebugFillingSwVTP(wb);
        }


        internal static void formatColumnsSwVTP()
        {
            Worksheet ws = RibbonHandler.ExcelApplication.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            ListObject testsTable = ws.ListObjects[1];
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.REQUIREMENT)].Range.ColumnWidth = 26; //Requirements
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.DESC)].Range.ColumnWidth = 24; //Description
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.COMMENT)].Range.ColumnWidth = 20; //Comment
        }

        private static void DebugFillingSwVTP(Workbook wb)
        {
            Worksheet SwvtpS = wb.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            //System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SwVTP_Creation));


            string input = Properties.Resources.SwVTP_stub; //"A2, \"FME-F12_040\nFME-F12_031\",Surveillance du CC-VE(URG)\r\nA2,tFME-F12_041,tSurveillance du CC-Q(URG)\t\r\nA2\tFME-F12_042\tSurveillance du CC1-Q(URG)\t\r\nA2\tFME-F12_043\tSurveillance du CC-Q-FU\t\r\nA2\tFME-F12_044\tSurveillance du CC(CFG)F\t\r\nA2\tFME-F12_046\tSurveillance du RB(IS)VV1(URG)\t\r\nA2\tFME-F12_047\tSurveillance du RB(IS)VV(RD)URG\t\r\nA2\tFME-F12_050\tSurveillance du BP1(URG)\tPrise en compte de la relecture du CC. A mettre à jour avec les bons noms de variables\r\n";// 
            string input2 = Regex.Replace(input, @"\t", ",");
            //string[,] text = new string[2, 2] { { "A2", "FME-F12_040" }, { "A3", "FME-F12_041" } };

            string[] text = Regex.Split(input2, ",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            SwvtpS.Range["B2:C3"].Value = text;
        }

        /// <summary>
        /// Format the sheet with grey background and no visible lines
        /// </summary>
        /// <param name="ws">The sheet to format</param>
        internal static void SetSheetFormatPattern(Worksheet ws)
        {
            Interior int_all = ws.Cells.Interior;
            int_all.Pattern = XlPattern.xlPatternSolid;
            int_all.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            int_all.ThemeColor = XlThemeColor.xlThemeColorDark1;
            int_all.TintAndShade = -0.349986266670736;
            int_all.PatternTintAndShade = 0;

            ws.Activate();
            RibbonHandler.ExcelApplication.ActiveWindow.DisplayGridlines = false;
        }
    }
}
