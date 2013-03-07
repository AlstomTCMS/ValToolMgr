using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using ExcelTools = Microsoft.Office.Tools.Excel;
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
            TabColorLightBlue(eps.Tab);

            eps.Range["B3:B10"].Value = new String[] { "Function", "Num_PR", "Indice_PR", "Date_PR", 
                                    "Ref_FRScc", "Ind_FRScc", "Versions MPU", "Aim of the function"};


            eps.Range["B2"].Value = "";//D4&" "&D5&" - "&D3;
        }

        private static void CreateEvolSheet(Workbook wb)
        {
            Worksheet es = wb.Sheets[2];
            es.Name = StringEnum.GetStringValue(SheetsNames.EVOLUTION);
            TabColorLightBlue(es.Tab);
            General.SetGreySheetPattern(es);

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
            bcs.Tab.ThemeColor = XlThemeColor.xlThemeColorDark2;
            bcs.Tab.TintAndShade = -9.99786370433668E-02;
        }

        private static void CreateSwVTPSheet(Workbook wb)
        {
            Worksheet SwvtpS = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            SwvtpS.Name = StringEnum.GetStringValue(SheetsNames.SW_VTP);
            TabColorLightBlue(SwvtpS.Tab);
            General.SetGreySheetPattern(SwvtpS);

            ListObject testsTable = SwvtpS.ListObjects.Add(XlListObjectSourceType.xlSrcRange, SwvtpS.Range["B2:F2"], XlYesNoGuess.xlYes);
            testsTable.Name = "TestsList_0";
            SwvtpS.Range["A2:F2"].Value = new String[] { StringEnum.GetStringValue(SwVTP_Columns.CATEGORY), 
                                                        StringEnum.GetStringValue(SwVTP_Columns.TEST), 
                                                        StringEnum.GetStringValue(SwVTP_Columns.BENCH_CONF), 
                                                        StringEnum.GetStringValue(SwVTP_Columns.REQUIREMENT),
                                                        StringEnum.GetStringValue(SwVTP_Columns.DESC),
                                                        StringEnum.GetStringValue(SwVTP_Columns.COMMENT) };

            testsTable.TableStyle = "TableStyleMedium2";
            //testsTable.Range.EntireColumn.AutoFit();

            General.UnformatGrey(SwvtpS.Range["A1", testsTable.Range.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing)]);
            
            // Tests title
            Range testsTitleRange = testsTable.Range.Offset[-1, 0].Rows[1];
            testsTitleRange.MergeCells = true;
            testsTitleRange.Value = "Tests";
            testsTitleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            testsTitleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            testsTitleRange.Font.Bold = true;
            testsTitleRange.EntireRow.RowHeight = 20;

            Border leftTestsBorder = testsTitleRange.Offset[0,-1].Cells[1,1].Borders[XlBordersIndex.xlEdgeRight];
            Border testsTableBorder = testsTable.TableStyle.TableStyleElements[XlTableStyleElementType.xlWholeTable].Borders[XlBordersIndex.xlEdgeLeft];
            //leftTestsBorder = testsTableBorder;
            leftTestsBorder.LineStyle = testsTableBorder.LineStyle;
            leftTestsBorder.Weight = testsTableBorder.Weight;
            leftTestsBorder.ThemeColor = testsTableBorder.ThemeColor;
            leftTestsBorder.TintAndShade = testsTableBorder.TintAndShade;

            //Category Style
            Range catRange = SwvtpS.Range["A2"];
            Interior cat_title_int = catRange.Interior;
            cat_title_int.Pattern = XlPattern.xlPatternSolid;
            cat_title_int.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            cat_title_int.ThemeColor = XlThemeColor.xlThemeColorAccent1;
            cat_title_int.TintAndShade = 0;
            cat_title_int.PatternTintAndShade = 0;
            Font font = catRange.Font;
            font.Bold = true;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;

            Range titleRow = SwvtpS.Rows["2:2"];
            titleRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRow.EntireRow.RowHeight = 35;
            //titleRow.WrapText = true;

            formatColumnsSwVTP();
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.TEST)].Range.EntireColumn.Hidden = true;

            SetBenchConfValidation(testsTable);


            // Add Tests's list event handler
            //ExcelTools.Worksheet SwvtpT = SwvtpS as ExcelTools.Worksheet;
            //ExcelTools.ListObject testsTableT = testsTable as ExcelTools.ListObject;
            //(testsTable as ExcelTools.ListObject).Change += new Microsoft.Office.Tools.Excel.ListObjectChangeHandler(list1_Change);
            //DebugFillingSwVTP(wb);
        }


        internal static void formatColumnsSwVTP()
        {
            Worksheet ws = RibbonHandler.ExcelApplication.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            ListObject testsTable = ws.ListObjects[1];
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.TEST)].Range.ColumnWidth = 4;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.BENCH_CONF)].Range.ColumnWidth = 6.57;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.REQUIREMENT)].Range.ColumnWidth = 16;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.DESC)].Range.ColumnWidth = 25;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.COMMENT)].Range.ColumnWidth = 25;
            testsTable.Range.WrapText = true;
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

        private static void TabColorLightBlue(Tab tab)
        {
            tab.ThemeColor = XlThemeColor.xlThemeColorLight2;
            tab.TintAndShade = 0.599993896298105;
        }

        private static void SetBenchConfValidation(ListObject testsTable)
        {
            ListColumn lc = testsTable.ListColumns[StringEnum.GetStringValue(SwVTP_Columns.BENCH_CONF)];
            //List<string> validConfs = new List<string>() { "A1", "A2", "B", "C", "D" };
            //string[] validConfs = new string[] { "A1", "A2", "B", "C", "D" };
            //string validConfs = "\"A1\",\"A2\",\"B\",\"C\",\"D\"";
            string validConfs = "A1;A2;B;C;D";

            Validation valid = lc.Range.Cells[2,1].Validation;
            valid.Delete();
            valid.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, 
                XlFormatConditionOperator.xlBetween, validConfs);
            valid.IgnoreBlank = true;
            valid.InCellDropdown = true;
            valid.InputTitle = "";
            valid.ErrorTitle = "";
            valid.InputMessage = "";
            valid.ErrorMessage = "";
            valid.ShowInput = true;
            valid.ShowError = true;
        }

        private static void list1_Change(Range targetRange, ExcelTools.ListRanges changedRanges)
        {
            string cellAddress = targetRange.get_Address(Excel.XlReferenceStyle.xlA1);

            switch (changedRanges)
            {
                case ExcelTools.ListRanges.DataBodyRange:
                    MessageBox.Show("The cells at range " + cellAddress +
                        " in the data body changed.");
                    break;
                case ExcelTools.ListRanges.HeaderRowRange:
                    MessageBox.Show("The cells at range " + cellAddress +
                        " in the header row changed.");
                    break;
                case ExcelTools.ListRanges.TotalsRowRange:
                    MessageBox.Show("The cells at range " + cellAddress +
                        " in the totals row changed.");
                    break;
                default:
                    MessageBox.Show("The cells at range " + cellAddress +
                        " changed.");
                    break;
            }
        }
    }
}
