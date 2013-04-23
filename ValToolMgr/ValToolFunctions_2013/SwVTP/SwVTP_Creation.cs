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
using System.Reflection;
using System.IO;

namespace ValToolFunctions_2013
{
    internal class SwVTP_Creation
    {
        static ExcelTools.ListObject testsTableT;
        static ExcelTools.ListObject evolTableT;

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
                    //http://msdn.microsoft.com/fr-fr/library/vstudio/3y21t6y4(v=vs.100).aspx
                    //http://lgmorand.developpez.com/dotnet/regex/
                    Regex rgx = new Regex(@"[a-zA-Z]\d_\d{3}_\d_[A-Za-z]\d");
                    if (rgx.IsMatch(filename))
                    {
                        PRname = rgx.Match(filename).ToString();
                        SaveExcelFile(PRname);
                    }
                    else
                    {
                        MessageBox.Show("Invalide file name format"); //throw 
                    }
                }
            }
        }

        private static void SaveExcelFile(string fileName)
        {
            Excel.Application app = RibbonHandler.ExcelApplication;
            Workbook wb = app.Workbooks.Add();

            //Add init sheets
            CreateEndpaperSheet(wb, fileName);
            CreateEvolSheet(wb);
            CreateBenchConfSheet(wb);
            CreateSwVTPSheet(wb); 

            app.DisplayAlerts = false;
            //Save file and show it
            wb.SaveAs(fileName);
            app.DisplayAlerts = true;
            wb.Saved = true;
            wb.AddToFavorites();
        }

        [System.Obsolete("Use createWholeTestFormat instead", true)]
        internal static void initThroughTemplate(string fileName)
        {
            Excel.Application app = RibbonHandler.ExcelApplication;


            //string tempPath = @"D:\Documents and Settings\e_dleona\Local Settings\Temp\";
            ////instal template vsto
            //File.WriteAllBytes(tempPath + "ExcelTemplate_2013.vsto", Properties.Resources.ExcelTemplate_2013_Vsto);
            //File.WriteAllBytes(tempPath + "ExcelTemplate_2013.dll.manifest", Properties.Resources.ExcelTemplate_2013_dll_manifest);
            //File.WriteAllBytes(tempPath + "ExcelTemplate_2013.dll", Properties.Resources.ExcelTemplate_2013_dll);

            //// Init a new workbook from our 2013 embeded template
            //// by copying it to a temporary file
            //string sPath = tempPath + "Template_2013.xltx"; 
            //File.WriteAllBytes(sPath, Properties.Resources.Template_2013);
            Workbook wb = app.Workbooks.Add("ExcelTemplate_2013");
            //File.Delete(sPath);

            initEvol(wb);
            initEndpaper(wb, fileName);

            //goto the function name in order to indiquate to the user that we want him to input it
            app.Goto("FunctionName");

            app.DisplayAlerts = false;
            //Save file and show it
            wb.SaveAs(fileName);
            app.DisplayAlerts = true;
            wb.Saved = true;
            wb.AddToFavorites();
        }
        [System.Obsolete("Use createWholeTestFormat instead")]
        internal static void initEndpaper(Workbook wb, string filename)
        {
            Worksheet eps = wb.Sheets[StringEnum.GetStringValue(SheetsNames.ENDPAPER_PR)];
            eps.Unprotect();

            Regex rgx = new Regex(@"^[a-zA-Z][0-9]_\d{3}_\d{1}");
            wb.Names.Item("Num_PR").RefersToRange.Value = rgx.Match(filename).ToString();
            rgx = new Regex(@"[a-zA-Z][0-9]{1,}$");
            wb.Names.Item("Indice_PR").RefersToRange.Value = rgx.Match(filename).ToString();
            eps.Range["D6"].Value = General.GetCurrentDate();

            eps.Protect(DrawingObjects: false, Contents: true, Scenarios: true);
        }
        [System.Obsolete("Use createWholeTestFormat instead")]
        internal static void initEvol(Workbook wb)
        {
            Worksheet es = wb.Sheets[StringEnum.GetStringValue(SheetsNames.EVOLUTION)];
            es.Range["evolListobject[Version]"].Value = "A0";
            es.Range["evolListobject[Date]"].Value = General.GetCurrentDate();
            es.Range["evolListobject[Name]"].Value = Environment.UserName;
            es.Range["evolListobject[Modification]"].Value = "Creation";
        }
        [System.Obsolete("Use createWholeTestFormat instead")]
        internal static void wbT_Open(object sender, EventArgs e)
        {
            //if (e.Item.isNew)
            //{
            Microsoft.Office.Tools.Excel.ListObject evolLIst = sender as Microsoft.Office.Tools.Excel.ListObject;
            evolLIst.Range.Offset[1, 0].EntireRow.Hidden = false;
            //}
            //System.Windows.Forms.MessageBox.Show("This data is read-only.");
        }

        /// <summary>
        /// Create the "Page de garde" of the book
        /// </summary>
        /// <param name="wb"></param>
        private static void CreateEndpaperSheet(Workbook wb, string filename)
        {
            Worksheet eps = wb.Sheets[1];
            eps.Name = StringEnum.GetStringValue(SheetsNames.ENDPAPER_PR); 
            TabColorLightBlue(eps.Tab);

            //Mask what is not on the area of the sheet
            //General.SetGreySheetPattern(eps);
            //General.UnformatGrey(eps.Range["A1", "O27"]);
            //ScrollArea 
            Range lastColumn = eps.Range["P1"].get_End(Excel.XlDirection.xlToRight);
            eps.Range["P1",lastColumn].EntireColumn.Hidden = true;
            eps.Rows["28:1048576"].EntireRow.Hidden = true; 
            eps.Activate();
            RibbonHandler.ExcelApplication.ActiveWindow.DisplayGridlines = false;


            eps.Range["B3:B20"].Value = RibbonHandler.ExcelApplication.WorksheetFunction.Transpose(
                                        new String[] { "Function", "Num_PR", "Indice_PR", "Date_PR", 
                                        "Ref_FRScc", "Ind_FRScc", "Versions              MPU", "DDU","TCU","ACU","BCU", 
                                        "ATESS", "LZB", "TRU", "Locomotive's type", "Locomotive's number", "Test's date", "Aim of the function"});


            eps.Columns["A:A"].ColumnWidth = 2.5;
            eps.Columns["B:B"].ColumnWidth = 10.5;
            eps.Columns["C:C"].ColumnWidth = 10.5;
            eps.Columns["D:D"].ColumnWidth = 2.5;
            eps.Columns["O:O"].ColumnWidth = 2.5;

            Range TotalEditZone = eps.Range["D3:N20"];

            //Title line
            SetTitles(eps.Range["B2:N2"]);

            //Function name (AF code)
            SetTitles(eps.Range["B3:C3"], false);
            SetEditZone(eps.Range["D3:N3"], ref TotalEditZone);

            //PR References zone
            SetTitles(eps.Range["B4:C8"], false);
            SetEditZone(eps.Range["D4:N8"], ref TotalEditZone);

            //Hardware versions zone
            SetTitles(eps.Range["B9:C19"], false);
            SetEditZone(eps.Range["D9:N19"], ref TotalEditZone);
            
            // Aim of the function
            Range functionGoal = eps.Range["B20:C20"];
            functionGoal.EntireRow.RowHeight = 40;
            SetTitles(functionGoal, false);
            SetEditZone(eps.Range["D20:N20"], ref TotalEditZone);


            // Approval titles
            Range swVTPWriter = eps.Range["B22:E22"];
            SetTitles(swVTPWriter);
            swVTPWriter.Value = "SwVTP's Writer";
            SetEditZone(eps.Range["B23:E26"], ref TotalEditZone, true);

            Range testWriter = eps.Range["F22:H22"];
            SetTitles(testWriter);
            testWriter.Value = "Tests's Writer";
            SetEditZone(eps.Range["F23:H26"], ref TotalEditZone, true);

            Range controller = eps.Range["I22:K22"];
            SetTitles(controller);
            controller.Value = "Controller";
            SetEditZone(eps.Range["I23:K26"], ref TotalEditZone, true);

            Range approver = eps.Range["L22:N22"];
            SetTitles(approver);
            approver.Value = "Approver";
            SetEditZone(eps.Range["L23:N26"], ref TotalEditZone, true);

            eps.Rows[22].EntireRow.RowHeight = 25;
            eps.Rows[26].EntireRow.RowHeight = 90;


            //Add Named Ranges
            wb.Names.Add("FunctionName", eps.Range["D3"], true);//RefersToR1C1: "=" + SheetsNames.ENDPAPER + "!R3C4");
            wb.Names.Add("Num_PR", eps.Range["D4"], true);//RefersToR1C1: "=" + SheetsNames.ENDPAPER + "!R4C4");
            wb.Names.Add("Indice_PR", eps.Range["D5"], true);//RefersToR1C1: "=" + SheetsNames.ENDPAPER + "!R5C4");
            eps.Range["B2:N2"].FormulaR1C1 = @"=Num_PR&"" ""&Indice_PR&"" - ""&FunctionName";

            Regex rgx = new Regex(@"^[a-zA-Z][0-9]_\d{3}_\d{1}");
            wb.Names.Item("Num_PR").RefersToRange.Value = rgx.Match(filename).ToString();
            rgx = new Regex(@"[a-zA-Z][0-9]{1,}$");
            wb.Names.Item("Indice_PR").RefersToRange.Value = rgx.Match(filename).ToString();
            eps.Range["D6"].Value = General.GetCurrentDate();

            //Template version
            Range endpaperVersionRange = eps.Range["B1"];
            wb.Names.Add("EndpaperVersion", endpaperVersionRange, true);
            endpaperVersionRange.Value = "v1.0.0";
            Font epVFont = endpaperVersionRange.Font;
            // Workaround of Excel 2010 formatting bug : http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/0fe66a4d-357a-4d74-b502-32848e7b44ba/
            //epVFont.ThemeColor = XlThemeColor.xlThemeColorDark1;
            //epVFont.TintAndShade = -4.99893185216834E-02;
            epVFont.Color = 15987699;

            // Protect Sheet Editing
            //Define editing zones
            //ActiveSheet.Protection.AllowEditRanges.Add Title:="Range123", Range:=Range("K4:L10")
            eps.Protection.AllowEditRanges.Add("EditZone", TotalEditZone, Type.Missing);
            eps.Protect(DrawingObjects: false, Contents: true, Scenarios: true);
        }

        private static void SetEditZone(Range range, ref Range totalEditZone, bool isApproval = false)
        {
            range.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlLineStyleNone;

            foreach (XlBordersIndex edge in new XlBordersIndex[] { XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight })
            {
                Border border = range.Borders[edge];
                border.LineStyle = XlLineStyle.xlContinuous;
                border.Weight = XlBorderWeight.xlMedium;
                border.ColorIndex = 0;
                border.TintAndShade = 0;
            }
            Border borderIH = range.Borders[XlBordersIndex.xlInsideHorizontal];
            borderIH.LineStyle = XlLineStyle.xlContinuous;
            borderIH.Weight = XlBorderWeight.xlThin;
            borderIH.ColorIndex = 0;
            borderIH.TintAndShade = 0;

            range.VerticalAlignment = XlVAlign.xlVAlignTop;
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            range.WrapText = false; ;
            range.Orientation = 0;
            range.AddIndent = false;
            range.IndentLevel = 0;
            range.ShrinkToFit = false;
            range.ReadingOrder = (int)Excel.Constants.xlContext;
            range.MergeCells = false;

            if (isApproval)
            {
                range.Columns[1].value = RibbonHandler.ExcelApplication.WorksheetFunction.Transpose(
                                        new String[] { "Name : ", "Entity : ", "Date :", "Stamp :" });
                if (totalEditZone != null)
                {
                    Range localEditZone = range.Columns[2];
                    for (int i = 3; i <= range.Columns.Count; i++)
                    {
                        localEditZone = RibbonHandler.ExcelApplication.Union(localEditZone, range.Columns[i]);
                    }
                    localEditZone.Merge(true);
                    totalEditZone = RibbonHandler.ExcelApplication.Union(localEditZone, totalEditZone);
                }
            }
            else
            {
                range.Merge(true);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="isZoneTitle"></param>
        private static void SetTitles(Range range, bool isZoneTitle = true)
        {
            range.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlLineStyleNone;

            foreach (XlBordersIndex edge in new XlBordersIndex[] { XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight })
            {
                Border border = range.Borders[edge];
                border.LineStyle = XlLineStyle.xlContinuous;
                border.Weight = XlBorderWeight.xlMedium;
                border.ColorIndex = 0;
                border.TintAndShade = 0;
            }

            Interior inte = range.Interior;
            inte.Pattern = XlPattern.xlPatternSolid;
            inte.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            inte.ThemeColor = XlThemeColor.xlThemeColorAccent1;
            inte.TintAndShade = 0; ;
            inte.PatternTintAndShade = 0;

            Font font = range.Font;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;
            font.Bold = true;
            font.Size = 12;

            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            range.WrapText = false; ;
            range.Orientation = 0;
            range.AddIndent = false;
            range.IndentLevel = 0;
            range.ShrinkToFit = false;
            range.ReadingOrder = (int)Excel.Constants.xlContext;
            range.MergeCells = false;

            if (isZoneTitle)
            {
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                range.Merge();
            }
            else
            {
                range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            }
        }

        private static void CreateEvolSheet(Workbook wb)
        {
            Worksheet es = wb.Sheets[2];
            es.Name = StringEnum.GetStringValue(SheetsNames.EVOLUTION);
            TabColorLightBlue(es.Tab);
            General.SetGreySheetPattern(es);

            Range lastColumn = es.Range["E1"].get_End(Excel.XlDirection.xlToRight);
            es.Range["E1", lastColumn].EntireColumn.Hidden = true;
            //es.Rows["4:1048576"].EntireRow.Hidden = true;
            es.Activate();
            RibbonHandler.ExcelApplication.ActiveWindow.DisplayGridlines = false;

            //ListObject evolTable = es.ListObjects.Add(XlListObjectSourceType.xlSrcRange, es.Range["A1:D1"], XlYesNoGuess.xlYes);

            System.Data.DataSet ds = new System.Data.DataSet();
            System.Data.DataTable dt = ds.Tables.Add("Evolutions");
            dt.Columns.Add(new System.Data.DataColumn("Version"));
            dt.Columns.Add(new System.Data.DataColumn("Date"));
            dt.Columns.Add(new System.Data.DataColumn("Name"));
            dt.Columns.Add(new System.Data.DataColumn("Modification"));

            // Add a new row to the DataTable.
            System.Data.DataRow dr = dt.NewRow();
            dr["Version"] = "A0";
            dr["Date"] = General.GetCurrentDate();
            dr["Name"] = Environment.UserName;
            dr["Modification"] = "Creation";
            dt.Rows.Add(dr);

            // Create a list object.
            evolTableT = RibbonHandler.Factory.GetVstoObject(es).Controls.AddListObject(es.Range["A1"], "evolListobject");

            // Bind the list object to the DataTable.
            evolTableT.AutoSetDataBoundColumnHeaders = true;
            //evolTable.DataSource = ds;
            //evolTable.DataMember = "Evolutions";
            evolTableT.SetDataBinding(ds, "Evolutions", "Version", "Date", "Name", "Modification");

            // Create the event handler.
            //evolTable.DataMemberChanged += new System.EventHandler(EvolList_Change);

            //es.Range["A1:D1"].Value = new String[] { "Version", "Date (M/D/Y)", "Name", "Modification" };
            //evolTable.Name = "evolTable";
            evolTableT.TableStyle = "TableStyleMedium2";

            Interior int_evol = evolTableT.Range.Interior;
            int_evol.Pattern = XlPattern.xlPatternNone;
            int_evol.TintAndShade = 0;
            int_evol.PatternTintAndShade = 0;

            evolTableT.Range.EntireColumn.AutoFit();
            evolTableT.ListColumns[3].Range.ColumnWidth = 15; //Name
            evolTableT.ListColumns[4].Range.ColumnWidth = 60; //Modif
            es.Range["evolListobject[Date]"].NumberFormat = "mm/dd/yyyy";

            evolTableT.Change += new ExcelTools.ListObjectChangeHandler(EvolList_Change);
        }

        //internal static void list1_BeforeAddDataBoundRow(object sender, Microsoft.Office.Tools.Excel.BeforeAddDataBoundRowEventArgs e)
        //{
        //    //if (e.Item.isNew)
        //    //{
        //        Microsoft.Office.Tools.Excel.ListObject evolLIst = sender as Microsoft.Office.Tools.Excel.ListObject;
        //        evolLIst.Range.Offset[1, 0].EntireRow.Hidden = false;
        //    //}
        //    //System.Windows.Forms.MessageBox.Show("This data is read-only.");
        //}

        //internal static void EvolList_Change(object sender, EventArgs e)
        //{
        //    Microsoft.Office.Tools.Excel.ListObject evolLIst = sender as Microsoft.Office.Tools.Excel.ListObject;
        //    evolLIst.Range.Offset[1, 0].EntireRow.Hidden = false;
        //}

        internal static void EvolList_Change(Range targetRange, ExcelTools.ListRanges changedRanges)
        {
            if (changedRanges == ExcelTools.ListRanges.DataBodyRange)
            {
                if (evolTableT == null)
                {
                    evolTableT = RibbonHandler.Factory.GetVstoObject(RibbonHandler.ExcelApplication.Sheets[
                        StringEnum.GetStringValue(SheetsNames.EVOLUTION)]);
                }
                if (evolTableT != null)
                {
                    //// Hidden the entire sheet
                    //Worksheet evolWS = (evolTableT.Parent as Worksheet);
                    //evolWS.Cells.EntireRow.Hidden = true;
                    //// Display table zone + Next line
                    //RibbonHandler.ExcelApplication.Union(evolTableT.Range, evolTableT.Range.Offset[1, 0]).EntireRow.Hidden = false;

                    //Si on modifie la dernière ligne
                    if (evolTableT.ListRows.Count +1 == targetRange.Row)
                    {
                        evolTableT.Change -= new ExcelTools.ListObjectChangeHandler(EvolList_Change);
                        evolTableT.DataBodyRange.Cells[evolTableT.ListRows.Count, 2].Value = General.GetCurrentDate();
                        evolTableT.DataBodyRange.Cells[evolTableT.ListRows.Count, 3].Value = Environment.UserName;
                        evolTableT.Change += new ExcelTools.ListObjectChangeHandler(EvolList_Change);
                    }
                    //targetRange.Select();
                }
            }
        }

        private static void CreateBenchConfSheet(Workbook wb)
        {
            //Worksheet bcs = wb.Sheets[3];
            //bcs.Name = StringEnum.GetStringValue(SheetsNames.BENCH_CONF);
            //bcs.Tab.ThemeColor = XlThemeColor.xlThemeColorDark2;
            //bcs.Tab.TintAndShade = -9.99786370433668E-02;
        }

        private static void CreateSwVTPSheet(Workbook wb)
        {
            Worksheet SwvtpS = wb.Sheets[3];//wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            SwvtpS.Name = StringEnum.GetStringValue(SheetsNames.SW_VTP);
            TabColorLightBlue(SwvtpS.Tab);
            General.SetGreySheetPattern(SwvtpS);

            ListObject testsTable = SwvtpS.ListObjects.Add(XlListObjectSourceType.xlSrcRange, SwvtpS.Range["B2:L2"], XlYesNoGuess.xlYes);
            testsTable.Name = "TestsList_0";
            SwvtpS.Range["A2:L2"].Value = new String[] { StringEnum.GetStringValue(SwVTx_Columns.CATEGORY), 
                                                        StringEnum.GetStringValue(SwVTx_Columns.TEST), 
                                                        StringEnum.GetStringValue(SwVTx_Columns.BENCH_CONF), 
                                                        StringEnum.GetStringValue(SwVTx_Columns.REQUIREMENT),
                                                        StringEnum.GetStringValue(SwVTx_Columns.STEP), 
                                                        StringEnum.GetStringValue(SwVTx_Columns.TITLE),
                                                        StringEnum.GetStringValue(SwVTx_Columns.ACTION),
                                                        StringEnum.GetStringValue(SwVTx_Columns.CHECK), 
                                                        StringEnum.GetStringValue(SwVTx_Columns.STATUS),
                                                        StringEnum.GetStringValue(SwVTx_Columns.JUSTIFICATION),
                                                        StringEnum.GetStringValue(SwVTx_Columns.R_COMMENT) };

            testsTable.TableStyle = "TableStyleMedium2";
            //testsTable.Range.EntireColumn.AutoFit();

            General.UnformatGrey(SwvtpS.Range["A1", testsTable.Range.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing)]);
            
            // Tests title
            Range testsTitleRange = testsTable.Range.Offset[-1, 0].Rows[1];
            testsTitleRange.MergeCells = true;
            testsTitleRange.Value = "TEST";
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
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.TEST)].Range.EntireColumn.Hidden = true;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.STEP)].Range.EntireColumn.Hidden = true;
            SwvtpS.Columns["H:L"].EntireColumn.Hidden = true;

            SetBenchConfValidation(testsTable);


            // Add Tests's list event handler
            testsTableT = RibbonHandler.Factory.GetVstoObject(testsTable);
            testsTableT.Change += new ExcelTools.ListObjectChangeHandler(TestsList_Change);

            Range lastTableColumn = testsTable.Range.Columns[testsTable.Range.Columns.Count].Offset[0, 1];
            Range lastColumn = lastTableColumn.get_End(Excel.XlDirection.xlToRight);
            SwvtpS.Range[lastTableColumn, lastColumn].EntireColumn.Hidden = true;

            // SwVTP Format sheet version
            Range SwVTPFormatVersion = SwvtpS.Range["A1"];
            wb.Names.Add("SwVTPFormatVersion", SwVTPFormatVersion, true);
            SwVTPFormatVersion.Value = "v1.0.0";
            Font epVFont = SwVTPFormatVersion.Font;
            epVFont.Color = 15987699;

            //DebugFillingSwVTP(wb);
        }


        internal static void formatColumnsSwVTP()
        {
            Worksheet ws = RibbonHandler.ExcelApplication.Sheets[StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            ListObject testsTable = ws.ListObjects[1];
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.TEST)].Range.ColumnWidth = 4;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.BENCH_CONF)].Range.ColumnWidth = 6.57;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.REQUIREMENT)].Range.ColumnWidth = 16;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.TITLE)].Range.ColumnWidth = 25;
            testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.ACTION)].Range.ColumnWidth = 25;
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
            ListColumn lc = testsTable.ListColumns[StringEnum.GetStringValue(SwVTx_Columns.BENCH_CONF)];
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

        internal static void TestsList_Change(Range targetRange, ExcelTools.ListRanges changedRanges)
        {
            Worksheet swVTP = RibbonHandler.ExcelApplication.Sheets[
                        StringEnum.GetStringValue(SheetsNames.SW_VTP)];
            int lastTestsRow = swVTP.ListObjects[swVTP.ListObjects.Count].Range.get_End(Excel.XlDirection.xlDown).Row;
            General.FormatGrey(swVTP.Range["A3:A65365"]);
            General.UnformatGrey(swVTP.Range["A3:A"+lastTestsRow]);

            //string cellAddress = targetRange.get_Address(Excel.XlReferenceStyle.xlA1);
            ////testsTableT.
            
            //switch (changedRanges)
            //{
            //    case ExcelTools.ListRanges.DataBodyRange:
            //        MessageBox.Show("The cells at range " + cellAddress +
            //            " in the data body changed.");
            //        break;
            //    case ExcelTools.ListRanges.HeaderRowRange:
            //        MessageBox.Show("The cells at range " + cellAddress +
            //            " in the header row changed.");
            //        break;
            //    case ExcelTools.ListRanges.TotalsRowRange:
            //        MessageBox.Show("The cells at range " + cellAddress +
            //            " in the totals row changed.");
            //        break;
            //    default:
            //        MessageBox.Show("The cells at range " + cellAddress +
            //            " changed.");
            //        break;
            //}
        }
    }
}
