﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using ValToolFunctionsStub;

namespace ValToolFunctions_2013
{
    /// <summary>
    /// This class create a test from scratch
    /// </summary>
    internal class CreateTest
    {
        [System.Obsolete("Use createWholeTestFormat instead", true)]
        internal static void NewPR()
        {
            try
            {
                //Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[1, 1];
                //string cellValue = range.Value.ToString();

                //demander à l//utilisateur le nom qu//il veut mettre
                ////filesavefullname = application.getsaveasfilename(initialfilename:="b2_xxx_y_a0", _
                ////filefilter:="xls files (*.xls), *.xls")
                string defaultvalue = "1.";
                string testname = defaultvalue;
                if (Dialog.InputBox("test name", "please, give a name to your test.", ref testname) == DialogResult.OK)
                {
                    //créer l//ensemble des éléments du format
                    if ((testname != "") && (testname != defaultvalue))
                    {
                        //todo: tester si le test existe deja...
                        createWholeTestFormat(testname);
                    }
                }
                ////Sauvegarder    
            }
            catch { }
        }

        /// <summary>
        /// Create a whole PR test sheet in 2013 format
        /// </summary>
        /// <param name="testName"></param>
        internal static void createWholeTestFormat(string sheetName)
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
            if (! General.HasActiveBook(false)) {
                RibbonHandler.ExcelApplication.Workbooks.Add();
            }
    

            Worksheet testSheet = General.InitSheet(sheetName);
            testSheet.Activate();
            testSheet.Tab.ThemeColor = XlThemeColor.xlThemeColorLight2;
            testSheet.Tab.TintAndShade = 0;
            //General.SetGreySheetPattern(testSheet);
            testSheet.Cells.ColumnWidth = 25;

            //try{
                AddTableDescription(testSheet);
            //}
            //catch { }
            AddTableAction(testSheet);
            AddTableCheck(testSheet);
            AddActionLabel(testSheet);
            AddCheckLabel(testSheet);
            FormatTestSheet(testSheet);
            AddTestTitle(testSheet);
        }

        private static void FormatTestSheet(Worksheet testSheet)
        {
            testSheet.Columns["C:C"].ColumnWidth = 25;
            testSheet.Rows["1:2"].Group();
            //testSheet.Range[TEST.TABLE.PREFIX.ACTION + General.getTestNumber(testSheet.Name) + "["+ TEST.STEP_PATERN + "]"].Select();
            testSheet.Range["D5"].Select();
            testSheet.Application.ActiveWindow.FreezePanes = true;
            testSheet.Range["A1"].Select();

            //Range bottomRightCorner = testSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);//testSheet.Range[TEST.TABLE.PREFIX.CHECK + General.getTestNumber(testSheet.Name)].End(XlDirection.xlDown);
            //General.UnformatGrey(testSheet.Range["A1", bottomRightCorner.Offset[1,1]]);
        }

        ////Ajoute la table de description en haut
        private static void AddTableDescription(Worksheet testSheet)
        {
    
            //on insert une ligne supplémentaire pour les titres (qu'il n'y a pas)
            testSheet.Rows["1:1"].Insert(XlDirection.xlDown);
            string tableName = TEST.TABLE.PREFIX.DESC + General.getTestNumber(testSheet.Name);
            testSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, testSheet.Range["$C$1:$D$4"], XlYesNoGuess.xlYes).Name = tableName;

            //try
            //{
            //    AddDescTableFormat();
            //}
            //catch { }
            ListObject descTable = testSheet.ListObjects[tableName];
            descTable.TableStyle = TEST.DESCRIPTION_TABLE_STYLE + " " + TEST.DESCRIPTION_TABLE_STYLE_VERSION;
            descTable.ShowHeaders = false;
            descTable.ShowTableStyleFirstColumn = true;
            descTable.ShowTableStyleColumnStripes = true;
            descTable.Range.VerticalAlignment = XlVAlign.xlVAlignTop;
            descTable.Range.WrapText = true;

            //On réefface cette ligne qui ne sert plus
            testSheet.Rows["1:1"].Delete (XlDirection.xlUp);
        
            //Ajoute les labels des titres verticaux
            testSheet.Range["C1"].Value = StringEnum.GetStringValue(TEST.TABLE.TYPE.ACTION);
            testSheet.Range["C2"].Value = StringEnum.GetStringValue(TEST.TABLE.TYPE.CHECK);
            testSheet.Range["C3"].Value = "Name";
        
            // Efface la mise en forme de la première case de la ligne des totaux
            Interior totalInterior = testSheet.Range["C4"].Interior;
                totalInterior.Pattern = XlPattern.xlPatternSolid;
                totalInterior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                totalInterior.ThemeColor = XlThemeColor.xlThemeColorDark1;
                totalInterior.TintAndShade = 0;
                totalInterior.PatternTintAndShade = 0;

            Range col1 = descTable.DataBodyRange.Columns[1];//testSheet.Range[tableName + "[[#All],[Colonne1]]"];
            col1.HorizontalAlignment = XlHAlign.xlHAlignRight;
            col1.VerticalAlignment = XlHAlign.xlHAlignCenter;
            Font font = col1.Font;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;
            font.Bold = true;
        }

        private static void AddCheckLabel(Worksheet testSheet)
        {
            DefineVerticalLabel(testSheet, StringEnum.GetStringValue(TEST.TABLE.TYPE.CHECK));
        }

        private static void AddActionLabel(Worksheet testSheet)
        {
            DefineVerticalLabel(testSheet, StringEnum.GetStringValue(TEST.TABLE.TYPE.ACTION));
        }

        private static void DefineVerticalLabel(Worksheet testSheet, String label)
        {
            
            testSheet.Columns["A:A"].ColumnWidth = 5.5;
        
            string tableAddress = testSheet.ListObjects[TEST.TABLE_PREFIX + label + "_" + General.getTestNumber(testSheet.Name)].Range.Address;
            tableAddress = "A" + tableAddress.Substring(3, 2) + "A" + tableAddress.Substring(8, 1);            

            Range LabelRange = testSheet.Range[tableAddress];                
                LabelRange.MergeCells = true;
                LabelRange.Value = label;
                LabelRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                LabelRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                LabelRange.WrapText = false;
                LabelRange.Orientation = 90;
                LabelRange.AddIndent = false;
                LabelRange.IndentLevel = 0;
                LabelRange.ShrinkToFit = false;
                LabelRange.ReadingOrder = (int)Excel.Constants.xlContext;
            
            Font font =LabelRange.Font;
                font.Name = "Calibri";
                font.Size = 14;
                font.Strikethrough = false;
                font.Superscript = false;
                font.Subscript = false;
                font.OutlineFont = false;
                font.Shadow = false;
                font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                font.ThemeColor = XlThemeColor.xlThemeColorLight1;
                font.TintAndShade = 0;
                font.ThemeFont = XlThemeFont.xlThemeFontMinor;
                font.Bold = true;
        }

        internal static void UpdateVerticalLabel(Worksheet testSheet, ListObject table, Boolean addMode)
        {
            string tableAddress = null;

            if (table != null)
            {
                tableAddress = table.Range.Address;
            }

            if (tableAddress != null)
            {
                string[] tArray = Regex.Split(tableAddress, @"\$");

                try
                {
                    //Unmerge current merged zone
                    tableAddress = "A" + tArray[2] + "A" + (int.Parse(tArray[4]) + (addMode ? -1 : 1));
                    if (testSheet.Range[tableAddress].MergeCells)
                    {
                        testSheet.Range[tableAddress].UnMerge();
                    }

                    //merge the new one
                    tableAddress = "A" + tArray[2] + "A" + tArray[4];
                    testSheet.Range[tableAddress].Merge();
                }
                catch { }
            }
        }

        private static void AddTestTitle(Worksheet testSheet)
        {
            Range titleRange = testSheet.Range["B3"];
            titleRange.Value = Regex.Replace(testSheet.Name , "_", " ") ;
                //TODO: Donner un nom
            Font font = titleRange.Font;
            font.Name = "Calibri";
            font.Size = 14;
            font.Bold = true;
            font.Strikethrough = false;
            font.Superscript = false;
            font.Subscript = false;
            font.OutlineFont = false;
            font.Shadow = false;
            font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;
            font.ThemeFont = XlThemeFont.xlThemeFontMinor;

            Interior interior = titleRange.Interior;
            interior.Pattern = XlPattern.xlPatternSolid;
            interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
            interior.ThemeColor = XlThemeColor.xlThemeColorLight1;
            interior.TintAndShade = 0;
            interior.PatternTintAndShade = 0;
        
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            titleRange.WrapText = false;
            titleRange.Orientation = 0;
            titleRange.AddIndent = false;
            titleRange.IndentLevel = 0;
            titleRange.ShrinkToFit = false;
            titleRange.ReadingOrder = (int)Excel.Constants.xlContext;
            titleRange.MergeCells = false;

            testSheet.Columns["B:B"].ColumnWidth = 25;
            testSheet.Rows["3:3"].RowHeight = 30;
        }

        private static void AddTableCheck(Worksheet testSheet)
        {
            string tableName = TEST.TABLE.PREFIX.CHECK + General.getTestNumber(testSheet.Name);

            testSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, testSheet.Range["$B$8"], XlYesNoGuess.xlYes).Name = tableName;
                
            ListObject checkTable = testSheet.ListObjects[tableName];
            checkTable.TableStyle = "TableStyleMedium12";
            testSheet.Range["B8:D8"].Value = new string[]{"Target", "Location", TEST.STEP_PATERN};
            Range stepRange = checkTable.HeaderRowRange[3]; // testSheet.Range[tableName + "[[#Headers],[" + TEST.STEP_PATERN + "]]"];
            stepRange.AddIndent = true;
            stepRange.IndentLevel = 1;
            checkTable.ShowHeaders = false;        
        
            //Coloration de la colonne des variables
            Interior totalInterior = testSheet.Range["B9"].Interior;
                totalInterior.Pattern = XlPattern.xlPatternSolid;
                totalInterior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                totalInterior.ThemeColor = XlThemeColor.xlThemeColorAccent4;
                totalInterior.TintAndShade = 0;
                totalInterior.PatternTintAndShade = 0;
            Font font =testSheet.Range["B9"].Font;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;
            font.Bold = true;
        }

        private static void AddTableAction(Worksheet testSheet)
        {
            string tableName = TEST.TABLE.PREFIX.ACTION + General.getTestNumber(testSheet.Name);

            testSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, testSheet.Range["$B$5"], XlYesNoGuess.xlYes).Name = tableName;
                
            ListObject actionTable = testSheet.ListObjects[tableName];
            actionTable.TableStyle = "TableStyleMedium9";
            testSheet.Range["B5:D5"].Value = new string[]{"Target", "Location", TEST.STEP_PATERN};
            actionTable.ShowTotals = true;


            Range targetRange = actionTable.TotalsRowRange[1]; //testSheet.Range[tableName + "[[#Totals],[Target]]"];
            targetRange.FormulaR1C1 = "DELAY";
            targetRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            targetRange.VerticalAlignment = XlVAlign.xlVAlignBottom;
            targetRange.WrapText = false;
            targetRange.Orientation = 0;
            targetRange.AddIndent = false;
            targetRange.IndentLevel = 0;
            targetRange.ShrinkToFit = false;
            targetRange.ReadingOrder = (int)Excel.Constants.xlContext;
            targetRange.MergeCells = false;
        
            testSheet.Range["D7"].Value = "";
            Range stepRange = actionTable.HeaderRowRange[3]; // testSheet.Range[tableName + "[[#Headers],[" + TEST.STEP_PATERN + "]]"];
            stepRange.AddIndent = true;
            stepRange.IndentLevel = 1;
        
        
            //Coloration de la colonne des variables
            Interior totalInterior = testSheet.Range["B6"].Interior;
                totalInterior.Pattern = XlPattern.xlPatternSolid;
                totalInterior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
                totalInterior.ThemeColor = XlThemeColor.xlThemeColorAccent1;
                totalInterior.TintAndShade = 0;
                totalInterior.PatternTintAndShade = 0;
            Font font =testSheet.Range["B6"].Font;
            font.ThemeColor = XlThemeColor.xlThemeColorDark1;
            font.TintAndShade = 0;
            font.Bold = true;
        }


        /// <summary>
        /// Ajoute au workbook le style de tableau pour la partie descriptive s'il n'existe pas déjà
        /// </summary>
        /// <returns>vrai s'il faut mettre à jour le style du tableau</returns>
        internal static Boolean AddDescTableFormat()
        {
            Boolean updateTable=false;
            Boolean addTable=true;
            string[] tableName;
            TableStyle ts=null;
            Boolean addDescTableFormat=false;

            Workbook wb =RibbonHandler.ExcelApplication.ActiveWorkbook;

            // Vérifie que le style existe déjà et à quelle version
            foreach (TableStyle style in wb.TableStyles)
            {
                if(Regex.IsMatch(style.Name, TEST.DESCRIPTION_TABLE_STYLE + "*"))
                {
                    addTable = false;
                    tableName = Regex.Split(style.Name, " "); 
                    if (tableName.Length == 3)
                    {// Si on a la version
                        if (String.Compare(tableName[tableName.Length-1], TEST.DESCRIPTION_TABLE_STYLE_VERSION) > 0)
                        {
                            updateTable = true;
                        }
                    }
                    else if (tableName.Length == 2) {// Si on a la version
                        //If tableName(2) < PR_TEST_DESCRIPTION_TABLE_STYLE_VERSION Then
                        updateTable = true;
                        //End If
                    }else{
                        updateTable = true;
                    }
                }
            }
        
        
            if (addTable){
                ts = wb.TableStyles.Add(TEST.DESCRIPTION_TABLE_STYLE + " " + TEST.DESCRIPTION_TABLE_STYLE_VERSION);
            }

            if (addTable | updateTable)
            {
                if (ts == null)
                {
                    ts = wb.TableStyles[TEST.DESCRIPTION_TABLE_STYLE + " " + TEST.DESCRIPTION_TABLE_STYLE_VERSION];
                }
                ts.ShowAsAvailablePivotTableStyle = false;
                ts.ShowAsAvailableTableStyle = true;
                ts.ShowAsAvailableSlicerStyle = false;

                if (updateTable)
                {
                    addDescTableFormat = true;
                    //efface les styles avant de les définir
                    foreach (TableStyleElement stylesElement in ts.TableStyleElements)
                    {
                        stylesElement.Clear();
                    }
                }

                // -------------------------------------------------------------
                // LA Première colonne (les titres)
                // -------------------------------------------------------------
                TableStyleElement firstCol = ts.TableStyleElements[XlTableStyleElementType.xlFirstColumn];
                Font font = firstCol.Font;
                font.ThemeColor = XlThemeColor.xlThemeColorDark1;
                font.TintAndShade = 0;
                font.Bold = true;
                Interior inte = firstCol.Interior; inte.Color = 12419407; inte.TintAndShade = 0;
                foreach (XlBordersIndex edge in new XlBordersIndex[] { XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlInsideHorizontal })
                {
                    Border border = firstCol.Borders[edge];
                    border.LineStyle = XlLineStyle.xlLineStyleNone;
                    border.Weight = XlBorderWeight.xlThick;
                    border.ThemeColor = XlThemeColor.xlThemeColorDark1;
                    border.TintAndShade = 0;
                }

                // -------------------------------------------------------------
                // Colonnes impaires
                // -------------------------------------------------------------
                Interior oddsColInterior = ts.TableStyleElements[XlTableStyleElementType.xlColumnStripe1].Interior;
                oddsColInterior.Pattern = XlPattern.xlPatternSolid;
                oddsColInterior.PatternColorIndex = 0;
                oddsColInterior.Color = 15853276;
                oddsColInterior.TintAndShade = 0;
                oddsColInterior.PatternTintAndShade = 0;

                // -------------------------------------------------------------
                // Lignes impaires
                // -------------------------------------------------------------
                Border oddsLinesBorder = ts.TableStyleElements[XlTableStyleElementType.xlRowStripe1].Borders[XlBordersIndex.xlEdgeTop];
                oddsLinesBorder.LineStyle = XlLineStyle.xlLineStyleNone;
                oddsLinesBorder.Weight = XlBorderWeight.xlThin;
                oddsLinesBorder.ThemeColor = XlThemeColor.xlThemeColorLight2;
                oddsLinesBorder.TintAndShade = 0.799981688894314; //0.599963377788629

                // -------------------------------------------------------------
                // Lignes paires
                // -------------------------------------------------------------
                Border evenLinesBorder = ts.TableStyleElements[XlTableStyleElementType.xlRowStripe2].Borders[XlBordersIndex.xlEdgeTop];
                evenLinesBorder.LineStyle = XlLineStyle.xlLineStyleNone;
                evenLinesBorder.Weight = XlBorderWeight.xlThin;
                evenLinesBorder.ThemeColor = XlThemeColor.xlThemeColorLight2;
                evenLinesBorder.TintAndShade = 0.799981688894314; //0.599963377788629

                // -------------------------------------------------------------
                // Ligne des Totaux
                // -------------------------------------------------------------
                TableStyleElement totalRow = ts.TableStyleElements[XlTableStyleElementType.xlTotalRow];
                Font tf = totalRow.Font; tf.TintAndShade = 0; tf.ThemeColor = XlThemeColor.xlThemeColorDark1;
                Border totalBorder = totalRow.Borders[XlBordersIndex.xlEdgeTop];
                totalBorder.LineStyle = XlLineStyle.xlLineStyleNone;
                totalBorder.Weight = XlBorderWeight.xlThin;
                totalBorder.ThemeColor = XlThemeColor.xlThemeColorLight2;
                totalBorder.TintAndShade = 0.799951170384838; //0.799981688894314
            }
            return addDescTableFormat;
        }
    }
}