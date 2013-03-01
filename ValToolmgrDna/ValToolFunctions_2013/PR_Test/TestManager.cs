using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ValToolFunctions_2013
{
    public static class TestManager
    {
        public static void AddNewStep(Excel.Application xlsApp)
        {
                ExcelApplication.setInstance(xlsApp);
                if (General.isActivesheet_a_PR_Test())
                {
                    Worksheet ws = ExcelApplication.getInstance().ActiveSheet;
                    string testNumber = General.getTestNumber();
                    // Ajouter une colonne à chaque tableau
                    ListColumns actionT = ws.ListObjects[TEST.TABLE.PREFIX.ACTION + testNumber].ListColumns;
                    ListColumns checkT = ws.ListObjects[TEST.TABLE.PREFIX.CHECK + testNumber].ListColumns;
                    ListColumns descT = ws.ListObjects[TEST.TABLE.PREFIX.DESC + testNumber].ListColumns;
            
            
                    // Si tous les tableaux ont la meme taille
                    if ( actionT.Count == checkT.Count && actionT.Count == descT.Count + 1){
                        actionT.Add();
                        checkT.Add();
                        descT.Add();
                    }else if ( false){
                        int stepNumber = actionT.Count;
                        if ( stepNumber == checkT.Count){
                            checkT.Add();
                        }else{
                            //checkT.Resize Range("$B$15:$U$15")
                        }
                
                        if ( stepNumber == descT.Count + 1){
                            descT.Add();
                        }else{
                            //descT.Resize Range("$B$15:$U$15")
                        }
                    }else{
                        MessageBox.Show("All tables are not at the same size");
                    }
                }
        }

        public static void RemoveStep(Excel.Application xlsApp, EditingZone editingMode)
        {

                ExcelApplication.setInstance(xlsApp);
                if (General.isActivesheet_a_PR_Test())
                {
                    Worksheet ws = ExcelApplication.getInstance().ActiveSheet;
                    string testNumber = General.getTestNumber();
                    // Ajouter une colonne à chaque tableau
                    ListColumns descT = ws.ListObjects[TEST.TABLE.PREFIX.DESC + testNumber].ListColumns;
                    ListColumns actionT = ws.ListObjects[TEST.TABLE.PREFIX.ACTION + testNumber].ListColumns;
                    ListColumns checkT = ws.ListObjects[TEST.TABLE.PREFIX.CHECK + testNumber].ListColumns;


                    // If all tables have the same size
                    if (actionT.Count == checkT.Count && actionT.Count == descT.Count + 1)
                    {
                        //Delete until first step
                        if (actionT.Count > 3)
                        {
                            descT[descT.Count].Delete();
                            actionT[actionT.Count].Delete();
                            checkT[checkT.Count].Delete();
                        }
                        else if (actionT.Count == 3)
                        {
                            descT[2].DataBodyRange.ClearContents();
                            actionT[3].DataBodyRange.ClearContents();
                            actionT[3].Total.ClearContents();
                            checkT[3].DataBodyRange.ClearContents();
                        }
                    }
                }
        }

        public static void AddVariable(Excel.Application xlsApp, TEST.TABLE.TYPE type ,EditingZone editingMode)
        {
                ExcelApplication.setInstance(xlsApp);
                if (General.isActivesheet_a_PR_Test())
                {
                    Worksheet ws = ExcelApplication.getInstance().ActiveSheet;
                    string testNumber = General.getTestNumber();
                    ListObject checkT = ws.ListObjects[TEST.TABLE.PREFIX.CHECK + testNumber];

                    switch (type)
                    {
                        case TEST.TABLE.TYPE.ACTION:
                            ListObject actionT = ws.ListObjects[TEST.TABLE.PREFIX.ACTION + testNumber];
                            try
                            {
                                actionT.TotalsRowRange.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            }
                            catch
                            {
                                actionT.ListRows.Add(1);
                                Range checkTitle = checkT.Range.Offset[-1, -1].Cells[1, 1];
                                checkTitle.Cut(checkTitle.Offset[2, 0]);
                                CreateTest.UpdateVerticalLabel(ws, actionT, true);
                            }
                            //// Si la table est neuve
                            //if (actionT.ListRows.Count == 0)
                            //{
                            //    try
                            //    {
                            //        actionT.TotalsRowRange.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //    }
                            //    catch
                            //    {
                            //        actionT.ListRows.Add();
                            //        actionT.TotalsRowRange.Offset[1, 0].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //        actionT.TotalsRowRange.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //        //MoveCheckList(checkT, true);
                            //    }
                            //}
                            //else
                            //{
                            //    actionT.TotalsRowRange.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //}
                            break;

                        case TEST.TABLE.TYPE.CHECK:
                            checkT.ListRows.Add();
                            UpdateCheckListHeight(checkT);
                            CreateTest.UpdateVerticalLabel(ws, checkT, true);
                            //try
                            //{
                            //    checkT.DataBodyRange.Offset[1, 0].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //}
                            //catch
                            //{
                            //    checkT.ListRows.Add();
                            //    checkT.DataBodyRange.Offset[1, 0].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //    checkT.DataBodyRange.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //}
                            break;
                        default:
                            break;
                    }
                }
        }

        public static void RemoveVariable(Excel.Application xlsApp, TEST.TABLE.TYPE type, EditingZone editingMode)
        {
                ExcelApplication.setInstance(xlsApp);
                if (General.isActivesheet_a_PR_Test())
                {
                    Worksheet ws = ExcelApplication.getInstance().ActiveSheet;
                    string testNumber = General.getTestNumber();
                    ListObject checkT = ws.ListObjects[TEST.TABLE.PREFIX.CHECK + testNumber];

                    switch (type)
                    {
                        case TEST.TABLE.TYPE.ACTION:
                            ListObject actionT = ws.ListObjects[TEST.TABLE.PREFIX.ACTION + testNumber];
                            // If the table is not empty
                            if (actionT.ListRows.Count > 1)
                            {
                                actionT.ListRows[actionT.ListRows.Count].Delete();

                                // Update titles size according to their sizes
                                CreateTest.UpdateVerticalLabel(ws, actionT, false);

                                // Move check title up
                                MoveCheckList(checkT, false);

                                //Range checkTitle = actionT.TotalsRowRange.Offset[3, -1].Cells[1, 1];
                                //checkTitle.Cut(checkTitle.Offset[-1, 0]);

                                UpdateCheckListHeight(checkT);
                            }
                            else if (actionT.ListRows.Count == 1)
                            {
                                actionT.DataBodyRange.ClearContents();
                            }
                            break;

                        case TEST.TABLE.TYPE.CHECK:
                            if (checkT.ListRows.Count > 1)
                            {
                                checkT.ListRows[checkT.ListRows.Count].Delete();
                                UpdateCheckListHeight(checkT);
                                CreateTest.UpdateVerticalLabel(ws, checkT, false);
                            }
                            else if (checkT.ListRows.Count == 1)
                            {
                                // Show headers in order to not nullify checkT
                                checkT.ShowHeaders = true;
                                checkT.DataBodyRange.ClearContents();
                                checkT.ShowHeaders = false;
                            }
                            break;
                        default:
                            break;
                    }
                }
        }

        /// <summary>
        /// Move check title up or down
        /// </summary>
        static void MoveCheckList(ListObject checkT, Boolean downDirection)
        {
            // down
            if (downDirection)
            {
                if (checkT.ListRows.Count > 1)
                {
                    Range checkTitle = checkT.Range.Offset[-1, -1].Columns[1];
                    try
                    {
                        checkTitle.UnMerge();
                    }
                    catch { }
                    checkTitle.Cut(checkTitle.Offset[1, 0]);
                    try
                    {
                        checkTitle.Merge();
                    }
                    catch { }
                }
                else
                {
                    Range checkTitle = checkT.Range.Offset[-1, -1].Cells[1, 1];
                    checkTitle.Cut(checkTitle.Offset[1, 0]);
                }
            }
            else //up
            {
                if (checkT.ListRows.Count > 1)
                {
                    Range checkTitle = checkT.Range.Offset[1, -1].Columns[1]; 
                    try
                    {
                        checkTitle.UnMerge();
                    }
                    catch { }
                    checkTitle.Cut(checkTitle.Offset[-1, 0]);
                    try
                    {
                        checkTitle.Merge();
                    }
                    catch { }
                }
                else
                {
                    Range checkTitle = checkT.Range.Offset[1, -1].Cells[1, 1];
                    checkTitle.Cut(checkTitle.Offset[-1, 0]);
                }
            }
                
            UpdateCheckListHeight(checkT);
        }

        static void UpdateCheckListHeight(ListObject checkT)
        {
            try
            {
                switch (checkT.ListRows.Count)
                {
                    case 0: checkT.Range.RowHeight = 42; break;
                    case 1: checkT.Range.RowHeight = 42; break;
                    case 2: checkT.Range.RowHeight = 25; break;
                    default: checkT.Range.RowHeight = 15; break;
                }
            }
            catch { }
        }

    }
}
