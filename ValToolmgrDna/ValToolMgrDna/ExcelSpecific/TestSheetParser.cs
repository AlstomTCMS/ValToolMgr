using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.ExcelApi.GlobalHelperModules;

namespace ValToolMgrDna.ExcelSpecific
{
    class TestSheetParser
    {
        public const string TABLE_PREFIX = "Table_";
        public const string PR_TEST_ACTION = "Action";
        public const string PR_TEST_CHECK = "Check";
        public const string PR_TEST_DESCRIPTION = "Desc";
        public const string PR_TEST_STEP_PATERN = "STEP 1";
        public const string PR_TEST_PREFIX = "Test_";
        public const string PR_TEST_SCENARIO_PREFIX = "TS_";
        public const string PR_TEST_TABLE_ACTION_PREFIX = TABLE_PREFIX + PR_TEST_ACTION + "_";
        public const string PR_TEST_TABLE_CHECK_PREFIX = TABLE_PREFIX + PR_TEST_CHECK + "_";
        public const string PR_TEST_TABLE_DESCRIPTION_PREFIX = TABLE_PREFIX + PR_TEST_DESCRIPTION + "_";

        private Excel.Worksheet sheet;

        public enum TableTypes
        {
            TABLE_ACTIONS,
            TABLE_CHECKS,
            TABLE_HEADER
        }

        public TestSheetParser(Excel.Worksheet sheet)
        {
            this.sheet = sheet;
        }

        public static CTest parseTest(string title, Excel.Worksheet sheet, Excel.ListObject header, Excel.ListObject loActionsTable, Excel.ListObject loChecksTable)
        {
            TestSheetParser analyser = new TestSheetParser(sheet);
            CTest test = null;
            if (analyser.isSheetValid())
            {
                test = analyser.parseAsTest(title, header, loActionsTable, loChecksTable);
            }

            if (test == null) throw new NullReferenceException();
            return test;
        }

        private CTest parseAsTest(string title, Excel.ListObject header, Excel.ListObject loActionsTable, Excel.ListObject loChecksTable) 
        {
            Excel.ListColumns lcActionsTableColumns = loActionsTable.ListColumns;
            Excel.ListColumns lcChecksTableColumns = loChecksTable.ListColumns;

            CTest parseSingleTest = new CTest();

            parseSingleTest.title = title;

            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            //' Writing inputs
            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            for (int CurrentColumn = 3; CurrentColumn <= lcActionsTableColumns.Count; CurrentColumn++)
            {
                // Writing header
                // Debug.Print "Processing Step : " & lcActionsTableColumns.Item(CurrentColumn)
                CStep o_step = new CStep();
                o_step.title = getComment();
                o_step.DescAction = getComment();
                o_step.DescCheck = getComment();

                fillWithActions(o_step, TableTypes.TABLE_ACTIONS, loActionsTable, CurrentColumn);
                addTempoIfExists(o_step, loActionsTable, CurrentColumn);
                fillWithActions(o_step, TableTypes.TABLE_CHECKS, loChecksTable, CurrentColumn);
                parseSingleTest.Add(o_step);
            }
            return parseSingleTest;
        }

        private void fillWithActions(CStep o_step, TableTypes typeOfTable, Excel.ListObject loSourceFiles, int ColumnIndex)
        {
            for (int line = 1; line <= loSourceFiles.ListRows.Count; line++)
            {
                Excel.ListRow lrCurrent = loSourceFiles.ListRows[line];

                string Target = (string)lrCurrent.Range[1, 1].Value;
                string Location = (string)lrCurrent.Range[1, 2].Value;
                object CellValue = lrCurrent.Range[1, ColumnIndex].Value;

                if(CellValue != null)
                {
                    CInstruction o_instruction = detectAndBuildInstruction(Target, Location, CellValue, typeOfTable);

                    if(typeOfTable == TableTypes.TABLE_ACTIONS)
                    {
                        o_step.actions.Add(o_instruction);
                    }
                    else if (typeOfTable == TableTypes.TABLE_CHECKS)
                    {
                        o_step.checks.Add(o_instruction);
                    }
                }
            }
        }

        private CInstruction detectAndBuildInstruction(string Target, string Location, object CellValue, TableTypes typeOfTable)
        {
            CInstruction detectAndBuildInstruction = new CInstruction();

    
            detectAndBuildInstruction.category = CInstruction.actionList.UNIMPLEMENTED;
            detectAndBuildInstruction.data = null;
    
            CVariable o_variable = buildVariable(Target, Location, CellValue);
                
            if (typeOfTable == TableTypes.TABLE_ACTIONS)
            {

                if(CellValue is String && o_variable.value == "U")
                {
                    detectAndBuildInstruction.category = CInstruction.actionList.A_UNFORCE;
                }
                else
                {
                     detectAndBuildInstruction.category = CInstruction.actionList.A_FORCE;
                }
                detectAndBuildInstruction.data = o_variable;
            }
            else if (typeOfTable == TableTypes.TABLE_CHECKS)
            {
                detectAndBuildInstruction.category = CInstruction.actionList.A_TEST;
                detectAndBuildInstruction.data = o_variable;
            }
            return detectAndBuildInstruction;
        }

        private void addTempoIfExists(CStep o_step, Excel.ListObject loSourceFiles, int ColumnIndex) 
        {
            //'Delay retrieval. We know that data is contained inside Total line property
            object delay = loSourceFiles.TotalsRowRange.Cells[1, ColumnIndex].Value;

            if (delay != null)
            {
                CInstruction o_tempo = new CInstruction();
                o_tempo.category = CInstruction.actionList.A_WAIT;
                o_tempo.data = (int)delay;
                o_step.actions.Add(o_tempo);
            }
        }

        private CVariable buildVariable(string Target, string Location, object CellValue) 
        {
            CVariable buildVariable = new CVariable();
            buildVariable.name = Target;
            buildVariable.path = Location;
            buildVariable.value = CellValue;

            if(Target.IndexOf("I:") == 1) 
            {
                buildVariable.typeOfVar = CVariable.E_varType.T_INTEGER;
                buildVariable.name = Target.Substring(3);
            }
            else if(Target.IndexOf("R:") == 1) 
            {
                buildVariable.typeOfVar = CVariable.E_varType.T_REAL;
                buildVariable.name = Target.Substring(3);
            }
            else if(Target.IndexOf("DT:") == 1)
            {
                buildVariable.name = Target.Substring(4);
                buildVariable.typeOfVar = CVariable.E_varType.T_DATE_AND_TIME;
            }
            else
            {
                buildVariable.typeOfVar = CVariable.E_varType.T_BOOLEAN;
            }
            return buildVariable;
        }

        private bool isSheetValid() 
        {
            return Regex.IsMatch(sheet.Name, PR_TEST_PREFIX + ".*");
        }

        private void repairSheet() 
        {
            throw new NotImplementedException();
        }

        private string getComment()
        {
            return "Not implemented";
            // Function getComment(wsCurrentTestSheet As Worksheet, lcTable As ListObject, CurrentColumn As Integer, OldComment As String) As String
            //    Dim ColumnsHeaderPosition As Integer
    
            //    getComment = OldComment
    
            //    xPosition = lcTable.HeaderRowRange.Row - 1
            //    yPosition = lcTable.ListColumns.Item(CurrentColumn).Range.Column
    
            //    If xPosition > 0 And Not IsEmpty(wsCurrentTestSheet.Cells(xPosition, yPosition)) Then
            //        getComment = wsCurrentTestSheet.Cells(xPosition, yPosition).value
            //    End If
            //End Function
        }

    }
}
