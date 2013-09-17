using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ValToolMgrInt;
using System.Text.RegularExpressions;
using Excel = NetOffice.ExcelApi;
using ValToolMgrDna.Report;

namespace ValToolMgrDna.ExcelSpecific
{
    class TestSheetParser
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public const string TABLE_PREFIX = "Table_";
        public const string PR_TEST_ACTION = "Action";
        public const string PR_TEST_CHECK = "Check";
        public const string PR_TEST_DESCRIPTION = "Desc";
        public const string PR_TEST_STEP_PATERN = "STEP 1";
        public const string PR_TEST_PREFIX = "Test";
        public const string PR_TEST_SCENARIO_PREFIX = "TS_";
        public const string PR_TEST_TABLE_ACTION_PREFIX = TABLE_PREFIX + PR_TEST_ACTION + "_";
        public const string PR_TEST_TABLE_CHECK_PREFIX = TABLE_PREFIX + PR_TEST_CHECK + "_";
        public const string PR_TEST_TABLE_DESCRIPTION_PREFIX = TABLE_PREFIX + PR_TEST_DESCRIPTION + "_";

        private Excel.Worksheet sheet;
        Excel.ListObject header;
        string actionTableName;
        string checkTableName;
        Excel.ListObject loActionsTable;
        Excel.ListObject loChecksTable;
        TableColumnsStructure tableStructure;
        Excel.ListColumns lcActionsTableColumns;
        Excel.ListColumns lcChecksTableColumns;

        private SheetReport report;

        public enum TableTypes
        {
            TABLE_ACTIONS,
            TABLE_CHECKS,
            TABLE_HEADER
        }

        private class TableColumnsStructure
        {
            public int TargetColumnIndex = -1;
            public int LocationColumnIndex = -1;
            public int PathColumnIndex = -1;
            public int FirstColumnIndex = -1;

            public bool isValid()
            {
                return (TargetColumnIndex >= 0 && LocationColumnIndex >= 0 && PathColumnIndex >= 0 && FirstColumnIndex >= 0) 
                    && (TargetColumnIndex != LocationColumnIndex && PathColumnIndex != FirstColumnIndex && TargetColumnIndex != FirstColumnIndex);
            }

            public void setFirstColumnIndex()
            {
                FirstColumnIndex = Math.Max(TargetColumnIndex, Math.Max(LocationColumnIndex, PathColumnIndex)) + 1;
            }
        }

        public TestSheetParser(Excel.Worksheet sheet, string headerTableName, string actionsTableName, string checksTableName, SheetReport report)
        {
            this.report = report;
            this.sheet = sheet;
                        if (!Regex.IsMatch(sheet.Name, PR_TEST_PREFIX + ".*"))
                throw new FormatException(String.Format("Sheet name doesn't comply with naming rules (begins with \"{0}\").", PR_TEST_PREFIX));

            try
            {
                logger.Debug(String.Format("Trying to retrieve action table \"{0}\".", actionsTableName));
                loActionsTable = sheet.ListObjects[actionsTableName];

                actionTableName = String.Format("'{0}'!{1}", sheet.Name, actionsTableName);
                checkTableName = String.Format("'{0}'!{1}", sheet.Name, checksTableName);
                    
                logger.Debug(String.Format("Extracting columns for action table."));
                lcActionsTableColumns = loActionsTable.ListColumns;
            }
            catch (Exception ex)
            {
                logger.Error(String.Format("Action table \"{0}\" retrieval has failed.", actionsTableName), ex);
                throw new FormatException(String.Format("Action table \"{0}\" retrieval has failed.", actionsTableName));
            }

            try
            {
                logger.Debug(String.Format("Trying to retrieve test table \"{0}\".", checksTableName));
                loChecksTable = sheet.ListObjects[checksTableName];

                logger.Debug(String.Format("Extracting columns for checks table."));
                lcChecksTableColumns = loChecksTable.ListColumns;
            }
            catch (Exception ex)
            {
                logger.Error(String.Format("Check table \"{0}\" retrieval has failed.", checksTableName), ex);
                throw new FormatException(String.Format("Check table \"{0}\" retrieval has failed.", checksTableName));
            }

            if(lcActionsTableColumns.Count != lcChecksTableColumns.Count)
                throw new FormatException(String.Format("Action ({0} columns) and check ({1} columns)  tables has not same number of columns", lcActionsTableColumns.Count, lcChecksTableColumns.Count));

        }

        public static CTest parseTest(string title, Excel.Worksheet sheet, WorkbookParser.ExcelTestStruct tableRefs, SheetReport report)
        {
            logger.Info(String.Format("Beginning Analysis of sheet {0}, using arrays {1} and {2}", sheet.Name, tableRefs.actionTableName, tableRefs.testTableName));

            TestSheetParser analyser = new TestSheetParser(sheet, tableRefs.descrTableName, tableRefs.actionTableName, tableRefs.testTableName, report); 
 

            logger.Debug("Sheet passed validity tests successfully");

            CTest test = null;
            test = analyser.parseAsTest(title);

            return test;
        }

        private CTest parseAsTest(string title) 
        {
            logger.Debug(String.Format("Extracting columns for action table."));
            Excel.ListColumns lcActionsTableColumns = loActionsTable.ListColumns;

            tableStructure = checkAndDetermineTablecolumns(lcActionsTableColumns);

            object[,] actionsValues = preloadTable(this.actionTableName);

            logger.Debug(String.Format("Extracting columns for checks table."));
            Excel.ListColumns lcChecksTableColumns = loChecksTable.ListColumns;

            object[,] checksValues = preloadTable(this.checkTableName);

            CTest parseSingleTest = new CTest(title, "Description");
            logger.Debug(String.Format("Creating Test : {0}", parseSingleTest.ToString()));

            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            //' Writing inputs
            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            logger.Debug(String.Format("Found {0} Excel columns to process.", lcActionsTableColumns.Count));
            for (int CurrentColumn = tableStructure.FirstColumnIndex; CurrentColumn < lcActionsTableColumns.Count; CurrentColumn++)
            {
                logger.Info(String.Format("Processing Column {0}.", lcActionsTableColumns[CurrentColumn+1].Name));
                CStep o_step = new CStep(lcActionsTableColumns[CurrentColumn+1].Name+" : Title retrieval " + getComment(), "Action comment retrieval " + getComment(), "Checks comment retrieval " + getComment());

                logger.Debug(String.Format("Processing Actions table."));
                fillWithActions(o_step, TableTypes.TABLE_ACTIONS, loActionsTable, actionsValues, CurrentColumn);

                logger.Debug(String.Format("Processing Timer table."));
                addTempoIfExists(o_step, loActionsTable, CurrentColumn);

                logger.Debug(String.Format("Processing Checks table."));
                fillWithActions(o_step, TableTypes.TABLE_CHECKS, loChecksTable, checksValues, CurrentColumn);

                logger.Debug(String.Format("Adding step to results."));
                parseSingleTest.Add(o_step);
            }
            return parseSingleTest;
        }

        private TableColumnsStructure checkAndDetermineTablecolumns(Excel.ListColumns lcActionsTableColumns)
        {
            TableColumnsStructure tableStructure = new TableColumnsStructure();

            for (int CurrentColumn = 1; CurrentColumn < 5; CurrentColumn++)
            {
                if (lcActionsTableColumns[CurrentColumn].Name.Equals("Target"))
                    tableStructure.TargetColumnIndex = CurrentColumn - 1; // Indexes from Excel are starting from 1, and we are using 0 based indexes
                if (lcActionsTableColumns[CurrentColumn].Name.Equals("Location"))
                    tableStructure.LocationColumnIndex = CurrentColumn - 1; // Indexes from Excel are starting from 1, and we are using 0 based indexes
                if (lcActionsTableColumns[CurrentColumn].Name.Equals("Path"))
                    tableStructure.PathColumnIndex = CurrentColumn - 1; // Indexes from Excel are starting from 1, and we are using 0 based indexes
            }

            tableStructure.setFirstColumnIndex();

            if(!tableStructure.isValid())
                throw new FormatException(String.Format("Table doesn't contains all necessary columns headers : {0}", tableStructure.ToString()));

            return tableStructure;
        }

        private object[,] preloadTable(string namedRange)
        {

            // Get a reference to the current selection
            object selection = XlCall.Excel(XlCall.xlfEvaluate, namedRange);
            if(selection is ExcelError)
            {
                throw new FormatException(String.Format("Excel returned an error : {0}", ((ExcelError)selection)));
            }
            
            // Get the value of the selection
            object selectionContent = ((ExcelReference)selection).GetValue();
            //object evalResult = XlCall.Excel(XlCall.xlfEvaluate, formula_text);
            // Make sure we dereference if needed.
            //return XlCall.Excel(XlCall.xlCoerce, evalResult); 
            if (selectionContent is object[,])
            {
                return (object[,])selectionContent;
            }
            else
                throw new Exception(String.Format("Calling named range \"{0}\" failed.", namedRange));
        }

        private void fillWithActions(CStep o_step, TableTypes typeOfTable, Excel.ListObject tableRef, object[,] table, int ColumnIndex)
        {
            logger.Debug(String.Format("Found {0} Excel lines to process.", table.GetLength(0)));
            for (int line = 0; line < table.GetLength(0); line++)
            {
                object CellValue = table[line, ColumnIndex];

                if(!(CellValue is ExcelDna.Integration.ExcelEmpty))
                {
                    string Target = "";
                    if (table[line, tableStructure.TargetColumnIndex] is string) Target = (string)table[line, tableStructure.TargetColumnIndex];
                    string Path = "";
                    if (table[line, tableStructure.PathColumnIndex] is string) Path = (string)table[line, tableStructure.PathColumnIndex];
                    string Location = "";
                    if (table[line, tableStructure.LocationColumnIndex] is string) Location = (string)table[line, tableStructure.LocationColumnIndex];

                    logger.Debug(String.Format("Found item [Target={0}, Location={1}, Path={2}, Value={3}].", Target, Location, Path, CellValue));

                    try
                    {
                        logger.Debug(String.Format("Analysing current item."));
                        CInstruction o_instruction = detectAndBuildInstruction(Target, Location, Path, CellValue, typeOfTable);
                        if (typeOfTable == TableTypes.TABLE_ACTIONS)
                        {
                            logger.Debug("Adding item to list of actions to perform");
                            o_step.actions.Add(o_instruction);
                        }
                        else if (typeOfTable == TableTypes.TABLE_CHECKS)
                        {
                            logger.Debug("Adding item to list of checks to perform");
                            o_step.checks.Add(o_instruction);
                        }
                        else
                            throw new NotImplementedException(String.Format("This type of table ({0}) is not currently implemented", typeOfTable));
                    }
                    catch(InvalidCastException ex)
                    {
                        logger.Error("Problem when trying to find an equivalence for item.", ex);
                        report.add(new MessageReport("Invalid value in cell", tableRef.Range[line + 2, ColumnIndex + 1], ex.Message, Criticity.Error));
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Invalid item processed.", ex);
                        report.add(new MessageReport("Cell problem", tableRef.Range[line + 2, ColumnIndex + 1], ex.Message, Criticity.Error));
                    }
                }
            }
        }

        private CInstruction detectAndBuildInstruction(string Target, string Location, string Path, object CellValue, TableTypes typeOfTable)
        {
            CInstruction Instruction;

            string CellValueStr = Convert.ToString(CellValue);

            List<char> detectedChars = extractSpecialProperties(ref Target, ref CellValueStr);
            if(detectedChars.Count > 0) logger.Debug(String.Format("Found {0} special properties.", detectedChars.Count));

                if (typeOfTable == TableTypes.TABLE_ACTIONS)
                {
                    if (CellValueStr.Equals("U"))
                    {
                        Instruction = new CInstrUnforce();
                        logger.Debug(String.Format("Detected Unforce step."));
                        Instruction.data = VariableParser.parseAsVariable(Target, Location, Path, null);
                    }
                    else if (String.Compare(Target, "@POPUP@") == 0)
                    {
                        Instruction = new CInstrPopup();
                        logger.Debug(String.Format("Detected Popup."));
                        Instruction.data = CellValueStr;
                    }
                    else
                    {
                        Instruction = new CInstrForce();
                        logger.Debug(String.Format("Detected Force step."));
                        Instruction.data = VariableParser.parseAsVariable(Target, Location, Path, CellValueStr);
                    }
                }
                else if (typeOfTable == TableTypes.TABLE_CHECKS)
                {
                    if (String.Compare(Target, "@POPUP@") == 0)
                    {
                        Instruction = new CInstrPopup();
                        logger.Debug(String.Format("Detected Popup."));
                        Instruction.data = CellValueStr;
                    }
                    else
                    {
                        Instruction = new CInstrTest();
                        logger.Debug(String.Format("Detected Test step."));
                        Instruction.data = VariableParser.parseAsVariable(Target, Location, Path, CellValueStr);
                    }
                }
                else
                {
                    throw new NotImplementedException("This step is not recognized as a correct step");
                }

                
                Instruction.ForceFailed = detectedChars.Contains('F');
                Instruction.ForcePassed = detectedChars.Contains('P');
                Instruction.Skipped = detectedChars.Contains('S');
                
                return Instruction;
        }

        private List<char> extractSpecialProperties(ref string TargetTotest, ref string valueToTest)
        {
            List<char> table = new List<char>();

            // Here we call Regex.Match.
            Match match = Regex.Match(valueToTest, @"^{([SPF]+)}(.*)",
                RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                string key = match.Groups[1].Value;
                valueToTest = match.Groups[2].Value;
                char[] tab = key.ToCharArray();
                table.AddRange(tab);

            }

            // Here we call Regex.Match.
            match = Regex.Match(TargetTotest, @"^{([SPF]+)}(.*)",
                RegexOptions.IgnoreCase);

            // Here we check the Match instance.
            if (match.Success)
            {
                // Finally, we get the Group value and display it.
                string key = match.Groups[1].Value;
                TargetTotest = match.Groups[2].Value;
                char[] tab = key.ToCharArray();
                table.AddRange(tab);
            }
            return table;
        }

        private void addTempoIfExists(CStep o_step, Excel.ListObject loSourceFiles, int ColumnIndex) 
        {
            //'Delay retrieval. We know that data is contained inside Total line property
            object delay = loSourceFiles.TotalsRowRange.Cells[1, ColumnIndex + 1].Value; // We get values from excel, and array indexes begin with 1, not 0

            if (delay != null)
            {
                try
                {
                    logger.Debug(String.Format("Trying to retrieve temporisation with value \"{0}\".", delay));
                    CInstrWait o_tempo = new CInstrWait();
                    o_tempo.data = Convert.ToInt32(delay);

                    logger.Debug("Adding temporisation to results");
                    o_step.actions.Add(o_tempo);
                }
                catch(Exception ex)
                {
                    logger.Error("Failed to parse temporisation.", ex);
                    report.add(new MessageReport("Invalid value for temporisation", loSourceFiles.TotalsRowRange.Cells[1, ColumnIndex + 1], String.Format("Invalid value, an integer was expected, but we had {0}", delay ), Criticity.Error));
                }
            }
        }

        public void repairSheet() 
        {
            Excel.ListObjects ListOfRanges = this.sheet.ListObjects;
            List<Excel.ListObject> listActionsTable = new List<Excel.ListObject>();
            List<Excel.ListObject> listChecksTable = new List<Excel.ListObject>();
            List<Excel.ListObject> listDescriptionTable = new List<Excel.ListObject>();
            List<Excel.ListObject> listSpecialActionsTable = new List<Excel.ListObject>();

            foreach(Excel.ListObject obj in ListOfRanges)
            {
                logger.Debug(String.Format("Analysing Range {0}", MessageReport.printRange(obj.Range)));
            }
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
