using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using ValToolMgrInt;
using System.Text.RegularExpressions;
using Excel = NetOffice.ExcelApi;

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
        public const string PR_TEST_PREFIX = "Test_";
        public const string PR_TEST_SCENARIO_PREFIX = "TS_";
        public const string PR_TEST_TABLE_ACTION_PREFIX = TABLE_PREFIX + PR_TEST_ACTION + "_";
        public const string PR_TEST_TABLE_CHECK_PREFIX = TABLE_PREFIX + PR_TEST_CHECK + "_";
        public const string PR_TEST_TABLE_DESCRIPTION_PREFIX = TABLE_PREFIX + PR_TEST_DESCRIPTION + "_";

        private Excel.Worksheet sheet;
        Excel.ListObject header;
        Excel.ListObject loActionsTable;
        Excel.ListObject loChecksTable;
        Excel.ListColumns lcActionsTableColumns;
        Excel.ListColumns lcChecksTableColumns;

        public enum TableTypes
        {
            TABLE_ACTIONS,
            TABLE_CHECKS,
            TABLE_HEADER
        }

        public TestSheetParser(Excel.Worksheet sheet, string headerTableName, string actionsTableName, string checksTableName)
        {
            this.sheet = sheet;
                        if (!Regex.IsMatch(sheet.Name, PR_TEST_PREFIX + ".*"))
                throw new FormatException(String.Format("Sheet name doesn't comply with naming rules (begins with \"{0}\").", PR_TEST_PREFIX));

            try
            {
                logger.Debug(String.Format("Trying to retrieve action table \"{0}\".", actionsTableName));
                loActionsTable = sheet.ListObjects[actionsTableName];
                    
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

        public static CTest parseTest(string title, Excel.Worksheet sheet, string headerTableName, string actionsTableName, string checksTableName)
        {
            logger.Info(String.Format("Beginning Analysis of sheet {0}, using arrays {1} and {2}", sheet.Name, actionsTableName, checksTableName));

            TestSheetParser analyser = new TestSheetParser(sheet, headerTableName, actionsTableName, checksTableName);
            CTest test = null;

            logger.Debug("Sheet passed validity tests successfully");
            test = analyser.parseAsTest(title);

            return test;
        }

        private CTest parseAsTest(string title) 
        {
            logger.Debug(String.Format("Extracting columns for action table."));
            Excel.ListColumns lcActionsTableColumns = loActionsTable.ListColumns;

            logger.Debug(String.Format("Extracting columns for checks table."));
            Excel.ListColumns lcChecksTableColumns = loChecksTable.ListColumns;

            CTest parseSingleTest = new CTest(title, "Description");
            logger.Debug(String.Format("Creating Test : {0}", parseSingleTest.ToString()));

            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            //' Writing inputs
            //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            logger.Debug(String.Format("Found {0} Excel columns to process.", lcActionsTableColumns.Count));
            for (int CurrentColumn = 3; CurrentColumn <= lcActionsTableColumns.Count; CurrentColumn++)
            {
                logger.Info(String.Format("Processing Column {0}.", lcActionsTableColumns[CurrentColumn].Name));
                CStep o_step = new CStep(lcActionsTableColumns[CurrentColumn].Name+" : Title retrieval " + getComment(), "Action comment retrieval " + getComment(), "Checks comment retrieval " + getComment());

                logger.Debug(String.Format("Processing Actions table."));
                fillWithActions(o_step, TableTypes.TABLE_ACTIONS, loActionsTable, CurrentColumn);

                logger.Debug(String.Format("Processing Timer table."));
                addTempoIfExists(o_step, loActionsTable, CurrentColumn);

                logger.Debug(String.Format("Processing Checks table."));
                fillWithActions(o_step, TableTypes.TABLE_CHECKS, loChecksTable, CurrentColumn);

                logger.Debug(String.Format("Adding step to results."));
                parseSingleTest.Add(o_step);
            }
            return parseSingleTest;
        }

        private void fillWithActions(CStep o_step, TableTypes typeOfTable, Excel.ListObject loSourceFiles, int ColumnIndex)
        {
            logger.Debug(String.Format("Found {0} Excel lines to process.", loSourceFiles.ListRows.Count));
            for (int line = 1; line <= loSourceFiles.ListRows.Count; line++)
            {
                Excel.ListRow lrCurrent = loSourceFiles.ListRows[line];
                logger.Debug(String.Format("Processing Excel line {0}.", lrCurrent.Range.AddressLocal));

                string Target = (string)lrCurrent.Range[1, 1].Value;
                string Location = (string)lrCurrent.Range[1, 2].Value;
                Excel.Range rangeToRetrieve = lrCurrent.Range[1, ColumnIndex];
                object CellValue = rangeToRetrieve.Value;

                if(CellValue != null)
                {
                    logger.Debug(String.Format("Found item [Target={{0}}, Location={{1}}, Value={{2}}].", Target, Location, CellValue));

                    try
                    {
                        logger.Debug(String.Format("Analysing current item."));
                        CInstruction o_instruction = detectAndBuildInstruction(Target, Location, CellValue, typeOfTable);
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
                        XlCall.Excel(XlCall.xlcAlert, "Invalid value in cell " + rangeToRetrieve.Address + " : "+ex.Message);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Invalid item processed.", ex);
                        XlCall.Excel(XlCall.xlcAlert, "Cell problem " + rangeToRetrieve.Address + " : " + ex.Message);
                    }
                }
            }
        }

        private CInstruction detectAndBuildInstruction(string Target, string Location, object CellValue, TableTypes typeOfTable)
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
                        Instruction.data = VariableParser.parseAsVariable(Target, Location, null);
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
                        Instruction.data = VariableParser.parseAsVariable(Target, Location, CellValueStr);
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
                        Instruction.data = VariableParser.parseAsVariable(Target, Location, CellValueStr);
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
            object delay = loSourceFiles.TotalsRowRange.Cells[1, ColumnIndex].Value;

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
                    XlCall.Excel(XlCall.xlcAlert, "Invalid value, expected : (int), have : " + delay.ToString());
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
                string range = obj.Range.AddressLocal;
                logger.Debug(String.Format("Analysing Range {0}", range));                                                                                                                                                 
                //public const string PR_TEST_TABLE_ACTION_PREFIX = TABLE_PREFIX + PR_TEST_ACTION + "_";
                //public const string PR_TEST_TABLE_CHECK_PREFIX = TABLE_PREFIX + PR_TEST_CHECK + "_";
                //public const string PR_TEST_TABLE_DESCRIPTION_PREFIX = TABLE_PREFIX + PR_TEST_DESCRIPTION + "_";
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
