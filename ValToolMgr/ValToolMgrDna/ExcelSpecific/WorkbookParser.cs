using System;
using ExcelDna.Integration;
using System.Collections;
using ValToolMgrInt;
using Excel = NetOffice.ExcelApi;

namespace ValToolMgrDna.ExcelSpecific
{
    class WorkbookParser
    {
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


        public struct ExcelTestStruct {
            public string actionTableName;
            public string testTableName;
            public string descrTableName;
        }

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static CTestContainer parseTestsOfWorkbook(Excel.Sheets sheets)
        {
            logger.Info("Begin Analysis of selected sheets");
            CTestContainer listOfTests = new CTestContainer();
            foreach (Excel.Worksheet wsCurrentTestSheet in sheets)
            {
                try
                {
                    logger.Debug(String.Format("Processing sheet \"{0}\".", wsCurrentTestSheet.Name));

                    ExcelTestStruct tableRefs = findTablesInSheet(wsCurrentTestSheet);
                    CTest result = TestSheetParser.parseTest(wsCurrentTestSheet.Name, wsCurrentTestSheet, tableRefs);

                    logger.Debug("Adding sheet to result list");
                    listOfTests.Add(result);
                }
                catch (Exception ex)
                {
                    logger.Error("Sheet cannot be parsed : ", ex);
                    XlCall.Excel(XlCall.xlcAlert, "Sheet \"" + wsCurrentTestSheet.Name + " was not analysed. Message is : "+ ex.Message); 
                }
            }
            return listOfTests;
        }

        private static ExcelTestStruct findTablesInSheet(Excel.Worksheet wsCurrentTestSheet)
        {
            Excel.ListObjects ListOfRanges = wsCurrentTestSheet.ListObjects;

            SortedList listActionTables = new SortedList();
            SortedList listCheckTables = new SortedList();
            SortedList listDescrTables = new SortedList();
            foreach (Excel.ListObject obj in ListOfRanges)
            {
                int range = obj.Range.Row;
                String tableName = obj.Name;
                if (tableName.StartsWith(PR_TEST_TABLE_ACTION_PREFIX))
                    listActionTables.Add(range, tableName);
                else if (tableName.StartsWith(PR_TEST_TABLE_CHECK_PREFIX))
                    listCheckTables.Add(range, tableName);
                else if (tableName.StartsWith(PR_TEST_TABLE_DESCRIPTION_PREFIX))
                    listDescrTables.Add(range, tableName);
                else
                    logger.Error(String.Format("Not recognized type of named range : \"{0}\"", tableName));
            }

            if(listActionTables.Count == 1 && listCheckTables.Count == 1 && listDescrTables.Count == 1)
            {
                ExcelTestStruct n = new ExcelTestStruct();
                n.actionTableName = (string)listActionTables.GetByIndex(0);
                n.testTableName = (string)listCheckTables.GetByIndex(0);
                n.descrTableName = (string)listDescrTables.GetByIndex(0);
                return n;
            }
            else
                throw new NotImplementedException(String.Format("It is not currently possible to repair sheet \"{0}\"", wsCurrentTestSheet.Name));
        }
    }
}
