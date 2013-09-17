using System;
using System.Collections;
using System.IO;
using System.Reflection;
using ValToolMgrDna.Report;
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

        public static CTestContainer parseTestsOfWorkbook(Excel.Sheets sheets, string filename)
        {
            logger.Info("Begin Analysis of selected sheets");

            CTestContainer listOfTests = new CTestContainer();
            WorkbookReport report = new WorkbookReport(filename);

            foreach (Excel.Worksheet wsCurrentTestSheet in sheets)
            {
                logger.Debug(String.Format("Processing sheet \"{0}\".", wsCurrentTestSheet.Name));
                SheetReport sheetReport = new SheetReport(wsCurrentTestSheet.Name);

                try
                {
                    ExcelTestStruct tableRefs = findTablesInSheet(wsCurrentTestSheet, sheetReport);
                    CTest result = TestSheetParser.parseTest(wsCurrentTestSheet.Name, wsCurrentTestSheet, tableRefs, sheetReport);

                    logger.Debug("Adding sheet to result list");

                    listOfTests.Add(result);

                }
                catch (Exception ex)
                {
                    logger.Fatal("Sheet cannot be parsed : ", ex);
                    sheetReport.add(new MessageReport("Parsing error", "Sheet", "Sheet was not analysed. Message is : "+ ex.Message, Criticity.Critical));
                }
                report.add(sheetReport);
            }

            
            string URIFilename = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase) + Path.DirectorySeparatorChar + "Report.html";
            Uri uri = new Uri(URIFilename);
            logger.Debug("Writing report in "+URIFilename);

            report.printReport(uri.LocalPath);

            if (report.NbrMessages > 0)
            {
                System.Diagnostics.Process.Start(URIFilename);
            }

            return listOfTests;
        }

        private static ExcelTestStruct findTablesInSheet(Excel.Worksheet wsCurrentTestSheet, SheetReport report)
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
                {
                    report.add(new MessageReport("Unrecognised table type", obj.Range, String.Format("\"{0}\" is not an authorized name for table.", tableName), Criticity.Critical));
                    logger.Error(String.Format("\"{0}\" is not an authorized name for table.", tableName));
                }
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
