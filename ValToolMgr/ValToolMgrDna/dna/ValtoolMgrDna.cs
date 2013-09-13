using System;
using ExcelDna.Integration;
using ValToolMgrInt;
using ValToolMgrDna.ExcelSpecific;

using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Globalization;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.ExcelApi.GlobalHelperModules;
using System.IO;

namespace ValToolMgrDna
{
    public class ValtoolMgrDna
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
       
        [ExcelCommand(MenuText = "Generate sequence")] 
        public static void GenerateSequence() 
        {
            Excel.Application application = new Excel.Application(null, ExcelDnaUtil.Application);

            //CallingFileName();

            string path = application.ActiveWorkbook.FullName;
            string filenameNoExtension = Path.GetFileNameWithoutExtension(path);
            string root = Path.GetDirectoryName(path) + Path.DirectorySeparatorChar;

            CTestContainer container = WorkbookParser.parseTestsOfWorkbook(application.ActiveWindow.SelectedSheets);

            try
            {
                TestStandGen.TestStandGen.genSequence(container, root+filenameNoExtension+".seq", "C:\\macros_alstom\\templates\\ST-TestStand3\\");
            }
            catch (Exception ex)
            {
                logger.Debug("Exception raised : ", ex);
                XlCall.Excel(XlCall.xlcAlert, ex.Message); 
            }

            XlCall.Excel(XlCall.xlcAlert, "Generation is finished"); 
        }
    }
}