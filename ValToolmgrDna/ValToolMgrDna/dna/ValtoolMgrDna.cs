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

namespace ValToolMgrDna
{
    public class ValtoolMgrDna
    {
        [ExcelFunction(Description="My first Excel-DNA function")]
        public static string MyFirstFunction(string name)
        {
            return "Hello " + name;
        }
        
        [ExcelCommand(MenuText = "Say Hello")] 
        public static void SayHello() 
        {
            Excel.Application application = new Excel.Application(null, ExcelDnaUtil.Application);

            CTestContainer container = WorkbookParser.parseTestsOfWorkbook(application.ActiveWindow.SelectedSheets);

            try
            {
                TestStandGen.TestStandGen.genSequence(container, "C:\\macros_alstom\\test\\genTest.seq", "C:\\macros_alstom\\templates\\ST-TestStand3\\");
            }
            catch (Exception ex)
            {
                XlCall.Excel(XlCall.xlcAlert, ex.ToString()); 
            }

            XlCall.Excel(XlCall.xlcAlert, "Generation is finished"); 
        } 
    }
}
