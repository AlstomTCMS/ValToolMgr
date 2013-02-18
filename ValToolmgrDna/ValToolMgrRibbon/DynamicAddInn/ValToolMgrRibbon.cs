using System;
using ExcelDna.Integration;
using ValToolMgrDna.Interface;
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
    public class ValToolMgrRibbon
    {
        [ExcelFunction(Description="My first Excel-DNA function")]
        public static string MyFirstFunction(string name)
        {
            return "Hello Ribbon " + name;
        }
        
        [ExcelCommand(MenuText = "Say Hello")] 
        public static void SayHello() 
        {
            Excel.Application application = new Excel.Application(null, ExcelDnaUtil.Application);

            CTestContainer t = WorkbookParser.parseTestsOfWorkbook(application.ActiveWindow.SelectedSheets);

            XlCall.Excel(XlCall.xlcAlert, "Generation is finished"); 
        } 
    }
}
