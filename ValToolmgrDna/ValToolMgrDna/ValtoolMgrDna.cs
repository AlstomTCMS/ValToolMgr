using System;
using ExcelDna.Integration;

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

            foreach (Excel.Worksheet sheet in application.ActiveWindow.SelectedSheets)
            {
                XlCall.Excel(XlCall.xlcAlert, "Selected sheet : " + sheet.Name); 
            }
        } 
    }
}
