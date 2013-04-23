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

            CTestContainer t = WorkbookParser.parseTestsOfWorkbook(application.ActiveWindow.SelectedSheets);

            XlCall.Excel(XlCall.xlcAlert, "Generation is finished"); 
        }

        [ExcelCommand(MenuText = "Workbook_Open")]
        public static void TPL_Workbook_Open()
        {
            XlCall.Excel(XlCall.xlcAlert, "Workbook_Open"); 
    //        'Refresh linked data sources
    //Me.RefreshAll
    
    //Application.EnableEvents = True

    //For Each sh In Me.Worksheets
    //    ' Hidden sheets are refences. They don't have to be count
    //    If Not sh.Visible = xlSheetHidden Then
    //        For Each oList In sh.ListObjects
    //            Set nm_c = MyName(oList.Name & ROWS, CStr(oList.ListRows.Count))
    //            Set nm_c = MyName(oList.Name & COLUMNS, CStr(oList.ListColumns.Count))
    //        Next oList
    //    End If
    //Next sh
        }
    }
}
