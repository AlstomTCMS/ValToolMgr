using System;
using ExcelDna.Integration;

namespace ValToolMgrDna.ExcelSpecific
{
    class Addin : IExcelAddIn
    {
        void IExcelAddIn.AutoOpen()
        {
            // Register Ctrl+Shift+H to call SayHello 
            XlCall.Excel(XlCall.xlcOnKey, "^H", "SayHello");

            //ThisWorkbook.Name = System.IO.Path.GetFileName(XlCall.Excel(XlCall.xlGetName));
            //ThisWorkbook.Path = System.IO.Path.GetDirectoryName(XlCall.Excel(XlCall.xlGetName));

            //Factory.Initialize();          
        }

        void IExcelAddIn.AutoClose()
        {

            // Clear the registration if the add-in is unloaded 
            XlCall.Excel(XlCall.xlcOnKey, "^H");
        }

 
    }
}
