using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    return new Ribbon();
        //}


        public Worksheet ListObject_Change()
        {
            Worksheet vstoWorksheet =
        Globals.Factory.GetVstoObject(this.Application.ActiveWorkbook.Worksheets[1]);
            //ListObject list1 =
            //    vstoWorksheet.Controls.AddListObject(
            //    vstoWorksheet.Range["A1", "C4"], "list1");
            //list1.Change += new ListObjectChangeHandler(list1_Change);
            return vstoWorksheet;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
