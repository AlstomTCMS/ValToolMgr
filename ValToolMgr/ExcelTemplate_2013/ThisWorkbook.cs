using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelTemplate_2013
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            //Check in each sheet the kies data

            // get evol list row number
            MessageBox.Show("This startup");

            // get SwVTP tests's list
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Open += new Excel.WorkbookEvents_OpenEventHandler(ThisWorkbook_Open);
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        void ThisWorkbook_Open()
        {
            //Check in each sheet the kies data

            // get evol list row number
            MessageBox.Show("This Open");

            // get SwVTP tests's list
        }

        #endregion

    }
}
