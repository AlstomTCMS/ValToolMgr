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
using System.Text.RegularExpressions;

namespace ExcelTemplate_2013
{
    public partial class Evol
    {
        public ListObject EvolList
        {
            get {
                foreach (ListObject list in this.ListObjects)
                {
                    if (Regex.IsMatch(list.Name, "evolListobject"))
                    {
                        return list;
                    }
                }
                return null;
            }
        }

        private int _lastEvolListRowNumber;
        public int LastEvolListRowNumber
        {
            get
            {
                return _lastEvolListRowNumber;
            }
        }

        private void Feuil2_Startup(object sender, System.EventArgs e)
        {
            if (EvolList != null)
            {
                _lastEvolListRowNumber = EvolList.ListRows.Count;
                MessageBox.Show("evollist rows: " + _lastEvolListRowNumber);
                EvolList.BeforeAddDataBoundRow += new BeforeAddDataBoundRowEventHandler(EvolList_BeforeAddDataBoundRow);

                //hide all sheet
                this.Cells.Hidden = true;

                //unhide data
                EvolList.Range.Hidden = false;
            }
        }

        void EvolList_BeforeAddDataBoundRow(object sender, BeforeAddDataBoundRowEventArgs e)
        {
            EvolList.Range.Offset[1, 0].EntireRow.Hidden = false;
        }

        private void Feuil2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Feuil2_Startup);
            this.Shutdown += new System.EventHandler(Feuil2_Shutdown);
        }

        #endregion

    }
}
