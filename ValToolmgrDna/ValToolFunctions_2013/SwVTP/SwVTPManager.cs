using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;

namespace ValToolFunctions_2013
{
    internal class SwVTPManager
    {
        internal static void AddCategory()
        {
            if (General.isActivesheet_a_SwVTPSheet())
            {
                Worksheet ws = RibbonHandler.Factory.GetVstoObject(RibbonHandler.ExcelApplication.ActiveSheet);
                //ListObject newCategoryTestsTableT = ws.Controls.AddListObject();

            }
        }

    }
}
