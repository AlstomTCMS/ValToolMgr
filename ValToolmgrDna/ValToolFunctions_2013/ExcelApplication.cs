using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ValToolFunctions_2013
{
    /// <summary>
    /// Singleton class for the Excel application
    /// </summary>
    static class ExcelApplication
    {
        static Application Instance;

        public static void setInstance(Application excelApp)
        {
            Instance = excelApp;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Application getInstance()
        {
            if (Instance == null)
            {
                throw new ExcelApplicationMissingException();
            }
            return Instance;
        }
    }
}
