using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = NetOffice.ExcelApi;
using System.Text.RegularExpressions;

namespace ValToolMgrDna.Report
{
    public enum Criticity { Debug, Information, Warning, Error, Critical };

    public class MessageReport
    {
        public string title;
        public string localisation;
        public string description;
        public Criticity criticity;

        public MessageReport(string title, string localisation, string description, Criticity criticity)
        {
            this.title = title;
            this.localisation = localisation;
            this.description = description;
            this.criticity = criticity;
        }

        public MessageReport(string title, Excel.Range excelRange, string description, Criticity criticity)
        {
            this.title = title;
            this.localisation = printRange(excelRange);
            this.description = description;
            this.criticity = criticity;
        }

        public static string printRange(Excel.Range range)
        {
            string text = range.Address;
            text = Regex.Replace(text, "\\$", "");

            return "["+text+"]";
        }
    }


}
