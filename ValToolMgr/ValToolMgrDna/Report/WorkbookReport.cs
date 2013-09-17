using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Antlr4.StringTemplate;
using Antlr4.StringTemplate.Misc;
using System.Globalization;

namespace ValToolMgrDna.Report
{
    public class WorkbookReport : ItemReport
    {

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public WorkbookReport(string name) : base (name)
        {
        }

        public void printReport(string filename)
        {
            TemplateGroup group = new TemplateGroupDirectory("C:\\ValToolMgr\\templates\\ST-HtmlReport\\", '$', '$');

            ErrorBuffer errors = new ErrorBuffer();
            group.Listener = errors;
            group.Load();

            Template st = group.GetInstanceOf("MainTemplate");

            st.Add("Report", this);

            string result = st.Render();

            if (errors.Errors.Count > 0)
            {
                foreach (TemplateMessage m in errors.Errors)
               {
                    logger.Error(m);
                    throw new Exception(m.ToString());
                }
            }

            StreamWriter output = new StreamWriter(filename, false, Encoding.GetEncoding("UTF-8"));

            output.Write(result);
            output.Close();
        }
    }
}
