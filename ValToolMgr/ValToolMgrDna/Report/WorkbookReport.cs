using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Antlr4.StringTemplate;
using Antlr4.StringTemplate.Misc;
using System.Globalization;
using System.Reflection;

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
            string URIFilename = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase) + Path.DirectorySeparatorChar + "templates" + Path.DirectorySeparatorChar + "ST-HtmlReport" + Path.DirectorySeparatorChar;
            Uri uri = new Uri(URIFilename);
            logger.Debug("Defining Template directory for HTML report to " + uri.LocalPath);
            TemplateGroup group = new TemplateGroupDirectory(uri.LocalPath, '$', '$');

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
