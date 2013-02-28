using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ValToolFunctions_2013;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class RibbonVisual
    {
        private void RibbonVisual_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button_NewPR_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

            CreateTest.NewPR(Globals.ThisAddIn.Application);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == "Void set_Interactive(Boolean)")
                {
                    MessageBox.Show("Please, unselect the cell you are editing. This may cause unexecepted behaviours", "Warning !");
                }
            }
            finally
            {
                xlsApp.ScreenUpdating = true;
                if (!xlsApp.Interactive)
                {
                    xlsApp.Interactive = true;
                }
            }
        }

        private void LayoutVersion_DD_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void autoUpdate_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void AddStep_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.AddNewStep(Globals.ThisAddIn.Application);
        }

        private void RemoveStep_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

            TestManager.RemoveStep(Globals.ThisAddIn.Application,EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == "Void set_Interactive(Boolean)")
                {
                    MessageBox.Show("Please, unselect the cell you are editing. This may cause unexecepted behaviours", "Warning !");
                }
            }
            finally
            {
                xlsApp.ScreenUpdating = true;
                if (!xlsApp.Interactive)
                {
                    xlsApp.Interactive = true;
                }
            }
        }

        private void AddActionVar_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

            TestManager.AddVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.ACTION, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == "Void set_Interactive(Boolean)")
                {
                    MessageBox.Show("Please, unselect the cell you are editing. This may cause unexecepted behaviours", "Warning !");
                }
            }
            finally
            {
                xlsApp.ScreenUpdating = true;
                if (!xlsApp.Interactive)
                {
                    xlsApp.Interactive = true;
                }
            }
        }

        private void AddCheckVar_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

            TestManager.AddVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.CHECK, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == "Void set_Interactive(Boolean)")
                {
                    MessageBox.Show("Please, unselect the cell you are editing. This may cause unexecepted behaviours", "Warning !");
                }
            }
            finally
            {
                xlsApp.ScreenUpdating = true;
                if (!xlsApp.Interactive)
                {
                    xlsApp.Interactive = true;
                }
            }
        }

        private void RemoveActionVar_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

            TestManager.RemoveVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.ACTION, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == "Void set_Interactive(Boolean)")
                {
                    MessageBox.Show("Please, unselect the cell you are editing. This may cause unexecepted behaviours", "Warning !");
                }
            }
            finally
            {
                xlsApp.ScreenUpdating = true;
                if (!xlsApp.Interactive)
                {
                    xlsApp.Interactive = true;
                }
            }
        }

        private void RemoveCheckVar_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

                TestManager.RemoveVariable(xlsApp, TEST.TABLE.TYPE.CHECK, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == "Void set_Interactive(Boolean)")
                {
                    MessageBox.Show("Please, unselect the cell you are editing. This may cause unexecepted behaviours", "Warning !");
                }
            }
            finally
            {
                xlsApp.ScreenUpdating = true;
                if (!xlsApp.Interactive)
                {
                    xlsApp.Interactive = true;
                }
            }
        }
    }
}
