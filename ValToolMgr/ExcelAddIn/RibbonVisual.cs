using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ValToolFunctions_2013;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using ValToolFunctionsStub;
using System.Resources;
using ExcelAddIn.Properties;

namespace ExcelAddIn
{
    public partial class RibbonVisual
    {
        internal const string EXCEPTION_USER_EDITING = "Void set_Interactive(Boolean)";
        internal RibbonHandler ribbonHandler_2013 = new RibbonHandler();


        #region Internal ribbon management

        private void RibbonVisual_Load(object sender, RibbonUIEventArgs e)
        {
            #region UI localization

            button_NewPR.Label = Language.NewPR_label;
            plan2Tests.Label = Language.PlanToTests_label;
            AddStep.Label = Language.AddStep_label;
            RemoveStep.Label = Language.RemoveStep_label;
            AddActionVar.Label = Language.AddActionVar_label;
            RemoveActionVar.Label = Language.RemoveActionVar_label;
            AddCheckVar.Label = Language.AddCheckVar_label;
            RemoveCheckVar.Label = Language.RemoveCheckVar_label;
            testCheck.Label = Language.CheckTest_label;
            testStand.Label = Language.ToTestStand;

            macroVersion.Label = String.Format(Language.macro_version_label, PublishVersion);
            autoUpdate.Label = Language.AutoUpdate_CB_label;
            autoUpdate.Checked = Properties.Settings.Default.AutoUpdate;

            #endregion
        }

        public string PublishVersion
        {
            get
            {
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                {
                    Version ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                    return string.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
                }
                else
                    return Language.NotPublished;
            }
        }

        private void LayoutVersion_DD_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void autoUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.AutoUpdate = autoUpdate.Checked;
            Properties.Settings.Default.Save();
        }

        #endregion

        #region Interactions with the RibbonHandler interface

        private void button_NewPR_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

                ribbonHandler_2013.NewPR(xlsApp, Globals.Factory);
            }
            catch (NotImplementedException ex)
            {
                MessageBox.Show(Language.Exception_NotImplemented);
            }
            catch (ExcelApplicationNotAvailableException ex)
            {
                MessageBox.Show(Language.Exception_ExcelApplication_NotAvailable);
            }
            catch (ExcelApplicationMissingException ex)
            {
                MessageBox.Show(Language.Exception_ExcelApplication_Missing);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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


        private void AddStep_Click(object sender, RibbonControlEventArgs e)
        {
            ribbonHandler_2013.AddStep(Globals.ThisAddIn.Application);
        }

        private void RemoveStep_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

                ribbonHandler_2013.RemoveStep(xlsApp, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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

                ribbonHandler_2013.AddActionVar(Globals.ThisAddIn.Application, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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

                ribbonHandler_2013.AddCheckVar(Globals.ThisAddIn.Application, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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

                ribbonHandler_2013.RemoveActionVar(Globals.ThisAddIn.Application, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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

                ribbonHandler_2013.RemoveCheckVar(xlsApp, EditingZone.NONE);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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
        #endregion

        private void plan2Tests_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

                ribbonHandler_2013.PlanToTests(xlsApp);
            }
            catch (NotImplementedException ex)
            {
                MessageBox.Show(Language.Exception_NotImplemented);
            }
            catch (ExcelApplicationNotAvailableException ex)
            {
                MessageBox.Show(Language.Exception_ExcelApplication_NotAvailable);
            }
            catch (ExcelApplicationMissingException ex)
            {
                MessageBox.Show(Language.Exception_ExcelApplication_Missing);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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

        private void toSwVTD_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            //try
            //{
                //xlsApp.ScreenUpdating = false;
                //xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

                ribbonHandler_2013.extractTests2SwVTD(xlsApp);
            //}
            //catch (NotImplementedException ex)
            //{
            //    MessageBox.Show(Language.Exception_NotImplemented);
            //}
            //catch (ExcelApplicationNotAvailableException ex)
            //{
            //    MessageBox.Show(Language.Exception_ExcelApplication_NotAvailable);
            //}
            //catch (ExcelApplicationMissingException ex)
            //{
            //    MessageBox.Show(Language.Exception_ExcelApplication_Missing);
            //}
            //catch (Exception ex)
            //{
            //    if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
            //    {
            //        MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
            //    }
            //}
            //finally
            //{
            //    xlsApp.ScreenUpdating = true;
            //    if (!xlsApp.Interactive)
            //    {
            //        xlsApp.Interactive = true;
            //    }
            //}
        }

        private void toSwVTDR_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void addCategory_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlsApp = Globals.ThisAddIn.Application;
            try
            {
                xlsApp.ScreenUpdating = false;
                xlsApp.Interactive = false; //http://msdn.microsoft.com/en-us/library/ff841248.aspx

                ribbonHandler_2013.AddCategory(xlsApp, Globals.Factory);
            }
            catch (NotImplementedException ex)
            {
                MessageBox.Show(Language.Exception_NotImplemented);
            }
            catch (ExcelApplicationNotAvailableException ex)
            {
                MessageBox.Show(Language.Exception_ExcelApplication_NotAvailable);
            }
            catch (ExcelApplicationMissingException ex)
            {
                MessageBox.Show(Language.Exception_ExcelApplication_Missing);
            }
            catch (Exception ex)
            {
                if (ex.TargetSite.ToString() == EXCEPTION_USER_EDITING)
                {
                    MessageBox.Show(Language.Error_Excel_User_Editing, Language.Warning_title);
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

        public void TestInterCom()
        {
            MessageBox.Show("Appel inter com marche !");
        }
    }
}
