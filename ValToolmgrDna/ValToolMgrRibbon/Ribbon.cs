using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices; //needed for <ComVisible(True)> 
using ExcelDna.Integration.Extensibility;

namespace ValToolMgrRibbon
{
        //class for the Ribbon handler.  
        //The handler for the button press is defined in the ribbon xml as onAction="RunButtonID"
        //'which we find in the code module inside the RiskGen.xlam file. 
        //'The ExcelRibbon-derived class must also be marked as ComVisible(True),
        //' or in the project properties, advanced options, the ComVisible option must be checked.
        //' (Note that this is not the ‘Register for COM Interop’ option, which must never be used with Excel-DNA.)
        [ComVisible(true)]
        public class Ribbon : ExcelRibbon
        {

            //ExcelDna provides a feature to use onAction="RunTagMacro" which will run a VBA void named in tag="MyVBAMacro"
            //I have used the ID for that purpose

            void AddNewPR(IRibbonControl ctl)
            {
                //‘example: id=”MyMacro” onaction=”RunButtonID”
                //Application.Run(ctl.Id);

            }

            void RunButtonIDWithTag(IRibbonControl ctl)
            {
                //‘ example: id=”TestTag” onaction=”RunButtonIDWithTag” tag=”Hello”
                //Application.Run(ctl.Id, ctl.Tag);
            }

            void Ancien_Vers_Nouveau(IRibbonControl ctl)
            {
                //If (HasActiveBook)
                //{
                Constants.LAYOUT getSelectedLayoutVersion = Constants.LAYOUT.L_2013;

                switch (getSelectedLayoutVersion)
                {                
                    case Constants.LAYOUT.L_2012:
                        //ValToolFunctions_2012.AncienVersNouveau
                        break;
                    case Constants.LAYOUT.L_2013:
                        //ValToolFunctions_2013.Ancien_Vers_Nouveau
                        break;
                    default:
                        break;
                }
                string str = Constants.LAYOUT.L_2012.ToString();
            }

            // Génère les onglets de test à partir de la synthèse
            //void Generer_Onglets_Tests(IRibbonControl ctl){
            //    If HasActiveBook Then
            //    {
            //        Select Case getSelectedLayoutVersion
            //            Case LAYOUT_2012
            //                ValToolFunctions_2012.Generer_OngletsTests
            //            Case LAYOUT_2013
            //                ValToolFunctions_2013.Generer_OngletsTests
            //        End Select
            //    }
            //}

            //void Reverse_Nvo_Vers_Ancien(IRibbonControl ctl){
            //    If HasActiveBook Then
            //        Select Case getSelectedLayoutVersion
            //            Case LAYOUT_2012
            //                ValToolFunctions_2012.Reverse_NvoVersAncien
            //            Case LAYOUT_2013
            //                ValToolFunctions_2013.Reverse_Nvo_Vers_Ancien
            //        End Select
            //}

            //void AddNewPR(IRibbonControl ctl){
            //    Select Case getSelectedLayoutVersion
            //        Case LAYOUT_2012
            //            ValToolFunctions_2012.CopyRef
            //        Case LAYOUT_2013
            //            ValToolFunctions_2013.NewPR
            //    End Select
            //    }

            //void AddNewStep(IRibbonControl ctl){
            //    Select Case getSelectedLayoutVersion
            //        Case LAYOUT_2012
            //            'ValToolFunctions_2012.CopyRef
            //            MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
            //        Case LAYOUT_2013
            //            ValToolFunctions_2013.AddNewStep
            //    End Select
            //}
        }

        //Public Module testRibbon
        //    void TestTag(ByVal sTag As String)
        //        MsgBox(sTag)
        //    }
        //End Module
}
