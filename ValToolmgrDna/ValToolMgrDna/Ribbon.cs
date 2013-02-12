using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices; //needed for <ComVisible(True)> 
using ExcelDna.Integration.Extensibility;

namespace ValToolMgrDna
{
    
    //class for the Ribbon handler.  
    //The handler for the button press is defined in the ribbon xml as onAction="RunButtonID"
     //'which we find in the code module inside the RiskGen.xlam file. 
    //'The ExcelRibbon-derived class must also be marked as ComVisible(True),
    //' or in the project properties, advanced options, the ComVisible option must be checked.
    //' (Note that this is not the ‘Register for COM Interop’ option, which must never be used with Excel-DNA.)
    //<ComVisible(True)>
    public class Ribbon : ExcelRibbon{
    
        //ExcelDna provides a feature to use onAction="RunTagMacro" which will run a VBA sub named in tag="MyVBAMacro"
        //I have used the ID for that purpose

        void AddNewPR(IRibbonControl ctl)
        {
	    //‘example: id=”MyMacro” onaction=”RunButtonID”
            //Application.Run(ctl.Id);
            
        }

        void RunButtonIDWithTag(IRibbonControl ctl  ){
	    //‘ example: id=”TestTag” onaction=”RunButtonIDWithTag” tag=”Hello”
            //Application.Run(ctl.Id, ctl.Tag);
        }
    }

//Public Module testRibbon
//    Sub TestTag(ByVal sTag As String)
//        MsgBox(sTag)
//    End Sub
//End Module
}





