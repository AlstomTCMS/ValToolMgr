using System;
using System.Collections.Generic;
using System.Text;

namespace ValToolFunctions_2013
{
    class LayoutVersion_DropDown
    {
        //int ItemCount;
        //Dim ListItemsRg As Range
        public static LAYOUT SelectedLayoutVersion = LAYOUT.NONE;
        public static Array LayoutVersions;

        LAYOUT getSelectedLayoutVersion()
        {
            if (SelectedLayoutVersion == LAYOUT.NONE)
            {
                SelectedLayoutVersion = LAYOUT.L_2013; //GetValue("LayoutVersion");
                //if (SelectedLayoutVersion == "Error"){
                    //AddValue("LayoutVersion", LAYOUT_2013);
                    //SelectedLayoutVersion = GetValue("LayoutVersion");
                //}
            }
            return SelectedLayoutVersion;
        }

        //Array getLayoutVersions() {
        //    if (LayoutVersions == null){
        //        LayoutVersions = Array(LAYOUT_2012, LAYOUT_2013);
        //    }
        //    return LayoutVersions;
        //}

        //int getLayoutArrayIndex(string version){
        //    int i;
        //    int getLayoutArrayIndex = 0;
        //    For i = LBound(LayoutVersions) To UBound(LayoutVersions)
        //        If LayoutVersions(i) = version Then
        //            getLayoutArrayIndex = i
        //            Exit For
        //        End If
        //    Next i
        //    return getLayoutArrayIndex;
        //}

        ////=========Drop Down Code =========

        ////Callback for Dropdown getItemCount.
        ////Tells Excel how many items in the drop down.
        //Sub DDItemCount(control As IRibbonControl, ByRef returnedVal)
        //    Call getLayoutVersions
        //    ItemCount = UBound(LayoutVersions) - LBound(LayoutVersions) + 1
        //    returnedVal = ItemCount
        //End Sub

        ////Callback for dropdown getItemLabel.
        ////Called once for each item in drop down.
        ////If DDItemCount tells Excel there are 10 items in the drop down
        ////Excel calls this sub 10 times with an increased "index" argument each time.
        ////We use "index" to know which item to return to Excel.
        //Sub DDListItem(control As IRibbonControl, index As Integer, ByRef returnedVal)
        //    returnedVal = getLayoutVersions(index)
        //    SelectedLayoutVersion = returnedVal
        //    //index is 0-based, our list is 1-based so we add 1.
        //End Sub

        ////Drop down change handler.
        ////Called when a drop down item is selected.
        //Sub DDOnAction(control As IRibbonControl, ID As String, index As Integer)
        //' Two ways to set the variable SelectedLayoutVersion to the dropdown value

        //'way 1
        //    SelectedLayoutVersion = getLayoutVersions(index)
        //    Call UpdateValue("LayoutVersion", SelectedLayoutVersion)

        //    //way 2
        //    'Call DDListItem(control, index, SelectedLayoutVersion)

        //End Sub

        ////Returns index of item to display.
        //Sub DDItemSelectedIndex(control As IRibbonControl, ByRef returnedVal)
        //    Dim LayoutStr As String
    
        //    returnedVal = 0
        //    'control.Tag = -è
        //    LayoutStr = GetValue("LayoutVersion")
        //    If LayoutStr = "Error" Then
        //        Call AddValue("LayoutVersion", LAYOUT_2013)
        //        LayoutStr = GetValue("LayoutVersion")
        //    End If
        //    SelectedLayoutVersion = LayoutStr
        //    returnedVal = getLayoutArrayIndex(SelectedLayoutVersion)
        //End Sub

        ////------- End DD Code --------


        ////Show the variable SelectedLayoutVersion (selected item in the dropdown)
        ////You can use this variable also in other macros
        //Sub ValueSelectedItem(control As IRibbonControl)
        //    MsgBox "The variable SelectedLayoutVersion have the value = " & SelectedLayoutVersion & vbNewLine & _
        //           "You can use SelectedLayoutVersion in other code now to use the dropdown value"
        //End Sub
    }
}
