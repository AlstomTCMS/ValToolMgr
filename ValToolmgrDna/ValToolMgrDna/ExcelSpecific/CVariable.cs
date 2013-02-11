using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.ExcelSpecific
{
    abstract class CVariable
    {
        public string name { get; set; }

        public string path { get; set; }

        public abstract object value { get; set; }

        public override string ToString()
        {
            return base.ToString();
            //Public Function getStringValue() As String
            //    Select Case typeOfVar
            //    Case T_BOOLEAN
            //         If varType(value) = vbInteger Or varType(value) = vbDouble Then
            //            If (value = 0) Then
            //                getStringValue = "False"
            //            Else
            //                getStringValue = "True"
            //            End If
            //        Else
            //            MsgBox "Variable value is not managed : " & TypeName(value)
            //        End If
            //    Case Else
            //        getStringValue = value
            //    End Select
            //End Function
        }
    }
}
