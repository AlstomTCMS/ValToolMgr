using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.ExcelSpecific
{
    class CVariable
    {
        public enum E_varType
        {
            T_BOOLEAN,
            T_REAL,
            T_DATE_AND_TIME,
            T_INTEGER,
            UNKNOWN
        }

        public string name { get; set; }

        public string path { get; set; }

        public object value { get; set; }

        public E_varType typeOfVar { get; set; }

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
