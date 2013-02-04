VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum E_varType
    T_BOOLEAN
    T_REAL
    T_DATE_AND_TIME
    T_INTEGER
    UNKNOWN
End Enum

Public typeOfVar As E_varType
Public name As String
Public path As String
Public value As Variant

Public Function getStringValue() As String
    Select Case typeOfVar
    Case T_BOOLEAN
         If varType(value) = vbInteger Or varType(value) = vbDouble Then
            If (value = 0) Then
                getStringValue = "False"
            Else
                getStringValue = "True"
            End If
        Else
            MsgBox "Variable value is not managed : " & TypeName(value)
        End If
    Case Else
    Else
        getStringValue = value
    End Select
End Function
