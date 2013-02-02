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

Public Enum varType
    T_BOOLEAN
    T_REAL
    t_DATE_AND_TIME
    T_INTEGER
End Enum

Public varType As varType
Public name As String
Public path As String
Public value As Variant
