VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestStandInstr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum categoryList
    TS_FORCE
    TS_UNFORCE
    TS_TEST
    TS_CALL
    TS_WAIT
    UNKNOWN
End Enum

Private Const C_TS_FORCE As String = "Force"
Private Const C_TS_UNFORCE As String = "UnForce"
Private Const C_TS_TEST As String = "Test"
Private Const C_TS_CALL As String = "SequenceCall"
Private Const C_TS_WAIT As String = "Wait"
Private Const C_UNKNOWN As String = "UNKNOWN"

Public category As categoryList
Public Data As Variant

Public Function getCategoryAsText() As String
    Select Case category
          Case TS_FORCE
              getCategoryAsText = C_TS_FORCE
          Case TS_UNFORCE
              getCategoryAsText = C_TS_UNFORCE
          Case TS_TEST
              getCategoryAsText = C_TS_TEST
          Case TS_CALL
              getCategoryAsText = C_TS_CALL
          Case TS_WAIT
              getCategoryAsText = C_TS_WAIT
          Case Else
              getCategoryAsText = C_UNKNOWN
    End Select
End Function
