VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestStandSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public title As String
Public Desc As String
Public identifier As String
Private InstrList As New Collection

Private Sub Class_Initialize()
    Set InstrList = New Collection
End Sub

Public Sub AddInstr(o_instr As CTestStandInstr)
    InstrList.Add o_instr
End Sub

