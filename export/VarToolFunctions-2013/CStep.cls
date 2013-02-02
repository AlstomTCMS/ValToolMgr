VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public title As String
Public DescAction As String
Public DescCheck As String
Private InstrList As New Collection

Private Sub Class_Initialize()
    Set InstrList = New Collection
End Sub

Public Sub AddInstruction(o_action As CInstruction)
    InstrList.Add o_action
End Sub

Public Sub AddReq(ByRef req As String)
    ' Not Implemented
End Sub
