VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestStandSeqContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public title As String
Public Desc As String

Private SeqList As New Collection

Private Sub Class_Initialize()
    Set SeqList = New Collection
End Sub

Public Sub AddSequence(o_instr As CTestStandSeq)
    SeqList.Add o_instr
End Sub

Public Function getCount() As Integer
    getCount = SeqList.Count
End Function

Public Function getSequence(index As Integer) As CTestStandSeq
    Set getSequence = SeqList.Item(index)
End Function
