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
Private actionList As New Collection
Private checkList As New Collection

Private Sub Class_Initialize()
    Set actionList = New Collection
    Set checkList = New Collection
End Sub

Public Sub AddAction(o_action As CInstruction)
    actionList.Add o_action
End Sub

Public Function getActionCount() As Integer
    getActionCount = actionList.Count
End Function

Public Function getAction(index As Integer) As CInstruction
    Set getAction = actionList.Item(index)
End Function

Public Sub AddCheck(o_check As CInstruction)
    checkList.Add o_check
End Sub

Public Function getcheckCount() As Integer
    getcheckCount = checkList.Count
End Function

Public Function getCheck(index As Integer) As CInstruction
    Set getCheck = checkList.Item(index)
End Function

Public Function getTotalCount() As Integer
    getTotalCount = getcheckCount() + getActionCount()
End Function

Public Sub AddReq(ByRef req As String)
    ' Not Implemented
End Sub
