VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public title As String
Public Desc As String
Private StepList As New Collection

Private Sub Class_Initialize()
    Set StepList = New Collection
End Sub

Public Sub AddStep(o_action As CStep)
    StepList.Add o_action
End Sub

Public Sub AddReq(ByRef req As String)
    ' Not Implemented
End Sub
