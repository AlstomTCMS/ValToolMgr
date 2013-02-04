VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public title As String
Public Desc As String
Private TestList As New Collection

Private Sub Class_Initialize()
    Set TestList = New Collection
End Sub

Public Sub AddTest(o_test As CTest)
    TestList.Add o_test
    TestList.Count
    
End Sub

Public Function getCount() As Integer
    getCount = TestList.Count
End Function

Public Function getTest(index As Integer) As CTest
    Set getTest = TestList.Item(index)
End Function
