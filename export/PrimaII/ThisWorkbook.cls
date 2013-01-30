VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents app As Application
Attribute app.VB_VarHelpID = -1
    
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    'MsgBox "on est content !"
    
    Call extractBAS
    
End Sub

Sub extractBAS()
Dim VBComp As VBIDE.VBComponent
    Dim Sfx As String
    
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComp.Type
        Case vbext_ct_ClassModule, vbext_ct_Document
        Sfx = ".cls"
        Case vbext_ct_MSForm
        Sfx = ".frm"
        Case vbext_ct_StdModule
        Sfx = ".bas"
        Case Else
        Sfx = ""
        End Select
        If Sfx <> "" Then
        VBComp.Export _
        Filename:=ActiveWorkbook.Path & "\export\PrimaII\" & VBComp.Name & Sfx
        End If
    Next VBComp
End Sub


