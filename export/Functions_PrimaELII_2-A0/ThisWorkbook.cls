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
    Call exportBas
End Sub

Sub exportBas()
    Dim VBComp As VBIDE.VBComponent
    Dim Sfx As String
    
    'On supprime tous les fichiers du dossier de sortie.
    On Error Resume Next
    Kill MacroPath & exportFolder & "*.*"
    On Error GoTo 0
    
    'On exporte tous les modules, les forms et cette feuille-ci
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComp.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                If Not VBComp.Name Like "Feuil*" Then
                    Sfx = ".cls"
                Else
                    Sfx = ""
                End If
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select
        
        If Sfx <> "" Then
            VBComp.Export Filename:=ActiveWorkbook.Path & exportFolder & VBComp.Name & Sfx
        End If
    Next VBComp
End Sub



Private Sub Workbook_Open()
    'Call importBas
End Sub

Sub importBas()
    Dim MyObj As Object, MySource As Object, file As Variant
    
    'Efface tous les Modules présents dans ce fichier
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.Type = vbext_ct_StdModule Then
            ThisWorkbook.VBProject.VBComponents.Remove VBComp
        End If
    Next
    
    ' Importe tous les fichiers .bas provenants de l'export à la sauvegarde
    file = Dir(MacroPath & exportFolder)
    While (file <> "")
        If InStr(file, ".bas") > 0 Then
            ThisWorkbook.VBProject.VBComponents.Import Filename:=MacroPath & exportFolder & file
        End If
        file = Dir
    Wend

End Sub


