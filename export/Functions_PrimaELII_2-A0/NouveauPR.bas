Attribute VB_Name = "NouveauPR"
'Permet de créer un PR à blanc

Sub AddNewPR(control As IRibbonControl)
    Call CopyRef
End Sub

Sub AddNewPR_FromSheet()
    Call CopyRef
End Sub

'------------------------------------------------------------------------
' Copie le fichier de références de C:\macros_alstom\Ref_PrimaELII_2-XX.xlsx
' Renvoie True si la copie s'est bien passée
' Macro crée le 11/01/2013 par dLeonardi
'------------------------------------------------------------------------
Public Function CopyRef() As Boolean
Dim fileSaveFullName, fileSaveName, RefFileName As String
Dim splitName As Variant
    
    CopyRef = True
    
    RefFileName = MacroPath & "\Ref_PrimaELII_2-" & refVersion & ".xltm"
    
    'Demander à l'utilisateur le nom qu'il veut mettre
    fileSaveFullName = Application.GetSaveAsFilename(InitialFileName:="", fileFilter:="Excel Files macro enabled (*.xlsm), *.xlsm")
    
    If fileSaveFullName <> False Then
        Application.ScreenUpdating = False
        'Call FileCopy(RefFileName, fileSaveFullName)
        Set newWB = Workbooks.Add(Template:=RefFileName)
        splitName = Split(fileSaveFullName, "\")
        fileSaveName = Replace(splitName(UBound(splitName)), ".xlsm", "")
        
        'faire une copie de Synthèse vierge
        Sheets(SYNTHESE_MODEL_NAME).visible = xlSheetVisible
        Sheets(SYNTHESE_MODEL_NAME).Copy After:=Sheets(Sheets.Count)
        ActiveSheet.Name = SYNTHESE_NAME
        Sheets(SYNTHESE_MODEL_NAME).visible = xlSheetHidden
        Sheets(ENDPAPER_PR_NAME).Activate
        
        newWB.SaveAs Filename:=fileSaveName, FileFormat:=XlFileFormat.xlOpenXMLWorkbookMacroEnabled
        
        Application.ScreenUpdating = True
    End If
            
ErrHandler:
    ' error handling code
    If Err.Number = 1004 Then
        Call MsgBox("Le fichier " & RefFileName & " est introuvable." _
                & vbCrLf & "Le processus ne peut continuer. ", vbExclamation, "Alerte")
        CopyRef = False
    End If
    
End Function

