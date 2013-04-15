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
    
    RefFileName = "C:\macros_alstom\Ref_PrimaELII_2-" & refVersion & ".xls"
    
    'ChDir "c:\Documents and Settings\"
    
    'Demander à l'utilisateur le nom qu'il veut mettre
    fileSaveFullName = Application.GetSaveAsFilename(InitialFileName:="B2_XXX_Y_A0", _
    fileFilter:="xls Files (*.xls), *.xls")
    
    If fileSaveFullName <> False Then
        Call FileCopy(RefFileName, fileSaveFullName)
        
        On Error GoTo ErrHandler:
        Workbooks.Open fileSaveFullName, 0, ReadOnly:=False
        On Error GoTo 0
        
        'dejà mettre le nom de la fonction dans la page de garde
        splitName = Split(fileSaveFullName, "\")
        fileSaveName = Replace(splitName(UBound(splitName)), ".xls", "")
        range(Names("Num_PR")) = Left(fileSaveName, 8)
        range(Names("Indice_PR")) = Mid(fileSaveName, 10, 2)
        
        With Sheets("Suivi Versions")
            .range("A2") = range(Names("Indice_PR"))
            .range("B2") = Date
            .range("C2") = Environ("username")
        End With
        
        'faire une copie de Synthèse vierge
        Sheets(SYNTHESE_MODEL_NAME).visible = xlSheetVisible
    
        Sheets(SYNTHESE_MODEL_NAME).Copy After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = SYNTHESE_NAME
        
        Sheets(SYNTHESE_MODEL_NAME).visible = xlSheetHidden
        
    End If
    
            
ErrHandler:
    ' error handling code
    If Err.Number = 1004 Then
        Call MsgBox("Le fichier " & RefFileName & " est introuvable." _
                & vbCrLf & "Le processus ne peut continuer. ", vbExclamation, "Alerte")
        CopyRef = False
    End If
    
End Function

