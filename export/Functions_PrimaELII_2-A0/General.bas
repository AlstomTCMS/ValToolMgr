Attribute VB_Name = "General"
'------------------------------------------------------------------------
' Initie une feuille par son nom
'------------------------------------------------------------------------
Function InitSheet(ByVal sheetName As String, Optional ByVal eraseContent As Boolean, Optional visible As Boolean = True, Optional sheetAlreadyExist As Boolean, Optional titles As Variant = Null) As Excel.Worksheet
    Dim WsExist As Boolean, range1 As range
 
On Error Resume Next
    WsExist = ActiveWorkbook.Sheets(sheetName).Index
On Error GoTo 0

    'Si la feuille n'existe pas, on l'ajoute
    If Not WsExist Then
        Worksheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
    Else
        sheetAlreadyExist = True
    End If
    
    On Error Resume Next
    ' On efface le contenu de la feuille
    'Sheets(sheetName).Cells.ClearContents
    If eraseContent Then
        Sheets(sheetName).Cells.ClearContents
        Sheets(sheetName).Cells.ClearContents
    End If
    
    'On ajoute les titres s'il y en a
    If Not titles Is Null Then
        With Sheets(sheetName)
        
            Set range1 = .range("A1", .Cells(1, UBound(titles) + 1))
            range1 = titles
            tableLiens = "Tableau" & sheetName
            .ListObjects.Add(xlSrcRange, range1, , xlYes).Name = tableLiens
            .ListObjects(tableLiens).TableStyle = "tableau de test"
            
            'enlève l'affichage grille
            .Activate
            ActiveWindow.DisplayGridlines = False
            
        End With
    End If
    On Error GoTo 0
           
    If Not visible Then
        Sheets(sheetName).visible = xlSheetHidden
    Else
        Sheets(sheetName).visible = xlSheetVisible
    End If
    
    'feuille renvoyée
    Set InitSheet = Sheets(sheetName)
    
End Function

'------------------------------------------------------------------------
' Dit si une feuille existe dans le fichier
'------------------------------------------------------------------------
Function WsExist(ByVal Nom$) As Boolean
'Nous dit si la feuille mis en paramètre existe
    On Error Resume Next
    WsExist = ActiveWorkbook.Sheets(Nom).Index
    On Error GoTo 0
End Function


'------------------------------------------------------------------------
' Copie les feuilles de références de C:\macros_alstom\Ref_PrimaELII_2-XX.xlsx dans le fichier PR ouvert
' Renvoie True si la copie s'est bien passée
' Macro crée le 28/11/2012 par dLeonardi
'------------------------------------------------------------------------
Function CopySheetsFromRef() As Boolean
Dim sheetName, thisFileName, RefFileName As String
Dim sheet As Worksheet
Dim has2Update, exist, hasCopyARefSheet, XlsConvert As Boolean
    
    CopySheetsFromRef = True
    hasCopyARefSheet = False
    XlsConvert = False

    thisFileName = ActiveWorkbook.Name
    workingDirectory = ActiveWorkbook.Path
    RefFileName = TEMPLATE_FULLPATH

    On Error GoTo ErrHandler:
    Workbooks.Open RefFileName, 0, ReadOnly:=False
    On Error GoTo 0
    
    RefWB = ActiveWorkbook.Name
    Workbooks(thisFileName).Activate
    'Convert file to Excel 2007 xlsx if not
    If thisFileName Like "*.xls" Then
        XlsConvert = True
    End If
    If XlsConvert Then
        fileSaveName = workingDirectory & "\" & Replace(thisFileName, ".xls", ".xlsx")
        Application.DisplayAlerts = False
        Workbooks(thisFileName).SaveAs Filename:=fileSaveName, FileFormat:=XlFileFormat.xlOpenXMLWorkbook, Addtomru:=True
        Application.DisplayAlerts = True
        thisFileName = Replace(thisFileName, ".xls", ".xlsx")
        Workbooks(thisFileName).Close 'Attention, le nom est changé automatiquement
        On Error GoTo ErrHandler:
        Workbooks.Open fileSaveName, 0, ReadOnly:=False
        On Error GoTo 0
    End If
        
    Call Add_AddInRef_to_WorkBook(Workbooks(thisFileName))
    
    indiceToCopy = 1
    If WsExist(2) Then
        indiceToCopy = 2
    End If
    
    
    For Each sheet In Workbooks(RefWB).Sheets
        exist = WsExist(sheet.Name)
        'Supprimer la feuille si elle doit être mise à jour
        If exist Then
            
            If has2Update Then
                Application.DisplayAlerts = False
                Workbooks(thisFileName).Sheets(sheet.Name).Delete
                Application.DisplayAlerts = True
                sheet.Copy After:=Workbooks(thisFileName).Sheets(indiceToCopy)
                hasCopyARefSheet = True
            End If
        ElseIf Not exist Then
            sheet.Copy After:=Workbooks(thisFileName).Sheets(indiceToCopy)
            hasCopyARefSheet = True
        End If
    Next sheet
    
    Application.DisplayAlerts = False
    Workbooks(RefWB).Close
    Application.DisplayAlerts = True

    'Remove links to ref file in named ranges
    If hasCopyARefSheet Then
        'If XlsConvert Then
            'Workbooks(thisFileName).ChangeLink Name:=RefWB, NewName:=fileSaveName, Type:=xlExcelLinks
        'Else
            Sheets("Endpaper_PR").Unprotect
            Sheets("Endpaper_PV").Unprotect
            Workbooks(thisFileName).ChangeLink Name:=RefWB, NewName:=thisFileName, Type:=xlExcelLinks
            Sheets("Endpaper_PR").Protect
            Sheets("Endpaper_PV").Protect
        'End
        Call RemoveDuplicateNamesRef(Workbooks(thisFileName))
    End If
            
ErrHandler:
    ' error handling code
    If Err.Number = 1004 Then
        Call MsgBox("Le fichier " & RefFileName & " est introuvable." _
                & vbCrLf & "Le processus ne peut continuer. ", vbExclamation, "Alerte")
        CopySheetsFromRef = False
    End If
    
End Function


Sub UpdateNamesRef()
    Dim name_ As Name
    
    For Each name_ In ActiveWorkbook.Names
        If name_.RefersTo Like "='C:*" Or name_.RefersTo Like "='D:*" _
                Or name_.RefersTo Like "=?REF!*" Then
            name_.Delete
        ElseIf name_.RefersTo Like "='" & TEMPLATE_FILE_PREFIX & "*" Then
            name_.RefersTo = "=" & Split(name_.RefersTo, "!")(1)
        ElseIf name_.RefersTo Like "='?" & TEMPLATE_FILE_PREFIX & "*" Then
            newRef = "='" & Split(name_.RefersTo, "]")(1)
            If ActiveWorkbook.Names(name_.Name).RefersTo = name_.RefersTo Then
                name_.RefersTo = newRef
            Else
                name_.Delete
            End If
        End If
    Next
End Sub

Sub RemoveDuplicateNamesRef(wb As Workbook)
    Dim name_ As Name
    Dim nm As Name
    
    For Each name_ In wb.Names
        If name_.Name Like "*!*" And Not (name_.Name Like "*Print_Area" Or name_.Name Like "*Zone_d_impression") Then   'Zone_d_impression
            On Error Resume Next
            Set nm = wb.Names(Split(name_.Name, "!")(1))
            On Error GoTo 0
            If Not nm Is Nothing Then
                If Not nm.Name = name_.Name Then
                    name_.Delete
                End If
            Else
                wb.Names.Add Name:=Split(name_.Name, "!")(1), RefersTo:=name_.RefersTo
                name_.Delete
            End If
            Set nm = Nothing
        End If
    Next
End Sub

' Vire les feuilles inutiles qui se trouvent dans les fichiers PR d'origine
Sub virerFeuillesInutiles()
Dim sheets2Delete As Variant

sheets2Delete = Array("feuil2", "feuil3", "ACU", "TCU", "BCU", "BT", "DESK1")
utilExist = False

'Teste s'il existe des feuilles inutiles pour ne pas proposer ce message si pas nécessaire
For i = 0 To UBound(sheets2Delete)
    If WsExist(sheets2Delete(i)) Then
        utilExist = True
        Exit For
    End If
Next

If utilExist Then
    If vbYes = MsgBox("Voulez vous supprimer les feuilles inutiles (feuil2, feuil3, ACU, TCU, BCU,BT,DESK1) ?", vbExclamation + vbYesNo, "") Then
        
        
        Application.DisplayAlerts = False
        
        For i = 0 To UBound(sheets2Delete)
            If WsExist(sheets2Delete(i)) Then
                Sheets(sheets2Delete(i)).Delete
            End If
        Next
        
        Application.DisplayAlerts = True
    End If
End If
End Sub

'Fonction à appeler depuis toute macro appelée par un bouton de barre de macro externe
'return vrai si il y a un fichier d'ouvert
Function HasActiveBook() As Boolean

    HasActiveBook = True
    On Error GoTo NoActiveWorkBook:
    'Si on a un nouveau classeur vide
    If ActiveWorkbook.Name Like "Classeur*" Then
        GoTo NoActiveWorkBook
    End If
    On Error GoTo 0
    Exit Function
    
NoActiveWorkBook:
    HasActiveBook = False
    Call MsgBox("Veuillez ouvrir un fichier PR pour utiliser cette fonctionnalité.", vbExclamation, "Alerte")
End Function



Public Function StringFormat(ByVal forFormat As String, ParamArray params() As Variant) As String
Dim i As Integer
Dim formatted As String

    formatted = forFormat
    For i = LBound(params()) To UBound(params())
        formatted = Replace(formatted, "{" & CStr(i) & "}", CStr(params(i)))
    Next
    StringFormat = formatted
End Function

