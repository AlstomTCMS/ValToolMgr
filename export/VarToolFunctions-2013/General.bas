Attribute VB_Name = "General"

'------------------------------------------------------------------------
' Initie une feuille par son nom
'------------------------------------------------------------------------
Function InitSheet(ByVal sheetName As String, Optional ByVal eraseContent As Boolean, Optional visible As Boolean = True, Optional sheetAlreadyExist As Boolean, Optional titles As Variant = Null) As Excel.Worksheet
    Dim WsExist As Boolean, range1 As Range
 
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
        
            Set range1 = .Range("A1", .Cells(1, UBound(titles) + 1))
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



'Fonction à appeler depuis toute macro appelée par un bouton de barre de macro externe
'return vrai si il y a un fichier d'ouvert
Function HasActiveBook(Optional ByVal displayMsg As Boolean = True) As Boolean

    HasActiveBook = True
    On Error GoTo NoActiveWorkBook:
    'Si on a un nouveau classeur vide
    'If ActiveWorkbook.Name Like "Classeur*" Or ActiveWorkbook.Name Like "Book*" Then
        'GoTo NoActiveWorkBook
    'End If
    Set wb = Workbooks(1)
    On Error GoTo 0
    Exit Function
    
NoActiveWorkBook:
    HasActiveBook = False
    If displayMsg Then
        Call MsgBox("Please open a PR file to use this feature.", vbExclamation, "Alerte")
    End If
End Function


'Réecri un String avec des parametres entre crochet {} remplacés par la liste de paramètres mis en argument
Public Function StringFormat(ByVal forFormat As String, ParamArray params() As Variant) As String
Dim i As Integer
Dim formatted As String

    formatted = forFormat
    For i = LBound(params()) To UBound(params())
        formatted = Replace(formatted, "{" & CStr(i) & "}", CStr(params(i)))
    Next
    StringFormat = formatted
End Function

