Attribute VB_Name = "Verifications"
Private firstError As Boolean

'Rend visible l'onglet des vérifications en fonction de ses éléments
Sub getVisibleVerifTab(control As IRibbonControl, ByRef visible)
    visible = True
    'Call setVisible_VerifTestButton(control, visible)
End Sub

Sub VerificationsTestCourant(control As IRibbonControl)
    If HasActiveBook Then
        'Vérifier si la feuille courante est un test
        With ActiveSheet
            testName = .Name
            If testName Like "B?_???_???" Then
                If Not checkTest(testName, True) Then
                    Call FormatErrorSheet
                    Message = "Le test " & testName & " contient des erreurs !"
                    Call MsgBox(Message, vbExclamation, "Attention")
                End If
            Else
                MsgBox "L'onglet courant n'est pas un onglet de test PRIMA.", vbOKOnly + vbExclamation, "Fonctionnalité non utilisable !"
            End If
        End With
    End If
End Sub

Sub setVisible_VerifTestButton(control As IRibbonControl, ByRef visible)
    visible = True
    'If ActiveSheet.Name Like "B?_???_???" Then
        'visible = True
    'Else
        'visible = False
    'End If
End Sub

' Vérifie si un onglet existe pour ce test
' Vérifie l'ordre des types dans chaque étapes du test
' Vérifie s'il y a des doublons
' Renvoie vrai si tout c'est bien passé
Function checkTest(ByVal testSheet As String, Optional ByVal fromCheckButton As Boolean = False) As Boolean
Dim nNbLigne, nNbLigneT, nNbLigneTotal As Long
Dim bErreur As Boolean

Dim Etape_EnCours As range
Dim nNbLigneEtape_Deb, nNbLigneEtape_Fin As Long 'Num Ligne Debut et Fin d'Etape

Dim TexteType_EnCours As String
Dim nNbLigneType_Deb, nNbLigneType_Fin As Integer 'Num Ligne Debut et Fin de Type (AEn,CEn,ACc,CCc)
    
    checkTest = True
    firstError = True
    
    'Tester si l'onglet de test existe
    If Not WsExist(testSheet) Then
        checkTypeOrder = False
        Exit Function
    End If
       
    
    If fromCheckButton Then
        'Si la feuille d'erreur existait déjà on la supprime
        If WsExist(ERROR_NAME) Then
            Application.DisplayAlerts = False
            Sheets(ERROR_NAME).Delete
            Application.DisplayAlerts = True
        End If
        'Verifier si le test est entièrement completé
        Call verif_remplissage(testSheet, True)
        'Voir s'il faut sortir ou pas si c'est pas bon
        Call initValidationSheet
    End If
    
    Call SetCellulesVidesRouges_TEST(testSheet)
    Call SetValidations_TEST(testSheet)
    
    With Sheets(testSheet)
        
        'on créé une colonne de comparaison par concaténation de type|section|Variable|chemin
        nNbLigneTotal = .range("A1").End(xlDown).row
        For nNbLigne = 2 To nNbLigneTotal
            .Cells(nNbLigne, TEST_COLUMN_DOUBLON_COMPARE) = .Cells(nNbLigne, 7) & .Cells(nNbLigne, 8) & .Cells(nNbLigne, 9) & .Cells(nNbLigne, 10)
        Next
         
        nNbLigneEtape_Deb = 2
        nNbLigne = 2
        While (.Cells(nNbLigne, 1) <> "") And (nNbLigne <= nNbLigneTotal)
            Set Etape_EnCours = .Cells(nNbLigneEtape_Deb, 1)
            nNbLigneEtape_Deb = nNbLigne
            
            'Trouve les bornes de l'Etape
            For nNbLigne = nNbLigneEtape_Deb To nNbLigneTotal
                If .Cells(nNbLigne, 1) <> Etape_EnCours Then ' Or (.Cells(nNbLigne, 1) = "") Then
                    nNbLigneEtape_Fin = nNbLigne - 1
                    Exit For
                End If
            Next
            
            If nNbLigne >= nNbLigneTotal Then
                nNbLigneEtape_Fin = nNbLigneTotal
            End If
            nNbLigneType_Deb = nNbLigneEtape_Deb
            
        '-----------------------------------------------
        '------------------------ACc
            TexteType_EnCours = .Cells(nNbLigneType_Deb, TEST_COLUMN_TYPE_ACTION)
            If (TexteType_EnCours = TYPE_VAR_ACTION_EMB) And (bErreur = False) Then
                For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                    If (.Cells(nNbLigne, TEST_COLUMN_TYPE_ACTION) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                        nNbLigneType_Fin = nNbLigne - 1
                        Exit For
                    End If
                Next
                bErreur = Doublon(testSheet, nNbLigneEtape_Deb, nNbLigneEtape_Fin)
                nNbLigneType_Deb = nNbLigneType_Fin + 1
            End If
        '------------------------AEn
            TexteType_EnCours = .Cells(nNbLigneType_Deb, TEST_COLUMN_TYPE_ACTION)
            If (TexteType_EnCours = TYPE_VAR_ACTION_ENV) And (bErreur = False) Then
                For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                    If (.Cells(nNbLigne, TEST_COLUMN_TYPE_ACTION) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                        nNbLigneType_Fin = nNbLigne - 1
                        Exit For
                    End If
                Next
                bErreur = Doublon(testSheet, nNbLigneEtape_Deb, nNbLigneEtape_Fin)
                nNbLigneType_Deb = nNbLigneType_Fin + 1
            End If
        '------------------------CCc
            TexteType_EnCours = .Cells(nNbLigneType_Deb, TEST_COLUMN_TYPE_ACTION)
            If (TexteType_EnCours = TYPE_VAR_CHECK_EMB) And (bErreur = False) Then
                For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                    If (.Cells(nNbLigne, TEST_COLUMN_TYPE_ACTION) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                        nNbLigneType_Fin = nNbLigne - 1
                        Exit For
                    End If
                Next
                bErreur = Doublon(testSheet, nNbLigneEtape_Deb, nNbLigneEtape_Fin)
                nNbLigneType_Deb = nNbLigneType_Fin + 1
            End If
        '------------------------CEn
            TexteType_EnCours = .Cells(nNbLigneType_Deb, TEST_COLUMN_TYPE_ACTION)
            If (TexteType_EnCours = TYPE_VAR_CHECK_ENV) And (bErreur = False) Then
                For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                    If (.Cells(nNbLigne, TEST_COLUMN_TYPE_ACTION) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                        nNbLigneType_Fin = nNbLigne - 1
                        Exit For
                    End If
                Next
                bErreur = Doublon(testSheet, nNbLigneEtape_Deb, nNbLigneEtape_Fin)
                nNbLigneType_Deb = nNbLigneType_Fin + 1
            End If
        '------------------------PGM
            TexteType_EnCours = .Cells(nNbLigneType_Deb, TEST_COLUMN_TYPE_ACTION)
            If (TexteType_EnCours = TYPE_VAR_PGM) And (bErreur = False) Then
                For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                    If (.Cells(nNbLigne, TEST_COLUMN_TYPE_ACTION) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                        nNbLigneType_Fin = nNbLigne - 1
                        Exit For
                    End If
                Next
                bErreur = False  'Doublon(testSheet, nNbLigneType_Deb, nNbLigneType_Fin)
                nNbLigneType_Deb = nNbLigneType_Fin + 1
            End If
        '-----------------------------------------------
        
    
            nNbLigne = nNbLigneEtape_Fin + 1
            nNbLigneEtape_Deb = nNbLigneEtape_Fin + 1
            
            If bErreur = True Then
                nNbLigne = nNbLigneTotal + 1
                'MsgBox "Doublon dans l'Etape " & Etape_EnCours, vbOKOnly + vbCritical, "Erreur doublon !"
                'Call AjoutErreur(testSheet, Etape_EnCours, StringFormat(ERROR_TYPE_DOUBLON, "TES3"))
                If checkTest Then
                    checkTest = False
                End If
            End If
            If (nNbLigneType_Deb <= nNbLigneEtape_Fin) And (bErreur = False) Then
                bErreur = True
                nNbLigne = nNbLigneTotal + 1
                'MsgBox "L'ordre des types de variables (ACc, AEn, CCc, CEn) n'a pas été respecté pour l'étape : " & Etape_EnCours, vbOKOnly + vbCritical, "Erreur !"
                Call AjoutErreur(testSheet, Etape_EnCours, ERROR_TYPE_ORDER)
                If checkTest Then
                    checkTest = False
                End If
            End If
        Wend
        
        ' On efface la colonne qui a servit de tampon pour vérifier les doublons
        .Columns(TEST_COLUMN_DOUBLON_COMPARE).EntireColumn.Delete
        
    End With
End Function

'Si la fiche de Doublons n'existe pas, la créer
Sub initErrorSheet()

    If firstError Then
        firstError = False
        
        If Not WsExist(ERROR_NAME) Then
            Call InitSheet(ERROR_NAME, True, , , Array("Test", "Etape", "Erreur"))
            
            'On colorie l'onglet en rouge
            With Sheets(ERROR_NAME)
                .ListObjects("TableauErreurs").TableStyle = "TableStyleMedium3"
                .Move After:=Sheets("Conf Banc")
                With .Tab
                    .Color = 255
                    .TintAndShade = 0
                End With
            End With
        End If
    End If
End Sub


' Vérifie si c'est en doublon dans l'étape courante (entre début et fin)
Private Function Doublon(ByVal testSheetName As String, ByVal Deb As Integer, ByVal Fin As Integer) As Boolean
Dim Resultat As Boolean
Dim ligne1, ligne2 As Integer
Dim a, b
Doublon = False

    With Sheets(testSheetName)
    For ligne1 = Deb To Fin
        b = .Cells(ligne1, TEST_COLUMN_DOUBLON_COMPARE)
        a = Application.WorksheetFunction.CountIf(.range(.Cells(Deb, TEST_COLUMN_DOUBLON_COMPARE), .Cells(Fin, TEST_COLUMN_DOUBLON_COMPARE)), b)
        ' S'il y a plus qu'une occurence, alors on a un doublon
        If a > 1 Then
            Doublon = True
            Call AjoutErreur(testSheetName, .Cells(ligne1, 1), StringFormat(ERROR_TYPE_DOUBLON, .Cells(ligne1, 9) & " " & .Cells(ligne1, 10))) 'Var et chemin
        End If
    Next
    End With

End Function

' Ajoute dans la fiche de Doublons les doublons qu'il y a eu.
Private Sub AjoutErreur(testSheetName As String, etapeRange As range, errorMsg As String)
Dim nouvelleLigne As range
    
    Call initErrorSheet
    
    With Sheets(ERROR_NAME)
        'Déterminer la dernière ligne
        If .range("A1").End(xlDown) = "" Then
            Set nouvelleLigne = .range("A1").End(xlDown).range("A1:C1")
        Else
            Set nouvelleLigne = .range("A1").End(xlDown).range("A2:C2")
        End If
        'ajout de l'étape
        
        'Si l'erreur est pour un test et pas une étape
        If etapeRange Is Nothing Then
            nouvelleLigne = Array(testSheetName, "", errorMsg)
            .Hyperlinks.Add Anchor:=nouvelleLigne.range("A1"), Address:="", _
                SubAddress:="'" & testSheetName & "'!A1", TextToDisplay:=testSheetName
        Else
            nouvelleLigne = Array(testSheetName, etapeRange.Value, errorMsg)
            linkAddress = "'" & testSheetName & "'!" & Replace(etapeRange.Address, "$", "")
            .Hyperlinks.Add Anchor:=nouvelleLigne.range("B1"), Address:="", _
                SubAddress:=linkAddress, TextToDisplay:=etapeRange.Value
            .Hyperlinks.Add Anchor:=nouvelleLigne.range("A1"), Address:="", _
                SubAddress:="'" & testSheetName & "'!A1", TextToDisplay:=testSheetName
        End If
    End With
End Sub



' Vérifie que les colonnes soient remplies jusqu'au bout
' Renvoie Vrai si toutes les colonnes sont bien remplies
Function verif_remplissage(sheetName As String, isTestSheet As Boolean) As Boolean
Dim i, j, columnIndex As Integer
Dim ligneFin, finFichier As Long
Dim errorColumns As String
verif_remplissage = True
    
    With Sheets(sheetName)
        If Not isTestSheet Then
            ' Vérification de la colonne Num_Etape
            ligneFin = .Cells(.Rows.Count, 6).End(xlUp).row
            If .Cells(ligneFin + 1, 1) <> "END" Then
                MsgBox "La colonne Num_Etape de la feuille " & sheetName & " s'est arrêtée à la ligneFin " & Str(ligneFin)
                verif_remplissage = False
            End If
            
            ' On vérifie les numéros de section
            nNbLigneTotal = nNbLigne
            For nNbLigne = 9 To nNbLigneTotal
                If StrComp(Left(.Cells(nNbLigne, 12), 1), "M", 1) = 0 Then 'chaine du style Menante ou Menee
                ElseIf StrComp(.Cells(nNbLigne, 12), "X", 1) = 0 Then 'chaine du style X
                ElseIf .Cells(nNbLigne, 12) <= 7 Then 'chaine numérique
                Else
                    MsgBox "La colonne Section de la feuille " & sheetName & " s'est arrêtée à la ligne " & Str(nNbLigne)
                    bErreur = True
                End If
            Next
        
            'Verif des autres colonnes
            For i = 11 To 15
                ligneFin = .Cells(.Rows.Count, i).End(xlUp).row
                If .Cells(ligneFin + 1, 1) <> "END" Then
                    MsgBox "La colonne " & .Cells(8, i) & " de la feuille " & sheetName & " s'est arrêtée à la ligneFin " & Str(ligneFin)
                    verif_remplissage = False
                End If
            Next i
        Else
            If .range("B1") = "Mode" Then
                columnIndex = 7
            Else
                columnIndex = 6
            End If
            
            finFichier = .Cells(.Rows.Count, 1).End(xlUp).row
            
            'Verif des autres colonnes
            For i = columnIndex To columnIndex + 4
                ligneFin = finFichier
                While ligneFin > 0 And Len(.Cells(ligneFin, i)) = 0
                    ligneFin = ligneFin - 1
                Wend
                
                If ligneFin < finFichier Then
                    If verif_remplissage Then
                        verif_remplissage = False
                    End If
                    errorColumns = errorColumns & ", " & .Cells(1, i)
                End If
            Next i
            
            If Not verif_remplissage Then
                
                'MsgBox "La feuille de test " & sheetName & " n'est pas entièrement remplie." _
                   '& vbCrLf & "Le PR ne peut pas être généré.", vbExclamation, "Alerte !"
                Call AjoutErreur(sheetName, Nothing, StringFormat(ERROR_TYPE_EMPTY, Right(errorColumns, Len(errorColumns) - 2)))
            End If
        End If
    End With

End Function


Sub testChckEmpty()
Dim emptyList As Variant

    emptyList = checkEmpty(ActiveSheet.range("G2:K126"))
    
    If UBound(emptyList) = -1 Then
        trs = "pas de cellule vide"
    End If
End Sub

' Donne la table des cellules vides d'une zone
Function checkEmpty(range2Check) As Variant
    Dim emptyCellsStr As String
    Dim cell As range


    For Each cell In range2Check
        If cell.Value = vbNullString Then
            If emptyCellsStr = "" Then
                emptyCellsStr = Replace(cell.Address, "$", "")
            Else
                emptyCellsStr = emptyCellsStr & ";" & Replace(cell.Address, "$", "")
            End If
        End If
    Next cell
    
    checkEmpty = Split(emptyCellsStr, ";")
End Function


'Ajoute la formatage conditionnel sur cellules vides
Public Sub SetCellulesVidesRouges_TEST(ByVal sheetName As String)
Dim formatAlreadyExist As Boolean

    With Sheets(sheetName)
        Fin = .range("A1").End(xlDown).row
        With .range("G2:K" & Fin)
    
            'Voir si le formatage existe deja
            For Each formatCond In .FormatConditions
                If formatCond.Formula1 = "=NBCAR(SUPPRESPACE(G2))=0" Then
                    formatAlreadyExist = True
                End If
            Next formatCond
            
            If Not formatAlreadyExist Then
                .FormatConditions.Add Type:=xlExpression, Formula1:="=NBCAR(SUPPRESPACE(G2))=0"
                
                With .FormatConditions(.FormatConditions.Count)
                    .SetFirstPriority
                    With .Interior
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                    End With
                    .StopIfTrue = True
                End With
            End If
            
        End With
    End With
End Sub

' Ajoute les validations automatiques des données de tests
Sub SetValidations_TEST(ByVal sheetName As String)

    With Sheets(sheetName)
        Fin = .range("A1").End(xlDown).row
        
        'Types de Véhicules permis
        With .range("H2:H" & Fin).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='" & VALID_NAME & "'!$C$2:$C$8"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
        'Types de variable permis
        With .range("G2:G" & Fin).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='" & VALID_NAME & "'!$B$2:$B$6"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End With
End Sub

' Ajoute les validations automatiques des données de tests
Sub SetValidations_SYNTH()

    With Sheets(SYNTHESE_NAME)
        Fin = .range("F1").End(xlDown).row
        
        'Types de Conf banc permis
        With .range("B2:B" & Fin).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="='" & VALID_NAME & "'!$A$2:$A$6"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End With
End Sub

Public Sub FormatErrorSheet()
    With Sheets(ERROR_NAME)
        .Columns("A:C").AutoFit
        .UsedRange.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
        .Activate
    End With
End Sub

Sub test()
    Call SetCellulesVidesRouges_TEST(ActiveSheet.Name)
    Call SetValidations_TEST(ActiveSheet.Name)
End Sub

'Initie la feuille de validation si elle n'existe pas
Public Sub initValidationSheet()
    If Not WsExist(VALID_NAME) Then
        
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = VALID_NAME
        With Sheets(VALID_NAME)
            .range("A1:C1") = Array("Conf banc", "Type Var", "Section")
            
            .ListObjects.Add(xlSrcRange, .range("$A$1:$C$1"), , xlYes).Name = "TableauDataValidation"
            .ListObjects("TableauDataValidation").TableStyle = "TableStyleMedium11"
            
            .range("A2:A6") = Application.transpose(Array("A1", "A2", "B", "C", "D"))
            .range("B2:B6") = Application.transpose(Array("AEn", "ACc", "CEn", "CCc", "PGM"))
            .range("C2:C8") = Application.transpose(Array("Menante", "Menee", "X", "1", "2", "3", "4"))
            
            .visible = False
        End With
    End If
End Sub
