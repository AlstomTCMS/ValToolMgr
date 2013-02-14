Attribute VB_Name = "VersAncienFormalisme"

Sub Reverse_Nvo_Vers_Ancien(control As IRibbonControl)
    If HasActiveBook Then
        Call Reverse_NvoVersAncien
    End If
End Sub

'Function Reverse
Public Sub Reverse_NvoVersAncien()
Dim ws As Worksheet
    Application.ScreenUpdating = False
    
    'On test si la fiche de synthèse et les onglets de tests existent sinon on fait pas le test
    If Not WsExist(SYNTHESE_NAME) Then
        Call MsgBox("La fiche de Synthèse n'existe pas !" _
            & vbCrLf & "Il est impossible de générer un PR sans cela !", vbExclamation, "Alerte")
        GoTo Finally
    End If
    
    Call initValidationSheet
    Call SetValidations_SYNTH
    
    ' Vérifier si les dernières colonnes sont renseignées en bouclant sur les onglets de tests
    anyTestSheet = False
    For Each ws In Sheets
        'Si c'est un onglet de test
        If ws.Name Like "B2_???_???" Then
            anyTestSheet = True
            If Not verif_remplissage(ws.Name, True) Then GoTo Finally
        End If
    Next
    If Not anyTestSheet Then
        Call MsgBox("Auncune fiche de test existe !" _
            & vbCrLf & "Il est impossible de générer un PR sans cela !", vbExclamation, "Alerte")
        GoTo Finally
    End If
    
    ' Si un ancien PR out existe déjà, on le supprime
    If WsExist(PR_OUT_NAME) Then
        Application.DisplayAlerts = False
        Sheets(PR_OUT_NAME).Delete
        Application.DisplayAlerts = True
    End If
    
    'Si la feuille d'erreur existait déjà on la supprime
    If WsExist(ERROR_NAME) Then
        Application.DisplayAlerts = False
        Sheets(ERROR_NAME).Delete
        Application.DisplayAlerts = True
    End If
    
    Sheets(PR_MODEL_NAME).visible = xlSheetVisible
    Sheets(PR_MODEL_NAME).Copy Before:=Sheets(1)
    Sheets(PR_MODEL_NAME).visible = xlSheetHidden
    For Each ws In Sheets
        If ws.Name Like PR_MODEL_NAME & " (?)" Then
            ws.Name = PR_OUT_NAME
            ws.Move Before:=Sheets(1)
            If WsExist(PR_IN_NAME) Then
                Sheets(PR_IN_NAME).Move After:=ws
            End If
        End If
    Next
    
    
    'copier la synthèse filtrée pour que lignes principales de tests
    With Sheets(SYNTHESE_NAME)
        fin = .range("F" & .Rows.Count).End(xlUp).row
        With .range("$A$2:$I" & fin)
            .AutoFilter Field:=1, Criteria1:="<>"
            
            ' ---------------------------------------------------------------------------------------------------------
            ' Spécifité Prima forme Kazak
            For Each objrange In .Columns(2).SpecialCells(xlCellTypeVisible).Areas
                If objrange.Rows.Count < 2 Then
                    ' Si conf banc pas B ou vide, mettre en rouge + message d'erreur à la fin
                    If objrange.Value <> "" And objrange.Value <> "B" Then
                        With objrange.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                
                    ' Mettre B dans tous les cas
                    objrange.Value = "B"
                End If
            Next
            ' ---------------------------------------------------------------------------------------------------------
            
            Application.CutCopyMode = False
            .Copy
            Sheets(PR_OUT_NAME).range("A9").PasteSpecial Paste:=xlValue, Operation:=xlNone, SkipBlanks:=False, transpose:=False
            Application.CutCopyMode = False
            .AutoFilter Field:=1
        End With
    End With
    
    finPROUT = Sheets(PR_OUT_NAME).range("A" & Sheets(PR_OUT_NAME).Rows.Count).End(xlUp).row
    
    With Sheets(PR_OUT_NAME).range("A9:E" & finPROUT).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    'on marque la fin
    Sheets(PR_OUT_NAME).range("A" & finPROUT + 1) = "END"
    
    
    '---------------------------------------------------------------------------
    ' Copie des tests
    
    'boucler sur les numéros de test et chercher l'onglet correspondant
    Set testRange = Sheets(PR_OUT_NAME).range("A9")
    
    Do
        'If next_testRange <> "END" Then
        Set next_testRange = testRange.range("A2")
        
        ' Si le test existe et si tout va bien dedans, on colle son contenu dans PR de sortie
        If checkTest(testRange.Value) Then
            With Sheets(testRange.Value)
                
                tailleColle = .range("A" & .Rows.Count).End(xlUp).row
                
                ' S'il y a plus d'une ligne (cas standard)
                If tailleColle >= 3 Then
                
                    'On déplace la colonne "Mode" avant de faire la copie
                    Application.CutCopyMode = False
                    .Columns(2).Cut
                    .Columns("L:L").Insert Shift:=xlToRight
                    
                
                    ' Copie les données de la feuille de test correspondant à la ligne de test
                    Application.CutCopyMode = False
                    .range("A3:J" & tailleColle).Copy
                    
                    'insérer le contenu de l'onglet à la suite de la ligne de test en décalant les lignes vers le bas
                    next_testRange.EntireRow.Insert Shift:=xlDown
                    Application.CutCopyMode = False
                    testRange.range("A2:J" & tailleColle - 1).Cut Destination:=testRange.range("F2")
                
                    'On remet la colonne "Mode" à sa place
                    Application.CutCopyMode = False
                    .Columns("K:K").Cut
                    .Columns("B:B").Insert Shift:=xlToRight
                    
                ' Sinon, on a la dernière ligne de POPUP pour les cas dégradés
                Else
                    'On ajoute une ligne pour les exigences
                    next_testRange.EntireRow.Insert Shift:=xlDown
                End If
                
                'Copie des infos de la ligne de test
                Application.CutCopyMode = False
                .range("F2:K2").Copy
                testRange.range("J1").PasteSpecial Paste:=xlValue
                
                '--------------------------------------------------------------------------------------
                'Copie des exigences
                testRange.range("C1") = ""
                Set exigencesRange = Sheets(SYNTHESE_NAME).Columns(1).Find(what:=testRange.Value, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)
                If Not exigencesRange Is Nothing Then
                    exigences = Strings.Split(exigencesRange.range("C1"), Chr(10))
                    For i = 0 To UBound(exigences)
                        testRange.range("E" & 2 + i) = exigences(i)
                    Next i
                End If
                
                'On copie la colonne "Mode" dans la synthèse
                Application.CutCopyMode = False
                .range("B2:B" & tailleColle).Copy
                testRange.range("C1").PasteSpecial Paste:=xlValue
            End With
        
            Set testRange = testRange.range("A2").End(xlDown)
            
        Else
            If testRange <> "END" Then
                'Call MsgBox("Le test " & testRange.Value & " n'a pas été rajouté dans l'onglet de sortie !", vbExclamation, "Alerte")
                
                row = testRange.row
                'Et on redéfini le test suivant
                Set testRange = testRange.range("A2")
                
                'On supprime la ligne du test qui n'est pas insérer
                Application.DisplayAlerts = False
                Sheets(PR_OUT_NAME).Rows(row).EntireRow.Delete
                Application.DisplayAlerts = True
                
            End If
        End If
        
    Loop While StrComp(testRange.Value, "END", vbTextCompare) <> 0 'And testRange.row < next_testRange.row
    
    'Copie de l'entete
    Sheets("PDG").range("C4:C9").Copy
    With Sheets(PR_OUT_NAME)
        .range("B1").PasteSpecial Paste:=xlValue, Operation:=xlNone, SkipBlanks:=False, transpose:=False
        'On intervertie la version MPU avec Ref_FRScc depuis la version A5
        versionMPU = .range("B6")
        .range("B6") = .range("B5")
        .range("B5") = .range("B4")
        .range("B4") = versionMPU
    End With
        
    Call formatagePR_OUT

    '---------------------------------------
    ' Vérification Variables
    ' Pas tant qu'il y ait un fonctionnement de fichiers excel en réseau pour les Tref_FBS et Tref_Equipement_CB
    '
    'Call testPR_chemin(Sheets(2).Name)
    '---------------------------------------
    
    ' On montre la fiche produite
    'Sheets(PR_OUT_NAME).Activate
    'Sheets(PR_OUT_NAME).range("A9").Activate
    'On cache la fiche produite. elle ne sert que pour David en BDD
    Sheets(PR_OUT_NAME).visible = xlSheetHidden
    
    If WsExist(ERROR_NAME) Then
        Call FormatErrorSheet
        
        Message = "Des tests n'ont pas été rajoutés dans " & PR_OUT_NAME & " !"
        Call MsgBox(Message, vbCritical, "Alerte")
    Else
        Call MsgBox("La création du PR Out s'est bien déroulé.", vbInformation, "")
    End If
Finally:
    Application.ScreenUpdating = True
End Sub



Sub formatagePR_OUT()
    'Formatage
    With Sheets(PR_OUT_NAME)
        'Enleve les conditions de formatage de toute la feuille
        .Cells.FormatConditions.Delete
        
        'Enleve les ajouts de couleurs foireux depuis la colonne Num Etape
        With .range("F9", .range("O" & .Rows.Count).End(xlUp))
        
            With .Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        ' Colorie en bleu tous les tests et étapes
        
        With .range("F9", .range("O" & .Rows.Count).End(xlUp))
            .AutoFilter Field:=7, Criteria1:="<>"
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            .AutoFilter Field:=7
        End With
    
        .Columns("F").EntireColumn.AutoFit 'taille auto num_etape
        .Columns("H:I").WrapText = True
    End With
End Sub
