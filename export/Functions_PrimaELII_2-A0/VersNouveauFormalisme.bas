Attribute VB_Name = "VersNouveauFormalisme"
Private PR_In_Sheet As Worksheet


Sub Ancien_Vers_Nouveau(control As IRibbonControl)
    If HasActiveBook Then
        Call AncienVersNouveau
    End If
End Sub

'NON UTILISEE POUR L'INSTANT !
Sub unmaskFirstLignes_aff(control As IRibbonControl)
    If HasActiveBook Then
        Call unmaskFirstLignes
    End If
End Sub


'Démasque les premières lignes d'un PR si elles sont cachées
Sub unmaskFirstLignes()
    With Sheets(1)
        If .Rows("1:7").EntireRow.Hidden Then
            .Rows("1:7").EntireRow.Hidden = False
            .range("B1").Select
        End If
    End With
End Sub


'Fonction principale qui à partir du PR créer une synthèse et les onglets de tests qui correspondent
Public Sub AncienVersNouveau()
    
    ' Empeche le rafraichissement de la fenetre
    Application.ScreenUpdating = False
    
    'Si une feuille de synthèse existe déjà, demander si l'utilisateur veut vraiment lancer le processus
    If WsExist(SYNTHESE_NAME) Then
        'Il faudrait pouvoir quitter la création de synthèse si on le désire
        If vbNo = MsgBox("Une synthèse de PR existe déjà." & vbCrLf & "Voulez vous écrire par dessus ses données ?", vbExclamation + vbYesNo, "Attention") Then
            GoTo Finally
        End If
    End If
    
    ' On vérifie si on a déjà une feuille nommée "PR IN" dans quel cas on la choisie
    If WsExist(PR_IN_NAME) Then
        Set PR_In_Sheet = Sheets(PR_IN_NAME)
    Else
    'Sinon on prend le premier onglet
        Set PR_In_Sheet = Sheets(1)
    End If
    
    
    'Verifier que le fichier a num_PR en A1
    If Not PR_In_Sheet.range("A1") = "Num_PR" Then
        Call MsgBox("La génération ne peut pas se faire car la feuille 1 n'est pas un PR.", vbExclamation, "Alerte")
        GoTo Finally
    End If
    
    Call CopySheetsFromRef
    
    'Vérifier que la fiche 1 est bien un PR.
    If Not matchRange(PR_In_Sheet.range("A1:A3"), Sheets(PR_MODEL_NAME).range("A1:A3")) Then
        Call MsgBox("La génération ne peut pas se faire car la feuille 1 n'est pas un PR.", vbExclamation, "Alerte")
        GoTo Finally
    End If
    
    Call virerFeuillesInutiles
    If Not checkIsPrimaOldVersion Then GoTo Finally
        
    'Copie de l'entete dans la page de garde
    PR_In_Sheet.range("B1:B6").Copy
    With Sheets("PDG")
        .range("C4").PasteSpecial Paste:=xlValue, Operation:=xlNone, SkipBlanks:=False, transpose:=False
        'On intervertie la version MPU avec Ref_FRScc depuis la version A5
        versionMPU = .range("C7")
        .range("C7") = .range("C8")
        .range("C8") = .range("C9")
        .range("C9") = versionMPU
        .Activate
        .range("A1").Select
    End With
    
    Call FillSynthese
    Call SupprimerOngletsTests
    Call CreateAndFill_Tests
    Call formatageSynthese
    
    '---------------------------------------
    ' Vérification Variables
    ' Pas tant qu'il y ait un fonctionnement de fichiers excel en réseau pour les Tref_FBS et Tref_Equipement_CB
    '
    'Call testPR_chemin(PR_In_Sheet.Name)
    '---------------------------------------
    
    'On cache la feuille de l'ancien format pour éviter les modifications de part et d'autres
    PR_In_Sheet.visible = xlSheetHidden
    
Finally:
    ' Permet le rafraichissement de la fenetre
    Application.ScreenUpdating = True
End Sub



'Function qui dit si le contenu de deux colonnes sont égaux
Function matchRange(column1 As range, column2 As range) As Boolean
    matchRange = True
    For Each cell In column1
        If Not column2(cell.row) = cell Then
            matchRange = False
            Exit For
        End If
    Next
End Function

'return false s'il y a eu une erreur
Function checkIsPrimaOldVersion() As Boolean
checkIsPrimaOldVersion = True
    
    With PR_In_Sheet
        .Activate
        'Si c'est bien un PR, la renommer PR_IN_NAME et faire la génération
        .Name = PR_IN_NAME
        
        
        ' Si le PR n'est pas dans le deuxième format de Prima, on le renumérote
        If Not .range("A9") Like "B2_???_???" Then
            
            '-----------------------------------------------------------------------------------------------------
            'Si la version actuelle n'est pas du prima 2.2 on demande à l'utilisateur de renseigner l'entete avant
            num_PR = .range("B1")
            'Si vide ou format KZH ou PRIMA2.1
            If Not num_PR Like "B2_???_?" Then
                'Ajout du commentaire pour les Type de variables permis
                With .range("B1")
                    If .Comment Is Nothing Then
                        .AddComment
                        .Comment.visible = True
                        .Comment.Text Text:="Format permis: B2_XXX_Y avec XXX numéro de fonction et Y index de feuille"
                        '.Comment.Shape.Left = 590
                        '.Comment.Shape.Top = 26
                    End If
                End With
                
                .Activate
                .range("B1").Select
                Call MsgBox("Il faut renseigner l'entête dans le format PRIMA ELII.2 (B2_XXX_Y) pour pouvoir générer !", vbExclamation, "Impossible de générer")
                checkIsPrimaOldVersion = False
                Exit Function
            Else
                ' Virer le commentaire s'il y est
            End If
            
            'Si on a pas les nouvelles colonnes du format KZH, on les rajoute
            If .range("B8") = "Des_Etape" Then
                .Columns(2).Insert Shift:=xlToRight
                .range("B7") = "A1,A2,B,C,D"
                .range("B8") = "Modes"
                .Columns(2).Insert Shift:=xlToRight
                .range("B7") = "MPU1,MPU2,MPUX"
                .range("B8") = "Conf_Banc"
                .Columns(3).EntireColumn.AutoFit
                .range("D8") = "Des_Test"
                
                'remettre les infos dans B et enlever formatage
                .Activate
                Application.CutCopyMode = False
                .range("D1:D6").Cut Destination:=range("B1:B6")
                Application.CutCopyMode = False
                With .range("C1:C6").Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
            
            'Si fichiers venant de chez ASSYSTEM
            If .range("D8") = "Des_Etape" Then
                .range("D8") = "Des_Test"
            End If
            
            
            If Not Numerotation_Automatique Then
                checkIsPrimaOldVersion = False
            End If
        End If
    End With
End Function

' Créé et remplie la fiche de synthèse à partir du PR
Sub FillSynthese()
Dim synthSheet As Worksheet
Dim synthExist As Boolean
Dim synthTitle As Variant
synthTitle = Array("Test", "Conf banc", "Exigence(s) associée(s)", "Description Test", "Commentaires Test", "Etapes", "Commentaires Etapes", "Description Actions", "Description Vérification")
    
    'Supprimer l'objet TableauSynthèse s'il existe déjà
    'Application.Goto Reference:="TableauSynthèse_1"
    'Selection.Delete
    'Rows("1:2").Delete Shift:=xlUp
    
    'Faire un onglet pour la Synthèse si il n'existe pas déjà
    Set synthSheet = InitSheet(SYNTHESE_NAME, , , synthExist, synthTitle)
    If synthExist Then
        Application.DisplayAlerts = False
        Sheets(SYNTHESE_NAME).Delete
        Application.DisplayAlerts = True
        Set synthSheet = InitSheet(SYNTHESE_NAME, , , synthExist, synthTitle)
    End If
    
    syntRange = synthSheet.range("A2")
        
    ' --------------------------------------------------------------------------------------------------
    ' Remplissage de la fiche de synthèse
    
    'On filtre sur les Com_etapes pour voir l'ensemble sous forme de synthèse
    'et on recopie le contenue dans la feuille de synthèse
    With PR_In_Sheet
                      
        
        'On enleve le filtre qu'il pourrait y avoir sur la colonne "Com_Etape"
        '.Columns(1).AutoFilter Field:=7
        fin = .Columns(1).Find(what:="END", MatchCase:=False).row
        With .range("$A$9:$I" & fin)
            .AutoFilter
            
            .AutoFilter Field:=7, Criteria1:="<>"
            Application.CutCopyMode = False
            .Copy
            'On copie dans la feuille de synthèse
            Sheets(SYNTHESE_NAME).range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False
            Application.CutCopyMode = False
            .AutoFilter Field:=7
            
            
        End With
        
        
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
    End With
            
    'Supprimer les Modes de la colonne exigence
    With Sheets(SYNTHESE_NAME)
        .range("C2:" & .range("C" & .Rows.Count).End(xlUp).Address).ClearContents
    End With
End Sub


Public Sub formatageSynthese()
Dim testRange, next_testRange As range
Dim zoneLien As range
Dim fin As Integer

    With Sheets(SYNTHESE_NAME)
        For Each Object In .ListObjects
            If Object.Name Like "TableauSynthèse*" Then
                Object.TableStyle = "tableau de Synthèse"
            End If
        Next
        
        'boucler par test
        Set testRange = .range("A2")
        fin = .UsedRange.Rows.Count
        Do
            'on cherche le suivant
            If testRange.range("A2") <> "" Then
                Set next_testRange = testRange.range("A2")
            Else
                Set next_testRange = testRange.End(xlDown)
            End If
            
            'formater à gauche et à droite
            If next_testRange.row >= fin Then
                Set next_testRange = .range("A" & fin)
                Call FormatSyntheseGauche(range(testRange, next_testRange.range("E1")))
                Call FormatSyntheseDroite(range(testRange.range("F1"), next_testRange.range("I1")))
                Set zoneLien = range(testRange.range("A1"), next_testRange.range("F1"))
            Else
                Call FormatSyntheseGauche(range(testRange, next_testRange.range("E1").Offset(-1, 0)))
                Call FormatSyntheseDroite(range(testRange.range("F1"), next_testRange.range("I1").Offset(-1, 0)))
                Set zoneLien = range(testRange.range("A1"), next_testRange.range("F1").Offset(-1, 0))
            End If
            
            Call MajLiensSynthese(zoneLien)
            
            Set testRange = next_testRange
        Loop While testRange <> "" And testRange.row < fin
        
        
        .Move After:=Sheets(2)
        
        .Activate
        .range("J1").Activate
        
        .Columns("B").ColumnWidth = 3
        .Columns("C").ColumnWidth = 18
        .Columns("D:E").ColumnWidth = 24
        .Columns("C:E").WrapText = True
        .Columns("F").EntireColumn.AutoFit 'Num_etape autofit
        'retour à la ligne de "Description Actions" et "Description Vérification"
        .Columns("G:I").ColumnWidth = 24
        .Columns("G:I").WrapText = True
        
        'POur toutes les colonnes
        With .Columns("A:I")
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With

    End With
    
    Call SetValidations_SYNTH
End Sub


' Ajoute lien sur le numéro de test et les numéros d'étapes de la fiche synthèse vers l'onglet correspondant
Sub MajLiensSynthese(zoneLien As range)
Dim test_num As String
Dim cellToLink As range

    test_num = zoneLien.range("A1").Value
    
    With Sheets(SYNTHESE_NAME)
        .Hyperlinks.Add Anchor:=zoneLien.range("A1"), Address:="", SubAddress:= _
            "'" & test_num & "'!A2", TextToDisplay:=test_num
        
        'Ajout des liens vers les étapes
        For Each cell In zoneLien.Columns("F").Rows
            Set cellToLink = Sheets(test_num).Columns(1).Find(what:=cell.Value, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=True)
            If Not cellToLink Is Nothing Then
            
                .Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:= _
                "'" & test_num & "'!" & Replace(cellToLink.Address, "$", ""), TextToDisplay:=cell.Value
            End If
        Next
    End With
End Sub

Private Sub FormatSyntheseGauche(range As range)

    With range
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
End Sub


Private Sub FormatSyntheseDroite(range As range)

    With range
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub

'Création des onglets des tests
Sub CreateAndFill_Tests()
Dim anc_testRange As range
Dim anc_nextTestRange As range
Dim nvo_syntRange As range
Dim debut As range: Dim fin As range
Dim testSheet As Worksheet
Dim testTitle As Variant


testTitle = Array("Num_Etape", "Com_Etape", "Com_act", "Com_chk", "Pause", "Type_Var", "Vehicule", "Variable", "Chemin", "Valeur")
    
    Set anc_testRange = PR_In_Sheet.range("A9")
    
    'On filtre Synthèse pour n'avoir que les lignes de test sans toutes les étapes
    Sheets(SYNTHESE_NAME).UsedRange.AutoFilter Field:=1, Criteria1:="<>"
    
    'Boucler sur Num_test
    Do
        ' On regarde où est le prochain test
        Set anc_nextTestRange = anc_testRange.range("A2").End(xlDown)
        
        ' ---------------------------------------------------------------------------------------------------------
        ' Spécifité Prima forme Kazak
        ' ---------------------------------------------------------------------------------------------------------
        ' Si conf banc A1 et Vehicule X, mettre B et 1
        If anc_testRange.range("B1") = "A1" Then
        ' Vérifier pour ce test si on a un véhicule X
            '  définir la zone du test
            With PR_In_Sheet.range(anc_testRange, anc_nextTestRange.range("L1"))
                .AutoFilter
                '  filtrer véhicule colonne L par "X"
                .AutoFilter Field:=12, Criteria1:="X"
                '  remplacer tous les X par 1
                ' Si la zone n'est pas vide il faut remplacer par 1 et conf_banc par B
                On Error Resume Next
                With .Columns("L").SpecialCells(xlCellTypeVisible)
                    ' Boucler sur les Areas
                    For Each objrange In .Areas
                        objrange.Columns(1) = "1"
                    Next objrange
                    anc_testRange.range("B1") = " B"
                End With
                On Error GoTo 0
                .AutoFilter Field:=12
            End With
        End If
        
        '------------------------------------------------------
        ' Copier tout le bloc des etapes dans l'onglet de test
        '------------------------------------------------------
        ' Créer une feuille du nom du test
        Set testSheet = InitSheet(anc_testRange.Value, True, , , testTitle)
        'connaitre les ranges des deux coins de la diagonale
        Set debut = anc_testRange.range("F1")
        Set fin = anc_nextTestRange.range("O1").Offset(-1, 0)
        'Copier cet ensemble dans l'onglet de test
        Application.CutCopyMode = False
        PR_In_Sheet.range(debut, fin).Copy
        testSheet.range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False
        Application.CutCopyMode = False
        
        '------------------------------------------------------
        ' Insérer les Modes
        '------------------------------------------------------
        testSheet.Columns(2).Insert Shift:=xlToRight
        testSheet.range("B1") = "Mode"
        Set debut = anc_testRange.range("C1")
        Set fin = anc_nextTestRange.range("C1").Offset(-1, 0)
        PR_In_Sheet.range(debut, fin).Copy
        testSheet.range("B2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False
        Application.CutCopyMode = False

        Call formatageFicheTest(testSheet.Name)
        Call copyExigencesFormPR(anc_testRange)
        
        'On va au prochain test
        Set anc_testRange = anc_nextTestRange
    Loop While anc_testRange <> "END"
    
    Sheets(SYNTHESE_NAME).UsedRange.AutoFilter Field:=1
End Sub

Sub test_Unit_formatageFicheTest()
    Call formatageFicheTest("B2_041_101")
End Sub

'Formatage final d'une feuille de test
Sub formatageFicheTest(ByVal testSheetName As String)
    With Sheets(testSheetName)
        .Columns("A:B").EntireColumn.AutoFit
        .Columns("C:E").ColumnWidth = 24
        .Columns("C:E").WrapText = True
        .Columns("F").EntireColumn.AutoFit
        .Columns("I").EntireColumn.AutoFit
        .Columns("G").ColumnWidth = 4
        .Columns("H").EntireColumn.AutoFit
        .Columns("J").ColumnWidth = 27
        .Columns("K").EntireColumn.AutoFit
        .Activate
        .range("A2:K2").Select
        ActiveWindow.Zoom = True
        .range("L1").Activate
    End With
    
    Call initValidationSheet
    Call SetConditionalFormat_TEST(testSheetName)
    Call SetValidations_TEST(testSheetName)
End Sub

'Copie des exigences
Sub copyExigencesFormPR(ByRef testRange As range)
Dim exigenceRange As range
Dim exigences As String

        exigences = ""
        'Sinon on les cherche dans la colonne E "Com_Test"
        If testRange.range("E2") <> "" Then
            Set exigenceRange = testRange.range("E2")
        'On prend les exigences de la colonne C "Mode"
        ElseIf testRange.range("C2") <> "" And Not testRange.range("C2") Like "MPU?" Then
            Set exigenceRange = testRange.range("C2")
        Else
            Exit Sub
        End If
        
        If Not exigenceRange Is Nothing Then
            For Each cell In range(exigenceRange, exigenceRange.End(xlDown))
                If cell.Value = "" Then
                    Exit For
                Else
                    If exigences = "" Then
                        exigences = cell.Value
                    Else
                        exigences = exigences & Chr(10) & cell.Value
                    End If
                End If
            Next cell
        End If
        Set nvo_syntRange = Sheets(SYNTHESE_NAME).Columns(1).Find(what:=testRange.Value, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)
        If Not nvo_syntRange Is Nothing Then
            nvo_syntRange.range("C1") = LTrim(exigences)
        End If
End Sub

' Supprime les onglets de tests s'il y en a et que l'utilisateur veut les supprimer
' Renvoie vrai s'il y en a et que l'utilisateur veut les supprimer
' Renvoie vrai s'il n'y en a pas
' Renvoie faux s'il y en a et que l'utilisateur ne veut pas les supprimer
Private Function SupprimerOngletsTests() As Boolean
    SupprimerOngletsTests = False
    
    'réinit les test à modifier
    modifiedTest = Null
    
    'On test d'abord s'il y a des onglets de test pour ne pas afficher le message inutilement
    ancienTestsExiste = False
    For Each ws In Sheets
        If ws.Name Like "B2_???_???" Or ws.Name Like "K8_*" Or ws.Name Like "B????_*" Or ws.Name Like "E????_*" Then
            ancienTestsExiste = True
            Exit For
        End If
    Next
    
    If ancienTestsExiste Then
        If MsgBox("Voulez vous supprimer les onglets de tests actuels ?" & vbCrLf & _
        "Si vous générez à partir de la synthèse, vous perdrez toutes les informations à partir de la colonne 'Pause'." _
        & vbCrLf & vbCrLf & "Si cela n'est pas fait, les onglets de tests peuvent être incohérent avec la synthèse." _
        , vbExclamation + vbOKCancel, "Suppression des tests") = vbOK Then
            Application.DisplayAlerts = False
            On Error GoTo Finally
            For Each ws In Sheets
                If ws.Name Like "B2_???_???" Or ws.Name Like "K8_*" Or ws.Name Like "B????_*" Or ws.Name Like "E????_*" Then
                    ws.Delete
                End If
            Next
            Application.DisplayAlerts = True
        Else
            SupprimerOngletsTests = False
        End If
    End If
    
Finally:
    On Error GoTo 0
    Application.DisplayAlerts = True
End Function
