Attribute VB_Name = "Generation_Onglets"
Public modifiedTests As Variant
Public cancel_Synth2Tests As Boolean

' G�n�re les onglets de test � partir de la synth�se
Sub Generer_Onglets_Tests(control As IRibbonControl)
    If HasActiveBook Then
        Call Generer_OngletsTests
    End If
End Sub

' Efface les onglets de tests d�j� pr�sents dans le fichiers
Sub Supprimer_Onglets_Tests(control As IRibbonControl)
    If HasActiveBook Then
        Call SupprimerOngletsTests
    End If
End Sub

' G�n�re les onglets de test � partir de la synth�se
Public Sub Generer_OngletsTests()
    Dim testRange, debut, fin, finSynthese As range
    Dim testSheet As Worksheet
    Dim testTitle As Variant
    
    Application.ScreenUpdating = False
    
    If Not WsExist(SYNTHESE_NAME) Then
        MsgBox "L'onglet de synth�se n'existe pas ou n'est pas d�fini comme tel.", vbOKOnly + vbExclamation, "Fonctionnalit� non utilisable !"
        GoTo fin
    End If
    
    If SupprimerOngletsTests Then
        Call RedefineSyntheseArray
        Call deleteExigencesFromSynth
    
        Set testRange = Sheets(SYNTHESE_NAME).range("A2")
        testTitle = Array("Num_Etape", "Com_Etape", "Com_act", "Com_chk", "Pause", "Type_Var", "Vehicule", "Variable", "Chemin", "Valeur")
        Set finSynthese = Sheets(SYNTHESE_NAME).range("F2").End(xlDown).range("D1") 'cellule de l'angle du tableau par Num_Etape
        
        Do
            Call reformatExigences(testRange.range("C1"))
        'If exigencesExist(testRange.range("C1")) Then
            ' Cr�er une feuille du nom du test
            
            'Si l'onglet de test n'existe pas d�j�
        If Not WsExist(testRange.Value) Then
            Set testSheet = InitSheet(testRange.Value, True, , , testTitle)
            
            With testSheet
                
                'coller la zone Synth�se 3 derni�res colonnes dans B2
                Set debut = testRange.range("F1")
                
                'Tester s'il n'y a qu'une ligne principale pour ce test
                If testRange.range("A2") <> "" Then
                    Set fin = testRange.range("I1")
                ' Si on atteind la fin du tableau de synth�se, on ne doit pas faire d'offset
                ElseIf testRange.range("A2").End(xlDown).row = finSynthese.row Then
                    Set fin = testRange.range("A2").End(xlDown).range("I1")
                Else
                    Set fin = testRange.range("A2").End(xlDown).range("I1").Offset(-1, 0)
                End If
                
                If fin.row > finSynthese.row Then
                    Set fin = finSynthese
                End If
                
                Sheets(SYNTHESE_NAME).range(debut, fin).Copy
                .range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False
                Application.CutCopyMode = False
    
                .Columns("A:B").EntireColumn.AutoFit
                .Columns("C:D").ColumnWidth = 24
                .Columns("C:D").WrapText = True
                
                'Num_Etape
                nbreEtape = fin.row - debut.row + 1
                For i = 1 To nbreEtape
                    .Cells(i + 1, 1) = testRange.Value & "-" & Format(i, "00")
                Next
                
                'Ajout du commentaire pour les Type de variables permis
                With .range("F1")
                    If .Comment Is Nothing Then
                        .AddComment
                        .Comment.visible = True
                        .Comment.Text Text:= _
                            "Types permis (dans l'ordre):" & Chr(10) & "AEn;CEn;ACc;CCc"
                        .Comment.Shape.Left = 590
                        .Comment.Shape.Top = 26
                    End If
                End With
                
                'Ajouter liens vers �tapes
                Call ajouteLiens(testRange.range("A1:A" & nbreEtape))
                
                ' Ins�rer la colonne "Modes"
                testSheet.Columns(2).Insert Shift:=xlToRight
                testSheet.range("B1") = "Mode"
    
                Call formatageFicheTest(testSheet.Name)
            End With
        End If
        
        'S'il n'y a qu'une ligne principale pour ce test
        If testRange.range("A2") <> "" Then
            Set testRange = testRange.range("A2")
        Else
            Set testRange = testRange.range("A2").End(xlDown)
        End If
            
        Loop While testRange <> ""
        
        Call formatageSynthese 'On reformate la synth�se qui a �t� modifi�
        
        Sheets(SYNTHESE_NAME).Columns("F").EntireColumn.AutoFit
        Sheets(SYNTHESE_NAME).Activate
        Sheets(SYNTHESE_NAME).range("J1").Activate
    End If
    
fin:
    Application.ScreenUpdating = True
End Sub

' Ajoute lien sur le num�ro de test et les num�ros d'�tapes de la fiche synth�se vers l'onglet correspondant
Sub ajouteLiens(zoneLien As range)
Dim test_num As String
Dim cellToLink As range

    test_num = zoneLien.range("A1").Value
    
    With Sheets(SYNTHESE_NAME)
        .Hyperlinks.Add Anchor:=zoneLien, Address:="", SubAddress:= _
            "'" & test_num & "'!A2", TextToDisplay:=test_num
        
        'Ajout des liens vers les �tapes
        For Each cell In zoneLien.Columns("F").Rows
            .Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:= _
            "'" & test_num & "'!" & Replace(cell.Offset(2 - zoneLien.row, -5).Address, "$", ""), TextToDisplay:=cell.Value
        Next
    End With
End Sub


'V�rifie si l'exigence existe dans la table de r�f�rencement des exigences du projet.
Function exigencesExist(ByVal num_ex As String) As Boolean
Dim exRow As range
Dim exigences As Variant

    'sur la colonne 3 "Exigence associ�e" de la fiche synth�se
    'voir si pr�sente dans la feuille des exigences d'Alexandra
    'pour l'instant sur la colonne 3 "Exigences_Titre" de Tref_Exigences
    exigencesExist = True
    
    exigences = Strings.Split(num_ex, Chr(10))
    
    For i = 0 To UBound(exigences)
        Set exRow = Sheets("Tref_Exigences").Columns(3).Find(what:=exigences(i), LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)
        If exRow Is Nothing Then
            exigencesExist = False
        End If
    Next i
    
End Function

Sub reformatExigences(exigenceRange As range)
Dim exigences As Variant
Dim exigencesString As String
Dim exigencesTampon As Variant

    exigencesTampon = Strings.Split(exigenceRange, Chr(10))
    'Si on a qu'un seul �l�ment on v�rifie que ce n'est pas � cause du mauvais s�parateur
    If UBound(exigencesTampon) = 0 Then
        exigences = Strings.Split(RTrim(exigenceRange), " ")
        'Si on a moins d'�l�ment en coupant par Alt+enter que par l'espace,
        'c'est que c'est s�par� par des espaces et qu'il faut r��crire avec Alt+enter
        If UBound(exigencesTampon) < UBound(exigences) Then
            exigencesString = exigences(0)
            For i = 1 To UBound(exigences)
                exigencesString = exigencesString & Chr(10) & exigences(i)
            Next
            exigenceRange = exigencesString
        End If
    End If
End Sub


' Supprime les onglets de tests s'il y en a et que l'utilisateur veut les supprimer
' Renvoie vrai s'il y en a et que l'utilisateur veut les supprimer
' Renvoie vrai s'il n'y en a pas
' Renvoie faux s'il y en a et que l'utilisateur ne veut pas les supprimer
Private Function SupprimerOngletsTests() As Boolean
    SupprimerOngletsTests = False
    
    'r�init les test � modifier
    modifiedTest = Null
    
    'On test d'abord s'il y a des onglets de test pour ne pas afficher le message inutilement
    ancienTestsExiste = False
    For Each ws In Sheets
        'Autres Formats
        If ws.Name Like "K8_*" Or ws.Name Like "B????_*" Or ws.Name Like "E????_*" Then
            ancienTestsExiste = True
            Exit For
        End If
    Next
    
    If ancienTestsExiste Then
        If MsgBox("Voulez vous supprimer les onglets de tests actuels ?" & vbCrLf & _
        "Si vous g�n�rez � partir de la synth�se, vous perdrez toutes les informations � partir de la colonne 'Pause'." _
        & vbCrLf & vbCrLf & "Si cela n'est pas fait, les onglets de tests peuvent �tre incoh�rent avec la synth�se." _
        , vbExclamation + vbOKCancel, "Suppression des tests") = vbOK Then
        
            SupprimerOngletsTests = False
            Application.DisplayAlerts = False
            On Error GoTo Finally
            For Each ws In Sheets
                If ws.Name Like "K8_*" Or ws.Name Like "B????_*" Or ws.Name Like "E????_*" Then
                    ws.Delete
                End If
                If PR2Synth And ws.Name Like "B2_???_???*" Then
                    ws.Delete
                End If
            Next
            Application.DisplayAlerts = True
        Else
            SupprimerOngletsTests = False
        End If
    End If
    
    If WsExist(SYNTHESE_NAME) Then
        'Afficher une fenetre pour choisir les tests qui ont �t� modifi�s
        UserForm_modifiedTests.Show
        
        ' Si l'utilisateur n'annule pas l'action
        If Not cancel_Synth2Tests Then
            ' Si des tests ont �t� d�finis comme modifi�s
            If modifiedTests <> "" Then
                'Supprimer les onglets de ces tests l� pour qu'ils soient reg�n�r�
                SupprimerOngletsTests = True
                sheets2Delete = Split(modifiedTests, ";")
                
                Application.DisplayAlerts = False
                On Error GoTo Finally
                For i = 0 To UBound(sheets2Delete) - 1
                    Sheets(sheets2Delete(i)).Delete
                Next i
                Application.DisplayAlerts = True
            End If
        End If
    End If
    
Finally:
    On Error GoTo 0
    Application.DisplayAlerts = True
End Function


' Red�finie la taille du tableau de Synth�se en fonction de la r�alit�
Sub RedefineSyntheseArray()

    Dim oNewRow As ListRow
    
    With Sheets(SYNTHESE_NAME)
        'R�cup�rer la vraie fin sur la colonne Etapes qui ne peut pas avoir d'�l�ment vide
        i = 2
        While Len(.Cells(i, 6)) > 0
            i = i + 1
        Wend
        
        'On cherche le tableau de synth�se qui se trouve dans la feuille de synth�se
        For Each Object In .ListObjects
            If Object.Name Like "TableauSynth�se*" Then
            'Si la fin du tableauSynth�se et la fin r�elle du tableau ne co�ncident pas, il faut r�ajuster
            If Object.ListRows.Count < i - 2 Then
            
                'ins�re une ligne suppl�mentaire � la fin de la zone tableau
                Set oNewRow = .ListObjects("TableauSynth�se").ListRows.Add(AlwaysInsert:=True)
                
                'on coupe la zone hors du tableau
                .range("A" & oNewRow.Index + 2 & ":I" & i).Cut '171
                
                'on r�insert cette zone dans le tableau � sa fin
                .Rows(oNewRow.Index + 1).Insert Shift:=xlDown
                
                'on supprime la derni�re ligne qui a �t� rajout�e dans le tableau
                .ListObjects("TableauSynth�se").ListRows(i - 1).Delete
                
                'Supprimer les lignes vides du UsedRange par le couper/coller
                'tailleTableau = .ListObjects("TableauSynth�se").ListRows.Count
                '.range("A" & tailleTableau + 2, "A" & 2 * (tailleTableau + 1) - oNewRow.Index + 1).EntireRow.Delete
                Exit For
            End If
            End If
        Next Object
        
    End With
    
    Call SetValidations_SYNTH
End Sub
