Attribute VB_Name = "verif_chemin_VS_type"
Public listvar As Variant

' Vérifie si les variables existent
Sub Verif_cheminVStype(control As IRibbonControl)
    If HasActiveBook Then

    'Call MsgBox("Cette fonctionnalité n'est pas utilisable pour le moment")
    '---------------------------------------
    ' Vérification Variables
    ' Pas tant qu'il y ait un fonctionnement de fichiers excel en réseau pour les Tref_FBS et Tref_Equipement_CB
    '
    Call testPR_chemin(PR_OUT_NAME)
    '---------------------------------------
    End If
    
End Sub

Sub testPR_chemin(sheetName As String)
Dim listvar As Variant
Dim erreurPossible As Boolean

    erreurPossible = False
    listvar = getListVar

    With Sheets(sheetName)
        With .range("A9:O" & .range("F" & .Rows.Count).End(xlUp).row)
            ' On efface les couleurs mises avant
            With .Columns("N").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
        
            .AutoFilter Field:=14, Criteria1:=listvar, Operator:=xlFilterValues
            
            'On les colorie en bleu
            With .Columns("N").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16763955
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
    
    
            ' on trie toutes celles qui ne sont pas bleu et on les met en rouge
            .AutoFilter Field:=14, Operator:= _
                xlFilterNoFill
            On Error Resume Next
            'With .SpecialCells(xlCellTypeVisible)
                With .Columns("N").SpecialCells(xlCellTypeVisible).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                    
                    erreurPossible = True
                End With
            'End With
            On Error GoTo 0
            
            .AutoFilter Field:=14, Criteria1:=RGB(51, _
                204, 255), Operator:=xlFilterCellColor
            With .Columns("N").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            .AutoFilter Field:=14
            
            If Sheets(sheetName).visible = xlSheetVisible Then
                Sheets(sheetName).Activate
                Sheets(sheetName).range("N9").Select
            End If
            
            ' on affiche un message d'erreur si rouge il y a
            If erreurPossible Then
                Call MsgBox("Il est possible qu'il y ait des erreurs sur les chemins. Ceux-ci sont indiqués en rouge dans la fiche " & sheetName & ".", , "Vérifications des chemins")
            End If
            
        End With
        
        
    End With
End Sub


' Récupère la liste des variables environnement et embarquées
Private Function getListVar() As Variant
    Dim listvar As Variant
    Dim row As Integer
    Dim taille As Long
    
    'On copie les variables environnement dedans
    With Sheets("Tref_FBS")
    
        taille = .range("B" & .Rows.Count).End(xlUp).row - 1
        ReDim listvar(1 To taille)
        For i = 1 To taille
            listvar(i) = .range("B" & i + 1)
        Next
    End With
    
    'On copie les variables Embarquées dedans
    With Sheets("Tref_EquipementCB")
    
        taille2 = .range("B" & .Rows.Count).End(xlUp).row - 1
        taille3 = UBound(listvar) + taille2
        ReDim Preserve listvar(1 To taille3)
        For i = 1 To taille2
            listvar(i + taille) = .range("B" & i + 1)
        Next
    End With
    
    getListVar = listvar
End Function
