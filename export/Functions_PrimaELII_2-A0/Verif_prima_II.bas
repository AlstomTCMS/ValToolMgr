Attribute VB_Name = "Verif_prima_II"
Option Explicit


Sub Verifications(control As IRibbonControl)
    If HasActiveBook Then
        Call Verification
    End If
End Sub


'Utilisée pour PELII forme Kazak
Sub Verification_Old()
' Version du 28/11/2012
'Vérification de la présence de "END" à la fin du fichier
'Vérification des doublons de variables
'Vérification de l'ordre des types (AEn, CEn, ACc, CCc)
'La fonction sort à la première erreur

Dim nNbLigne, nNbLigneT, nNbLigneTotal As Long
Dim bErreur As Boolean

Dim TexteEtape_EnCours As String
Dim nNbLigneEtape_Deb, nNbLigneEtape_Fin As Long 'Num Ligne Debut et Fin d'Etape

Dim TexteType_EnCours As String
Dim nNbLigneType_Deb, nNbLigneType_Fin As Integer 'Num Ligne Debut et Fin de Type (AEn,CEn,ACc,CCc)
Dim PositionTypeAction As Integer

Dim Sheet2Verif As String
Sheet2Verif = PR_OUT_NAME

PositionTypeAction = 11
bErreur = False

If WsExist(Sheet2Verif) Then
With Sheets(Sheet2Verif)

' Forcage de la colonne "ID VARIABLE" en texte
    .Columns("P:P").NumberFormat = "@"
    
' Vérification En-tete du fichier
    For nNbLigne = 1 To 6
        If Len(.Cells(nNbLigne, 2)) <= 1 Then
            MsgBox "Vérifier l'En-Tête du fichier"
            GoTo Erreur
        End If
    Next
    
    'La vérification est déjà faite lors du collage des onglets de tests dans l'onglet "PR Out"
    'if Not verif_remplissage(Sheet2Verif, False) then Goto Erreur
    
    
    'on continue la suite des vérifs, on vérifie les doublons pour chaque Action/vérif d'1 Etape
    For nNbLigne = 9 To nNbLigneTotal
        .Cells(nNbLigne, 17) = .Cells(nNbLigne, 12) & .Cells(nNbLigne, 13) & .Cells(nNbLigne, 14)
    Next
     
    nNbLigneEtape_Deb = 9
    nNbLigne = 9
    While (.Cells(nNbLigne, 1) <> "END") And (nNbLigne <= nNbLigneTotal)
        TexteEtape_EnCours = .Cells(nNbLigneEtape_Deb, 6)
        nNbLigneEtape_Deb = nNbLigne
        For nNbLigne = nNbLigneEtape_Deb To nNbLigneTotal  'Trouve les bornes de l'Etape
            If (.Cells(nNbLigne, 6) <> TexteEtape_EnCours) Then ' Or (.Cells(nNbLigne, 6) = "") Then
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
        TexteType_EnCours = .Cells(nNbLigneType_Deb, PositionTypeAction)
        If (TexteType_EnCours = "ACc") And (bErreur = False) Then
            For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                If (.Cells(nNbLigne, PositionTypeAction) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                    nNbLigneType_Fin = nNbLigne - 1
                    Exit For
                End If
            Next
            bErreur = Doublon(nNbLigneType_Deb, nNbLigneType_Fin)
            nNbLigneType_Deb = nNbLigneType_Fin + 1
        End If
    '------------------------AEn
        TexteType_EnCours = .Cells(nNbLigneType_Deb, PositionTypeAction)
        If (TexteType_EnCours = "AEn") And (bErreur = False) Then
            For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                If (.Cells(nNbLigne, PositionTypeAction) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                    nNbLigneType_Fin = nNbLigne - 1
                    Exit For
                End If
            Next
            bErreur = Doublon(nNbLigneType_Deb, nNbLigneType_Fin)
            nNbLigneType_Deb = nNbLigneType_Fin + 1
        End If
    '------------------------CCc
        TexteType_EnCours = .Cells(nNbLigneType_Deb, PositionTypeAction)
        If (TexteType_EnCours = "CCc") And (bErreur = False) Then
            For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                If (.Cells(nNbLigne, PositionTypeAction) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                    nNbLigneType_Fin = nNbLigne - 1
                    Exit For
                End If
            Next
            bErreur = Doublon(nNbLigneType_Deb, nNbLigneType_Fin)
            nNbLigneType_Deb = nNbLigneType_Fin + 1
        End If
    '------------------------CEn
        TexteType_EnCours = .Cells(nNbLigneType_Deb, PositionTypeAction)
        If (TexteType_EnCours = "CEn") And (bErreur = False) Then
            For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                If (.Cells(nNbLigne, PositionTypeAction) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                    nNbLigneType_Fin = nNbLigne - 1
                    Exit For
                End If
            Next
            bErreur = Doublon(nNbLigneType_Deb, nNbLigneType_Fin)
            nNbLigneType_Deb = nNbLigneType_Fin + 1
        End If
    '------------------------PGM
        TexteType_EnCours = .Cells(nNbLigneType_Deb, PositionTypeAction)
        If (TexteType_EnCours = "PGM") And (bErreur = False) Then
            For nNbLigne = nNbLigneType_Deb To nNbLigneEtape_Fin + 1
                If (.Cells(nNbLigne, PositionTypeAction) <> TexteType_EnCours) Or (nNbLigne > nNbLigneEtape_Fin) Then
                    nNbLigneType_Fin = nNbLigne - 1
                    Exit For
                End If
            Next
            bErreur = False  'Doublon(nNbLigneType_Deb, nNbLigneType_Fin)
            nNbLigneType_Deb = nNbLigneType_Fin + 1
        End If
    '-----------------------------------------------
    
    
        nNbLigne = nNbLigneEtape_Fin + 1
        nNbLigneEtape_Deb = nNbLigneEtape_Fin + 1
        
        If bErreur = True Then
            nNbLigne = nNbLigneTotal + 1
            'MsgBox "Doublon dans l'Etape " & TexteEtape_EnCours, vbOKOnly + vbCritical, "Erreur doublon !"
            Call AjoutDoublon
        End If
        If (nNbLigneType_Deb <= nNbLigneEtape_Fin) And (bErreur = False) Then
            bErreur = True
            nNbLigne = nNbLigneTotal + 1
            MsgBox "L'ordre des types de variables (ACc, AEn, CCc, CEn) n'a pas été respecté pour l'étape : " & TexteEtape_EnCours, vbOKOnly + vbCritical, "Erreur !"
        End If
    Wend
    
    ' On efface la colonne qui a servit de tampon pour vérifier les doublons
    .Columns("Q:Q").ClearContents
    
End With

MsgBox "La Vérification du PR s'est achevée avec succès !", vbOKOnly + vbInformation, "Fin de la vérification !"
Exit Sub

Erreur:
    MsgBox "La Vérification du PR a été un échec !", vbOKOnly + vbCritical, "Fin de la vérification !"

End Sub


' Vérifie si c'est en doublon dans l'étape courante (entre début et fin)
'
Private Function Doublon(ByVal Deb As Integer, ByVal Fin As Integer) As Boolean
Dim Resultat As Boolean
Dim ligne1, ligne2 As Integer
Dim a, b
Doublon = False

    For ligne1 = Deb To Fin
        b = Cells(ligne1, 17)
        a = Application.WorksheetFunction.CountIf(range(Cells(Deb, 17), Cells(Fin, 17)), b)
        ' S'il y a plus qu'une occurence, alors on a un doublon
        If a > 1 Then
            Doublon = True
            
            'Créer une feuille des erreurs si elle n'existe pas
        End If
    Next

End Function
