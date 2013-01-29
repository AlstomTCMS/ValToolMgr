Attribute VB_Name = "Verif_prima_II"
Option Explicit


Sub Verifications(control As IRibbonControl)
    If HasActiveBook Then
        Call Verification
    End If
End Sub


'Utilis�e pour PELII forme Kazak
Sub Verification_Old()
' Version du 28/11/2012
'V�rification de la pr�sence de "END" � la fin du fichier
'V�rification des doublons de variables
'V�rification de l'ordre des types (AEn, CEn, ACc, CCc)
'La fonction sort � la premi�re erreur

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
    
' V�rification En-tete du fichier
    For nNbLigne = 1 To 6
        If Len(.Cells(nNbLigne, 2)) <= 1 Then
            MsgBox "V�rifier l'En-T�te du fichier"
            GoTo Erreur
        End If
    Next
    
    'La v�rification est d�j� faite lors du collage des onglets de tests dans l'onglet "PR Out"
    'if Not verif_remplissage(Sheet2Verif, False) then Goto Erreur
    
    
    'on continue la suite des v�rifs, on v�rifie les doublons pour chaque Action/v�rif d'1 Etape
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
            MsgBox "L'ordre des types de variables (ACc, AEn, CCc, CEn) n'a pas �t� respect� pour l'�tape : " & TexteEtape_EnCours, vbOKOnly + vbCritical, "Erreur !"
        End If
    Wend
    
    ' On efface la colonne qui a servit de tampon pour v�rifier les doublons
    .Columns("Q:Q").ClearContents
    
End With

MsgBox "La V�rification du PR s'est achev�e avec succ�s !", vbOKOnly + vbInformation, "Fin de la v�rification !"
Exit Sub

Erreur:
    MsgBox "La V�rification du PR a �t� un �chec !", vbOKOnly + vbCritical, "Fin de la v�rification !"

End Sub


' V�rifie si c'est en doublon dans l'�tape courante (entre d�but et fin)
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
            
            'Cr�er une feuille des erreurs si elle n'existe pas
        End If
    Next

End Function
