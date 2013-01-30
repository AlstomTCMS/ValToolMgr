Attribute VB_Name = "Num_Automatique"

'Renvoie True si �a s'est bien pass�
Function Numerotation_Automatique() As Boolean

Dim NumTest, NumPR, Projet As String
Dim x, y, test, NumEtape, Colonne_Des_Test, Colonne_Num_Etape, Colonne_Com_Etape As Integer
Dim checkRange, numPR_range As range
Dim columns2check_String, columns2check_int As Variant

'Init Variables
Numerotation_Automatique = True
x = 1
test = 1
columns2check_String = Array("Num_PR", "Des_Test", "Num_Etape", "Com_Etape")

ReDim columns2check_int(0 To 3)

With Sheets(PR_IN_NAME)
    .Activate
    
    Set numPR_range = .Columns(1).Find(what:="Num_PR", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    If numPR_range Is Nothing Then
        Call MsgBox("Num�ro de PR introuvable!" & vbCr & "Arr�t de la num�rotation!", vbExclamation, "Attention!")
        Numerotation_Automatique = False
        Exit Function
    Else
        NumPR = numPR_range.range("B1")
        Set numPR_range = .Columns(1).Find(what:="Num_Test", LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
        If numPR_range Is Nothing Then
            Call MsgBox("La colonne Num_Test est introuvable!" & vbCr & "Arr�t de la num�rotation!", vbExclamation, "Attention!")
            Numerotation_Automatique = False
            Exit Function
        Else
            For i = 1 To UBound(columns2check_String)
                Set checkRange = Rows(numPR_range.row).Find(what:=columns2check_String(i), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
                If checkRange Is Nothing Then
                    Call MsgBox("Colonne " & columns2check_String(i) & " introuvable!" & vbCr & "Arr�t de la num�rotation!", vbExclamation, "Attention!")
                    Numerotation_Automatique = False
                    Exit Function
                Else
                    columns2check_int(i) = checkRange.Column
                End If
            Next
        End If

        x = 9
        Colonne_Des_Test = columns2check_int(1)
        Colonne_Num_Etape = columns2check_int(2)
        Colonne_Com_Etape = columns2check_int(3)
    
        'Boucle de num�rotation
        Do Until .Cells(x, 1) = "END"
            'V�rifie si la cellule doit contenir un num�ro de test
            If (.Cells(x, Colonne_Des_Test) <> Empty) Then
                'V�rifie si le PR est Prima 2
                If Left(NumPR, 1) = "B" Then
                    .Cells(x, 1) = NumPR & Format(test, "00")
                    NumTest = .Cells(x, 1)
                    NumEtape = 1
                    'ou autres PR
                Else
                    .Cells(x, 1) = NumPR & Format(test, "00")
                    NumTest = .Cells(x, 1)
                    NumEtape = 1
                End If
                
                'Incr�ment du num�ro de test
                test = test + 1
            Else
                .Cells(x, 1) = Empty
            End If
            
            'V�rifie si la cellule de le colonne Com_Etape est vide
            If (.Cells(x, Colonne_Com_Etape) <> Empty) Then
                'Nouveau num�ro d'�tape
                .Cells(x, Colonne_Num_Etape) = NumTest & "-" & Format(NumEtape, "00")
                NumEtape = NumEtape + 1
            Else
                'Sinon recopie le pr�c�dent num�ro d'�tape
                .Cells(x, Colonne_Num_Etape) = .Cells(x - 1, Colonne_Num_Etape)
            End If
            'Incr�ment du num�ro de ligne
            x = x + 1
        Loop
        
    End If
End With


End Function


