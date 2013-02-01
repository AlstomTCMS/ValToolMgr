Attribute VB_Name = "FromUI"


Sub Generer_OngletsTests()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub Ancien_Vers_Nouveau()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub Reverse_Nvo_Vers_Ancien()
    If isActivesheet_a_PR_Test Then
        Call Generate_scenario(getTestNumber)
    End If
End Sub

Sub AddNewStep()
    If isActivesheet_a_PR_Test Then
        With ActiveSheet
            testNumber = getTestNumber
            ' Ajouter une colonne à chaque tableau
            Set actionT = .ListObjects(PR_TEST_TABLE_ACTION_PREFIX & testNumber).ListColumns
            Set checkT = .ListObjects(PR_TEST_TABLE_CHECK_PREFIX & testNumber).ListColumns
            Set descT = .ListObjects(PR_TEST_TABLE_DESCRIPTION_PREFIX & testNumber).ListColumns
            
            
            ' Si tous les tableaux ont la meme taille
            If actionT.Count = checkT.Count And actionT.Count = descT.Count + 1 Then
                actionT.Add
                checkT.Add
                descT.Add
            ElseIf False Then
                stepNumber = actionT.Count
                If stepNumber = checkT.Count Then
                    checkT.Add
                Else
                    'checkT.Resize Range("$B$15:$U$15")
                End If
                
                If stepNumber = descT.Count + 1 Then
                    descT.Add
                Else
                    'descT.Resize Range("$B$15:$U$15")
                End If
            Else
                MsgBox "All tables are not at the same size"
            End If
                       
            
        End With
    End If
End Sub

' Détecter si c'est bien un onglet de test au bon format
' sortir avec message sinon
'version bidon
Function isActivesheet_a_PR_Test(Optional ByVal displayMsg As Boolean = True) As Boolean
    If ActiveSheet.Name Like PR_TEST_PREFIX & "*" Then
        isActivesheet_a_PR_Test = True
    Else
        isActivesheet_a_PR_Test = False
        If displayMsg Then
            MsgBox "This sheet is not a PR test. You cannot use this function on this sheet."
        End If
    End If
End Function


Function getTestNumber() As String
    getTestNumber = Split(ActiveSheet.Name, "_")(1)
End Function
