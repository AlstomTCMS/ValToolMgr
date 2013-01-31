Attribute VB_Name = "FromUI"


Sub Generer_OngletsTests()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub Ancien_Vers_Nouveau()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub Reverse_Nvo_Vers_Ancien()
    If isActivesheet_a_PR_Test Then
        With ActiveSheet
            Call Generate_scenario
        End With
    End If
End Sub

Sub AddNewStep()
    If isActivesheet_a_PR_Test Then
        With ActiveSheet
            testNumber = Split(.Name, "_")(1)
            ' Ajouter une colonne à chaque tableau
            .ListObjects("TableAction" & testNumber).ListColumns.Add
            .ListObjects("TableCheck" & testNumber).ListColumns.Add
            .ListObjects("TableDesc" & testNumber).ListColumns.Add
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
