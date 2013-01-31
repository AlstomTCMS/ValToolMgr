Attribute VB_Name = "FromUI"


Sub Generer_OngletsTests()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub Ancien_Vers_Nouveau()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub Reverse_Nvo_Vers_Ancien()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
    
    'Call Generate_scenario
End Sub

Sub AddNewStep()
    With ActiveSheet
        ' détecter si c'est bien un onglet de test au bon format
        ' sortir avec message sinon
        
        'version bidon
        If .Name Like PR_TEST_PREFIX & "*" Then
        
            testNumber = Split(.Name, "_")(1)
            ' Ajouter une colonne à chaque tableau
            .ListObjects("TableAction" & testNumber).ListColumns.Add
            .ListObjects("TableCheck" & testNumber).ListColumns.Add
            .ListObjects("TableDesc" & testNumber).ListColumns.Add
        
        Else
            MsgBox "This sheet is not a PR test. You cannot use this function on this sheet."
        End If
    End With
End Sub
