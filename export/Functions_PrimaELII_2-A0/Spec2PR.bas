Attribute VB_Name = "Spec2PR"
Sub UnmaskRequirements()
    
    'Afficher les textes cachés
    ActiveWindow.ActivePane.View.ShowAll = True 'Not ActiveWindow.ActivePane.View.ShowAll
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=36, Extend:=wdExtend
    With Selection.Font
        .Hidden = False
    End With
    Selection.Copy
End Sub
