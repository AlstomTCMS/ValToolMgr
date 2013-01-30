Attribute VB_Name = "UI"


Sub Ancien_Vers_Nouveau(control As IRibbonControl)
    If HasActiveBook Then
        Select Case SelectedLayoutVersion
            Case LAYOUT_2012
                ValToolFunctions_2012.AncienVersNouveau
            Case LAYOUT_2013
                ValToolFunctions_2013.AncienVersNouveau
        End Select
    End If
End Sub

' Génère les onglets de test à partir de la synthèse
Sub Generer_Onglets_Tests(control As IRibbonControl)
    If HasActiveBook Then
        Select Case SelectedLayoutVersion
            Case LAYOUT_2012
                ValToolFunctions_2012.Generer_OngletsTests
            Case LAYOUT_2013
                ValToolFunctions_2013.Generer_OngletsTests
        End Select
    End If
End Sub

Sub Reverse_Nvo_Vers_Ancien(control As IRibbonControl)
    If HasActiveBook Then
        Select Case SelectedLayoutVersion
            Case LAYOUT_2012
                ValToolFunctions_2012.Reverse_NvoVersAncien
            Case LAYOUT_2013
                ValToolFunctions_2013.Reverse_NvoVersAncien
        End Select
    End If
End Sub

Sub AddNewPR(control As IRibbonControl)
    Select Case SelectedLayoutVersion
        Case LAYOUT_2012
            ValToolFunctions_2012.CopyRef
        Case LAYOUT_2013
            ValToolFunctions_2013.CopyRef
    End Select
End Sub
