Attribute VB_Name = "AddinsManager"
Public AutoUpdate As Boolean
'TODO : la définir à vraie par défaux
Dim barreMacro As IRibbonUI


' Ouvre le fichier de l'historique des évolutions
Sub ShowHistory(control As IRibbonControl)
    Dim myShell As Object
    Set myShell = CreateObject("WScript.Shell")
    myShell.Run Chr(34) & serverPath & "\Historique_evolutions_macro.pdf" & Chr(34)
End Sub


' Met à jour la valeur de AutoUpdate à partir du checkBox de la barre de Addin
Sub ToggleMAJAuto(control As IRibbonControl, majauto As Boolean)
    AutoUpdate = majauto
    works = UpdateValue("AutoUpdate", IIf(AutoUpdate = "Vrai", "True", "False"))
    'Call CheckAutoUpdateBat
    Call RefreshRibbon
End Sub

'initie la valeur de mise à jour depuis le fichier de configuration
Sub InitMajAuto(control As IRibbonControl, ByRef majauto)
    AutoUpdateStr = GetValue("AutoUpdate")
    If AutoUpdateStr = "Error" Then
        Call AddValue("AutoUpdate", "True")
        AutoUpdateStr = GetValue("AutoUpdate")
    End If
    AutoUpdate = AutoUpdateStr
    majauto = AutoUpdate
    Call CheckAutoUpdateBat
    Call RefreshRibbon
End Sub

Sub CheckAutoUpdateBat()
    'Copie ou efface le fichier .bat dans le startup
    Call Shell("cmd /c " & MacroPath & "\UpdateMacroTCMS.exe checkStartup " & AutoUpdate, vbHide)
End Sub

' MAJ manuelle
Sub UpdateManual(control As IRibbonControl)
    'demande à l'utilisateur s'il veut sauvegarder ou pas
    choice = MsgBox("Cette action va fermer Excel !" & vbCrLf & "Voulez vous enregistrer tous les fichiers ouverts ?", vbExclamation + vbYesNoCancel, "Attention")
    If choice = vbYes Then
        'Enregistrer tous les fichiers
        On Error Resume Next
        For Each Wbk In Workbooks
          Wbk.Save
        Next Wbk
        On Error GoTo 0
    End If
    'Si l'utilisateur n'a pas annulé, on lance le .bat qui appelle la MAJ
    If Not choice = vbCancel Then
        Call Shell("cmd /c " & Chr(34) & MacroPath & "\UpdateMacroTCMS.exe" & Chr(34) & " manuel " & macroVersion, vbNormalFocus)
    End If
    
End Sub

' MAJ forcée depuis la version sur le serveur
Sub ForceUpdate(control As IRibbonControl)
    'demande à l'utilisateur s'il veut sauvegarder ou pas
    choice = MsgBox("Cette action va fermer Excel !" & vbCrLf & "Voulez vous enregistrer tous les fichiers ouverts ?", vbExclamation + vbYesNoCancel, "Attention")
    If choice = vbYes Then
        'Enregistrer tous les fichiers
        On Error Resume Next
        For Each Wbk In Workbooks
          Wbk.Save
        Next Wbk
        On Error GoTo 0
    End If
    'Si l'utilisateur n'a pas annulé, on lance le .bat qui appelle la MAJ
    If Not choice = vbCancel Then
        Call Shell("cmd /c " & Chr(34) & serverPath & "\install_auto_macro_alstom_tcms_prima.exe" & Chr(34), vbHide) ' Chr(34) = quote
    End If
    
End Sub

Sub CallbackGetVisible(control As IRibbonControl, ByRef visible)
    visible = Not AutoUpdate
End Sub

Sub SetVersion(control As IRibbonControl, ByRef label)
    label = "Version: " & macroVersion
End Sub

Sub SetUpdateDate(control As IRibbonControl, ByRef label)
    label = "Date MaJ: " & macroUpdateDate
End Sub

Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set barreMacro = ribbon
End Sub

Sub RefreshRibbon()
    If Not barreMacro Is Nothing Then
        barreMacro.Invalidate
    End If
End Sub


' Elimine les macros internes venant du fichier d'exemple du PR
' XL2007: Options Excel/Centre de gestion de la confidentialité/Paramètres du centre de gestion...
' et cocher Accès approuvé au modèle d'objet du projet VBA
Sub deleteOldMacrosFromXLSM()
    Dim moduleToDelete As Variant
    Dim i As Integer
    
    moduleToDelete = Array("General", "Generation_Onglets", "Num_Automatique", "Synthese", "verif_chemin_VS_type", "Verif_prima_II", "VersAncienFormalisme", "VersNouveauFormalisme")
    
    'Boucler sur les Modules à supprimer et voir si le fichier actif les contient
    'S'il y sont, les supprimer
    With ActiveWorkbook.VBProject
    For i = 1 To UBound(moduleToDelete)
        For Each d In .VBComponents
          Debug.Print d.Name
          If d.Type = vbext_ct_StdModule Then
               .VBComponents.Remove .VBComponents(d.Name)
          End If
        Next d
    Next i
    End With
End Sub

Sub UpdateAddIn()
    Dim fs As Object
    Dim Profile As String

    If Workbooks.Count = 0 Then Workbooks.Add
    Profile = Environ("userprofile")
    Set fs = CreateObject("Scripting.FileSystemObject")
    AddIns("MyAddIn").Installed = False
    Call ClearAddinList
    fs.CopyFile "\\SourceOfLatestAddIn\MyAddIn.xla", Profile & "\Application Data\Microsoft\AddIns\", True
    AddIns.Add Profile & "\Application Data\Microsoft\AddIns\MyAddIn.xla"
    AddIns("MyAddIn").Installed = True
End Sub

Sub ClearAddinList()
    Dim MyCount As Long
    Dim GoUpandDown As String

    'Turn display alerts off so user is not prompted to remove Addin from list
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Do
        'Get Count of all AddIns
        MyCount = Application.AddIns.Count

        'Create string for SendKeys that will move up & down AddIn Manager List
        'Any invalid AddIn listed will be removed
        GoUpandDown = "{Up " & MyCount & "}{DOWN " & MyCount & "}"
        Application.SendKeys GoUpandDown & "~", False
        Application.Dialogs(xlDialogAddinManager).Show
    Loop While MyCount <> Application.AddIns.Count

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

