VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScenarioTs301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IsMaster As Variant
Private instance(1000, 5)
Private check(1000, 1)
Private Numcol
Private col_projet As Variant

Private Sub Class_Initialize()
    MsgBox "Initialized!"
End Sub

Private Sub Class_Terminate()
    MsgBox "Terminated!"
End Sub

Public Sub writeBegin()
    MsgBox "write Beginning!"
End Sub

Public Sub writeSheet(Sheet As Worksheet)
    MsgBox "write sheet!"
End Sub

Public Sub writeEnd()
    MsgBox "write end!"
End Sub

Private Sub Generation_TestStand()
    Dim Acronym As String
    Dim AcronymOk As Boolean
    Dim sFileseq As String
    Dim Compteur As Long
    Dim EnsembleLignes0 As New Collection
    Dim EnsembleLignes1 As New Collection
    Dim tableau_instance(1000, 1000) As String
    Dim proj_conf() As Variant     'parametre valeur type
    Dim teststand_conf() As String      'parametre valeur type
    Dim proj_filglo() As String    'parametre valeur type
    Dim proj_inst() As String      'parametre valeur
    Dim Path As String
    Dim k As Integer

    'optimisation excel
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    Application.DisplayAlerts = False

    'teste de génération
    If LCase(ActiveWorkbook.ActiveSheet.Name) = "param" Or LCase(ActiveWorkbook.ActiveSheet.Name) = "sousfonct" Or _
        LCase(ActiveWorkbook.ActiveSheet.Name) = "history" Or LCase(ActiveWorkbook.ActiveSheet.Name) = "rapport" Or _
        LCase(ActiveWorkbook.ActiveSheet.Name) = "outils" Or LCase(ActiveWorkbook.ActiveSheet.Name) = "context_error" Or _
        LCase(ActiveWorkbook.ActiveSheet.Name) = "events_error" Or LCase(ActiveWorkbook.ActiveSheet.Name) = "variable_cont" Or _
        LCase(ActiveWorkbook.ActiveSheet.Name) = "parameters" Or LCase(ActiveWorkbook.ActiveSheet.Name) = "tabconv" Then
        MsgBox ("This sheet can not be generated.")
        End
    End If

    'gestion config
    For k = 1 To 100
        If LCase(Worksheets("Parameters").Cells(k, 2)) Like "sections number" Then
            NbreSection = Worksheets("Parameters").Cells(k, 4)
            Exit For
        End If
        Next
        For k = 1 To 100
            If LCase(Worksheets("Parameters").Cells(k, 2)) Like "environments number" Then
            NbreEnv = Worksheets("Parameters").Cells(k, 4)
            Exit For
        End If
    Next

    Embedded = False
    i = 4
    j = 2

    'parametres
    col_desc = 9
    seq_nb_senar = 35000 'nombre de senarios dans le fichier teststand avant le découpage (découpage a la fin du senario)

    'Initialisation
    Compteur = 0

    'Par défaut, le nom de la fonction à traiter est celui de l'onglet actif
    Acronym = ActiveWorkbook.ActiveSheet.Name
    sFileseq = Acronym & ".seq"

    If LCase(Acronym) Like "**slave**" Then
        IsMaster = False
    Else
        IsMaster = True
    End If

    'Barre d'attente
    F_BarreAttente.Show
    Call BarreAttente("En cours", "Analyse de la configuration")

    'lecture de la configuration
    proj = lecture_conf(proj_conf, teststand_conf, proj_filglo, proj_inst)

    'mise en page et renumerotation
    Sheets("SousFonct").Select
    Call Renumerotation_sousstep_TestStand

    Sheets(Acronym).Select
    Range(Cells(1, 1), Cells(65000, 1)) = ""
    Call Renumerotation_step_TestStand

    'lecture chemin
    directory = lecture_param(teststand_conf, "Generation_directory")
    If directory = "" Then directory = ThisWorkbook.Path

    'lecture automatic restart
    'Auto_restart = LCase(lecture_param(teststand_conf, "Automatic restart"))

    'Barre d'attente
    Call BarreAttente("En cours", "Génération du Teststand de " & Acronym)

    'decoupage du fichier teststand
    file_num = 1


    Numcol = 4
    k = 1
    'debut
    While Cells(i, j) <> "END"
    'If Auto_restart <> "" Then
    '    If LCase(Cells(i, j)) = "restart" And Auto_restart = "false" Then
    '        Call TraitementSequenceCall("Restart", EnsembleLignes0, EnsembleLignes1, Compteur, "Restart", "CurFile", proj_filglo, proj_conf, NSection)
    '        i = i + 1
    '    ElseIf Auto_restart = "true" Then Call TraitementSequenceCall("Restart", EnsembleLignes0, EnsembleLignes1, Compteur, "Restart", "CurFile", proj_filglo, proj_conf, NSection): i = i + 1
    '    End If
    'End If

    If Cells(i, j) <> "" Then
    If LCase(Cells(i, j)) = "nominal cases" Or LCase(Cells(i, j)) = "degraded cases" Or LCase(Cells(i, j)) = "end" Or LCase(Cells(i, j)) = LCase(proj) Then
    If LCase(Cells(i, j)) = LCase(proj) Then
    Call TraitementLabel("Specific senario for " & proj, EnsembleLignes0, EnsembleLignes1, Compteur)
    Else
    Call TraitementLabel(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur)
    End If
    i = i + 1
    Else
    Do While LCase(Cells(i, j)) <> "nominal cases" And LCase(Cells(i, j)) <> "degraded cases" And LCase(Cells(i, j)) <> "end" And LCase(Cells(i, j)) <> LCase(proj)
    i = i + 1
    Loop
    i = i - 1
    End If
    End If
    If Cells(i, j + 1) <> "" Then
    Call recuperation_instance(tableau_instance, proj_inst)
    Erase instance
    tableau_courant tableau_instance, 1
    l = TraitementExcel_master(Numcol, EnsembleLignes0, EnsembleLignes1, Compteur, i, teststand_conf, proj, proj_filglo, proj_conf)
    calcul_k tableau_instance, k
    Cells(i, 1) = k
    For t = 2 To k
    tableau_courant tableau_instance, k
    l = TraitementExcel_master(Numcol, EnsembleLignes0, EnsembleLignes1, Compteur, i, teststand_conf, proj, proj_filglo, proj_conf)
    Next t
    i = l + 1
    Else
    i = i + 1
    End If

    If Compteur > seq_nb_senar And (Cells(i, 2) <> "END" And Cells(i + 1, 2) <> "END" And Cells(i + 2, 2) <> "END") Then
    Call write_teststand(fic, sFileseq, Compteur, EnsembleLignes0, EnsembleLignes1, proj_filglo, proj_conf, teststand_conf, directory, file_num, False)
    End If
    Wend

    'Recuperation des evenements
    'For k = 1 To NbreSection
    '    Path = "C:\Teststand\Download_Event" & k & ".seq"
    '    Path = Replace(Path, "\", "\\")
    '    Call TraitementSequenceCall("Download_Events", EnsembleLignes0, EnsembleLignes1, Compteur, "MainSequence", Path, proj_filglo, proj_conf, "_1")
    'Next k

    'Barre d'attente
    Call BarreAttente("En cours", "Ecriture du SEQ de " & Acronym)

    Call write_teststand(fic, sFileseq, Compteur, EnsembleLignes0, EnsembleLignes1, proj_filglo, proj_conf, teststand_conf, directory, file_num, True)

    'Barre d'attente
    Unload F_BarreAttente
    F_BarreAttente.Hide

    'optimisation excel
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    If Sheet(0) = "" Then MsgBox "The TestStand file has correctly been generated", vbOKOnly: ActiveWorkbook.Save

End Sub

Function write_teststand(fic As Variant, sFileseq As String, Compteur As Long, EnsembleLignes0, EnsembleLignes1, proj_filglo, proj_conf, teststand_conf, directory, ByRef file_num As Variant, fin As Boolean)

    fic = FreeFile

    If fin = False Then  'si file_num=0 dernier fichier
    If file_num > 1 And file_num < 10 Then
    nextseq = Left(sFileseq, Len(sFileseq) - 5) & file_num & ".seq"
    ElseIf file_num < 10 Then
    nextseq = Left(sFileseq, Len(sFileseq) - 4) & file_num & ".seq"
    Else
    nextseq = Left(sFileseq, Len(sFileseq) - 6) & file_num & ".seq"
    End If
    Call TraitementSequenceCall("Call " & nextseq, EnsembleLignes0, EnsembleLignes1, Compteur, "MainSequence", nextseq, proj_filglo, proj_conf, NSection)
    End If

    'Si le fichier existe ,on le renomme en .old
    If Dir(directory & "\" & sFileseq) <> vbNullString Then
    FileSystem.FileCopy directory & "\" & sFileseq, directory & "\" & sFileseq & ".old"
    FileSystem.Kill directory & "\" & sFileseq
    End If

    Open directory & "\" & sFileseq For Output As #fic
    Call CreationFicSeq(fic, sFileseq, Compteur, EnsembleLignes0, EnsembleLignes1, proj_filglo, proj_conf, teststand_conf, directory)
    Close #fic

    sFileseq = nextseq

    'vider les donnée ecrites
    Compteur = 0
    Set EnsembleLignes0 = Nothing
    Set EnsembleLignes0 = New Collection
    Set EnsembleLignes1 = Nothing
    Set EnsembleLignes1 = New Collection

    'incremente le fichier
    file_num = file_num + 1

    'passe le fichier en slave pour la suite du senario
    'embedded = True
    IsMaster = False

End Function

Function TraitementExcel_master(Numcol, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, ByVal i, teststand_conf, proj, proj_filglo, proj_conf) As Variant

    'lecture automatic restart
    'Auto_restart = LCase(lecture_param(teststand_conf, "Automatic restart"))

    Dim j As Integer
    j = 2

    'Call TraitementLabel("****" & (Cells(i, j + 1) & " -- " & Cells(i, j + 2)), EnsembleLignes0, EnsembleLignes1, Compteur)
    ' redemarrage de controlbuild
    If Compteur > 2 Then
        If LCase(Cells(i, j)) = "restart" Then Call TraitementSequenceCall("Restart", EnsembleLignes0, EnsembleLignes1, Compteur, "Restart", "CurFile", proj_filglo, proj_conf, NSection): i = i + 1
    End If

    While UCase(Cells(i, j + 2)) <> "END"
    If Cells(i, j + 1) Like "*_N_*" Or Cells(i, j + 1) Like "*_D_*" Then Call TraitementLabel("****" & (Cells(i, j + 1) & " -- " & Cells(i, j + 2)), EnsembleLignes0, EnsembleLignes1, Compteur)
    'i = i - 2
    ' redemarrage de controlbuild
    If LCase(Cells(i, j)) = "restart" Then
        Call TraitementSequenceCall("Restart", EnsembleLignes0, EnsembleLignes1, Compteur, "Restart", "CurFile", proj_filglo, proj_conf, NSection)
        i = i + 1
    End If

    If Cells(i, j) = "END" Then
        MsgBox ("Borne END du senario non trouvée")
        End
    End If

    If Cells(i, j + 2) <> "" Then
    Call TraitementLabel((Cells(i, j + 2) & " -- " & Cells(i, j + Numcol)), EnsembleLignes0, EnsembleLignes1, Compteur)
    k = 1
    
    While Cells(i + k, j + 2) = ""
    If Cells(i + k, j + Numcol) <> "" Then
    DoEvents
    Call TraitementAction(EnsembleLignes0, EnsembleLignes1, Compteur, Cells(i + k, j + Numcol), i + k, j + Numcol, proj_filglo, proj_conf)
    End If
    If Cells(i + k, j) = "END" Then
    MsgBox ("Borne END du senario non trouvée")
    End
    End If
    k = k + 1
    Wend

    If Cells(i, j + Numcol + 5) <> "" Then
    Call TraitementLabel((Cells(i, j + 2) & " -- " & Cells(i, j + Numcol + 5)), EnsembleLignes0, EnsembleLignes1, Compteur)
    End If
    k = 1
    While Cells(i + k, j + 2) = ""
    If Cells(i + k, j + Numcol + 5) <> "" Then
    DoEvents
    Call TraitementAction(EnsembleLignes0, EnsembleLignes1, Compteur, Cells(i + k, j + Numcol + 5), i + k, j + Numcol + 5, proj_filglo, proj_conf)
    End If
    If Cells(i + k, j) = "END" Then
    MsgBox ("Borne END du senario non trouvée")
    End
    End If
    k = k + 1
    Wend
    If Cells(i, j + 1) = "" And k > 2 Then k = k - 2
    i = i + k
    Else
    i = i + 1
    End If
    'i = i + 1
    Wend

    If Cells(i, j) = "END" Then
    MsgBox ("Borne END du senario non trouvée")
    End
    End If
    TraitementExcel_master = i

End Function

Private Sub TraitementExcel_sousseq(Numcol, EnsembleLignes0, EnsembleLignes1, sous_fonct, ByRef Compteur)

    Dim i, j As Integer

    i = 4
    j = 2

    While Cells(i, j) <> "END"
    While Cells(i, j + 1) = sous_fonct
    If Cells(i, j + 1) <> "" Then
    Call TraitementLabel("______" & (Cells(i, j + 1) & " -- " & Cells(i, j + 2)), EnsembleLignes0, EnsembleLignes1, Compteur)
    i = i + 1
    While UCase(Cells(i, j + 2)) <> "END"
    If Cells(i, j + 2) <> "" Then
    Call TraitementLabel("______" & Cells(i, j + 2) & " -- " & (Cells(i, j + Numcol)), EnsembleLignes0, EnsembleLignes1, Compteur)
    k = 1
    While Cells(i + k, j + 2) = ""
    If Cells(i + k, j + Numcol) <> "" Then
    DoEvents
    Call TraitementAction(EnsembleLignes0, EnsembleLignes1, Compteur, Cells(i + k, j + Numcol), i + k, j + Numcol, proj_filglo, proj_conf)
    End If
    k = k + 1
    Wend

    If Cells(i, j + Numcol + 5) <> "" Then
    Call TraitementLabel("______" & Cells(i, j + 2) & " -- " & Cells(i, j + Numcol + 5), EnsembleLignes0, EnsembleLignes1, Compteur)
    End If
    k = 1
    While Cells(i + k, j + 2) = ""
    If Cells(i + k, j + Numcol + 5) <> "" Then
    DoEvents
    Call TraitementAction(EnsembleLignes0, EnsembleLignes1, Compteur, Cells(i + k, j + Numcol + 5), i + k, j + Numcol + 5, proj_filglo, proj_conf)
    End If
    k = k + 1
    Wend

    i = i + k
    End If
    Wend
    Exit Private Sub
    i = i + 1

    End If

    Wend
    i = i + 1
    Wend

End Sub

Private Sub TraitementAction(EnsembleLignes0, EnsembleLignes1, ByRef Compteur, MotAction As String, ByVal i As Integer, ByVal j As Integer, proj_filglo, proj_conf)
    Dim strLib As String
    Dim variable As String
    Dim NSection As String
    Dim Location As String

    'traitement du nom de la variable
    If Not LCase(Cells(i, j)) Like "attrib" And InStr(1, Cells(i, j + 1), "<") <> 0 Then
    variable = Cells(i, j + 1)
    Do While variable Like "*<*>*"
    temp = Mid(variable, InStr(1, variable, "<"), InStr(1, variable, ">") + 1 - InStr(1, variable, "<"))
    X = 0
    Do While instance(X, 0) <> "" Or instance(X + 1, 0) <> ""
    If temp = instance(X, 0) Then
    If Not temp Like "*_*" Then instance(X, 2) = 1
    If temp Like "*_*" Then instance(X, 2) = 1
    variable = Replace(variable, instance(X, 0), instance(X, 1))
    Exit Do
    End If
    X = X + 1
    Loop
    If variable Like "*<*>*" And instance(X, 0) = "" Then
    MsgBox ("Instance non trouvée pour la variable " & variable & "   ligne " & i)
    End
    End If
    Loop
    Else
    variable = Cells(i, j + 1)
    End If


    'traitement du nom de la valeur
    'If LCase(Cells(i, j)) <> "attrib" And InStr(1, Cells(i, j + 3), "<") <> 0 Then
    '    valeur = Cells(i, j + 3)
    '    Do While valeur Like "**<**>**"
    '        temp = Mid(valeur, InStr(1, valeur, "<"), InStr(1, valeur, ">") + 1 - InStr(1, valeur, "<"))
    '        X = 0
    '            Do While instance(X, 0) <> "" Or instance(X + 1, 0) <> ""
    '                If temp = instance(X, 0) Then
    '                    If Not temp Like "*_*" Then instance(X, 2) = 1
    '                    If temp Like "*_*" Then instance(X, 2) = 2
    '                    valeur = Replace(valeur, instance(X, 0), instance(X, 1))
    '                    Exit Do
    '                End If
    '                X = X + 1
    '            Loop
    '    Loop
    'Else
    'valeur = Cells(i, j + 3)


    'pas = 0
    '
    'If LCase(Cells(i, j)) Like "attrib" And InStr(1, Cells(i, j + 1), "<") <> 0 Then
    '    Do While Not variable Like Left(instance(pas, 0), Len(instance(pas, 0)) - 1) & "*"
    '        pas = pas + 1
    '        If instance(pas, 0) = "" Then Exit Do
    '    Loop
    'End If

    'If InStr(1, Cells(i, j + 1), "<") <> 0 And valeur Like "*,*" Then
    '    instance_number = instance(pas, 3) - 1
    'Else
    '    instance_number = 1
    'End If

    'If LCase(Cells(i, j)) Like "attrib" Then instance(pas, 5) = nb_inst(valeur)
    '    If valeur Like "*,*" Then
    '        debut = 0
    '        temp = valeur
    '        fin = 0
    '       If instance_number <> 1 Then
    '            For w = 1 To instance_number - 1
    '            debut = debut + InStr(1, temp, ",")
    '            temp = Right(valeur, Len(valeur) - debut)
    '            Next w
    '        End If
    '        If temp Like "*,*" Then
    '            Do While valeur Like "*,*"
    '                fin = fin + InStr(1, temp, ",")
    '                valeur = Left(temp, fin - 1)
    '            Loop
    '        Else
    '            valeur = temp
    '        End If
    '    Else
    '        fin = Len(valeur)
    '    End If
    'End If

    'changement des reels avec "," en reels avec "."
    valeur = Cells(i, j + 3)
    If valeur Like "*,*" Then
    w = InStr(1, valeur, ",")
    fin = Len(valeur)
    temp = Left(valeur, w - 1)
    temp = temp & "." & Right(valeur, fin - w)
    valeur = temp
    End If

    'ecriture du label teststand avec le nom de variable
    strLib = Cells(i, j) & ";" & variable & ";" & valeur
    localise = Cells(i, j + 2)

    'traitement de la section
    If Int(Cells(i, j + 4)) > NbreSection Then
    MsgBox ("Section non trouvée pour la variable " & variable & "   ligne " & i)
    End
    End If

    ' Location (ENV_1=loco1&2, ENV_2=loco3&4)
    NSection = Cells(i, j + 4)
    Select Case Cells(i, j + 2)
    Case "Env"   'Pour le test des Rioms
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "Environment"
    Case "Sys"   'Pour le test des Rioms
    If Cells(i, j + 4) < 3 Then
    NSection = "1"
    Else
    NSection = "2"
    End If
    strLib = "Entree Riom" & ";" & strLib
    Location = "Application_TrTrans/Embedded/System"
    Case "ACU11"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "ACU12"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "TCU1"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "TCU2"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "TCU3"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "TCU4"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "BCU"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "LV_LOCO"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "BEHAVIOR"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "DESK1"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case "ROOF"
    If Cells(i, j + 4) < 3 Then
    NSection = "_1"
    Else
    NSection = "_2"
    End If
    strLib = "EnvSection" & Cells(i, j + 4) & ";" & strLib
    Location = "environment/env_Section" & Cells(i, j + 4) & "/taskenv_sil0_Section" & Cells(i, j + 4) & ""
    Case ""
    Case Else
    For k = 2 To 500
    temp = ""
    If Cells(i, j + 2) Like "*/*" Then
    temp = Mid(Cells(i, j + 2), InStr(Cells(i, j + 2), "/"), Len(Cells(i, j + 2)))
    Location = Mid(Cells(i, j + 2), 1, InStr(Cells(i, j + 2), "/") - 1)
    Else
    Location = Cells(i, j + 2)
    End If
    If Worksheets("TabConv").Cells(k, 2) Like Location Then
    Location = "Application_TrTrans/Embedded/Control_Command/FBS/" & Worksheets("TabConv").Cells(k, 3) & temp
    Exit For
    End If
    Next
    If Worksheets("TabConv").Cells(k, 2) = "" Then
    MsgBox ("Unexpected location for the variable " & variable & "   row " & i)
    End
    End If
    End Select


    ' path management for callsequence
    Path = ThisWorkbook.Path & "\" & variable & ".seq"
    Path = Replace(Path, "\", "\\")


    Select Case UCase(MotAction)
    Case UCase("LoadParamFile")
    Call TraitementCB_LoadParamFile(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location)
    Case UCase("MsgPopUp")
    Call TraitementCB_MsgPopUp(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, valeur)
    Case UCase("Force")
    Call TraitementCB_Force(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, valeur, NSection)
    Case UCase("ForceArrayElt")
    Call TraitementCB_ForceArrayElt(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, CInt(Cells(i, j + 4)), valeur)
    Case UCase("UnForce")
    Call TraitementCB_UnForce(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, NSection)
    Case UCase("Write")
    Call TraitementCB_Write(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, valeur, NSection)
    Case UCase("WriteArrayElt")
    Call TraitementCB_WriteArrayElt(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, valeur, Cells(i, j + 4))
    Case UCase("Read")
    Call TraitementCB_Read(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, valeur, NSection)
    Case UCase("Test")
    Call TraitementCB_Test(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, valeur, NSection, IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)))
    Case UCase("TestDT")
    Call TraitementCB_TestDT(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, Cells(i, j + 4), valeur, NSection)
    Case UCase("TestBool")
    Call TraitementCB_TestBool(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, Cells(i, j + 3), IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)), NSection)
    Case UCase("TestAna")
    Call TraitementCB_TestAna(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, valeur, Cells(i, j + 4), Cells(i, j + 6), IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)))
    Case UCase("TestArrayElt")
    Call TraitementCB_TestArrayElt(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Location, Cells(i, j + 4), valeur, , , IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)))
    Case UCase("ForceNN")
    Call TraitementCB_ForceNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, NSection)
    Case UCase("UnForceNN")
    Call TraitementCB_UnForceNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, NSection)
    Case UCase("WriteNN")
    Call TraitementCB_SetNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, NSection)
    Case UCase("ReadNN")
    Call TraitementCB_ReadNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, NSection)
    Case UCase("TestNN")
    Call TraitementCB_TestNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, NSection, IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)))
    Case UCase("TestBoolNN")
    Call TraitementCB_TestBoolNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, NSection, IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)))
    Case UCase("TestAnaNN")
    Call TraitementCB_TestAnaNN(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, NSection, Cells(i, j + 4), Cells(i, j + 6), IIf(Cells(i, j + 5) = "", "FileGlobals.TimeOut", Cells(i, j + 5)))
    Case UCase("Call")
    Call TraitementSequenceCall(Cells(i, j) & ";" & variable, EnsembleLignes0, EnsembleLignes1, Compteur, variable, Path, proj_filglo, proj_conf, NSection)
    Case UCase("Wait")
    Call TraitementNI_Wait(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, valeur)
    Case UCase("CB_StopTask")
    Call TraitementCB_StopTask(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur)
    Case UCase("CB_StartTask")
    Call TraitementCB_StartTask(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur)
    Case UCase("TestDDU")
    Call TraitementTestDDU(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur)
    Case UCase("WriteDDU")
    Call TraitementWriteDDU(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur)
    Case UCase("Attrib")
    If j < 9 Then
    Call Traitementinstance(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur)
    Else
    Call Traitementcheck(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur)
    End If
    Case UCase("Run")
    Call Traitement_sousseq(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur, i)
    Case UCase("Statement")
    Call Traitement_Statement(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur, variable, valeur)
    Case UCase("CompareScreenshot")
    Call Traitement_CompareScreenshot(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case UCase("MakeScreenshot")
    Call Traitement_Initialisation(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case UCase("KeyboardControl")
    Call Traitement_KeyboardControl(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case UCase("MouseControl")
    Call Traitement_MouseControl(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Call TraitementNI_Wait("wait", EnsembleLignes0, EnsembleLignes1, Compteur, "1") ''1s attente pour traitement
    Case UCase("OpticalCharacterRecognition")
    Call Traitement_OpticalCharacterRecognition(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case UCase("RemoteControlStart")
    Call Traitement_RemoteControlStart(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case UCase("RemoteControlStop")
    Call Traitement_RemoteControlStop(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case UCase("SendText")
    Call Traitement_SendText(strLib, EnsembleLignes0, EnsembleLignes1, Compteur, variable, localise, valeur)
    Case Else
    Call TraitementLabel(Cells(i, j), EnsembleLignes0, EnsembleLignes1, Compteur)
    End Select

End Sub

Private Sub Traitement_sousseq(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, writeValue As Variant, ligne As Variant)
    Dim X
    Dim feuille As String
    X = 2
    Do While Sheets("param").Cells(X, 1) <> ""
    If Sheets("param").Cells(X, 2) = strNickName Then
    sousseq = Sheets("param").Cells(X, 1)
    Exit Do
    End If
    X = X + 1
    Loop
    If sousseq = "" Then
    MsgBox ("La sous fonction """ & strNickName & """  (ligne " & ligne & ") n'existe pas" & Chr(13) & "Génération annulée")
    End
    End If
    feuille = ActiveSheet.Name
    Sheets("SousFonct").Select
    Call TraitementExcel_sousseq(Numcol, EnsembleLignes0, EnsembleLignes1, sousseq, Compteur)
    Sheets(feuille).Select


End Sub

Private Sub Traitementcheck(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, writeValue As Variant)
    Dim j
    Do While check(j, 0) <> ""
    If check(j, 0) = strNickName Then check(j, 1) = writeValue
    j = j + 1
    Loop

    If check(j, 0) = "" Then check(j, 0) = strNickName
    If check(j, 1) = "" Then check(j, 1) = writeValue

End Sub

Private Sub Traitementinstance(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, writeValue As Variant)
    Dim j
    Do While instance(j, 0) <> ""
    If instance(j, 0) = strNickName Then instance(j, 1) = writeValue
    j = j + 1
    Loop

    If instance(j, 0) = "" Then instance(j, 0) = strNickName
    If instance(j, 1) = "" Then instance(j, 1) = writeValue

End Sub
Private Sub TraitementCB_LoadParamFile(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, FilePath As String, strInstanceName As String)

    Dim TmpLine As String
    Dim TmpArray
    Dim TmpArray2
    Dim VarName As String
    Dim VarType As String
    Dim VarPath As String
    Dim VarVal As Variant
    Dim ligne As String
    Dim IsAnArray As Boolean
    Dim IsATime As Boolean
    Dim IsToSkip As Boolean
    Dim ArrayIndex As Integer
    Dim IsFirstLine As Boolean
    Dim VarTimeDay As Integer
    Dim VarTimeHour As Integer
    Dim VarTimeMin As Integer
    Dim VarTimeSec As Integer
    Dim VarTimeMilSec As Integer
    Dim tmpInt As Long

    Set FS = CreateObject("Scripting.FileSystemObject")
    IsFirstLine = True

    'ouverture du fichier de paramétrage
    If FS.FileExists(FilePath) Then
    Set A = FS.OpenTextFile(FilePath, 1)

    'saute la premiere ligne (code inconnu)
    A.readline

    'parcours du fichier
    While Not A.AtEndOfStream
    IsAnArray = False
    IsATime = False
    IsToSkip = False

    TmpLine = A.readline
    TmpArray = Split(TmpLine, " ")

    '
    ' Variable detection
    '

    VarType = TmpArray(0)
    VarVal = TmpArray(2)

    'cas de la première ligne : la première ligne comporte un integer à ignorer, suivi d'une ligne à traiter
    If IsFirstLine Then
    VarType = Right(VarType, 1)
    IsFirstLine = False
    End If

    'get last occurence of char '/' to retreive path and var name
    TmpArray2 = Split(TmpArray(3), "/")

    VarPath = ""
    For i = 4 To UBound(TmpArray2) - 1
    VarPath = VarPath & "/" & TmpArray2(i)
    Next
    If VarPath <> "" Then
    VarPath = Right(VarPath, Len(VarPath) - 1)
    End If


    VarName = TmpArray2(UBound(TmpArray2))


    'check if the variable is a time element
    If (Left(VarName, 7) = "CFG_XT_" Or _
    Left(VarName, 3) = "PT_" Or _
    Left(VarName, 7) = "CFG_PT_") And _
    InStr(1, VarName, "__") <> 0 Then

    If Right(VarName, 3) = "__1" Then
    IsToSkip = True
    Else
    IsATime = True

    VarName = Left(VarName, Len(VarName) - 3)

    'time formatting
    VarTimeDay = 0
    VarTimeHour = 0
    VarTimeMin = 0
    VarTimeSec = 0
    VarTimeMilSec = 0
    tmpInt = 0

    VarTimeDay = VarVal / 864000000
    tmpInt = VarVal Mod 864000000
    VarTimeHour = (VarVal Mod 864000000) / 36000000
    VarTimeMin = ((VarVal Mod 864000000) Mod 36000000) / 600000
    VarTimeSec = (((VarVal Mod 864000000) Mod 36000000) Mod 600000) / 10000
    VarTimeMilSec = ((((VarVal Mod 864000000) Mod 36000000) Mod 600000)) Mod 10000

    VarVal = """T#" & Format(VarTimeHour, "00") & "h" _
    & Format(VarTimeMin, "00") & "m" _
    & Format(VarTimeSec, "00") & "s" _
    & Format(VarTimeMilSec, "000") & "ms"""
    End If

    ElseIf InStr(1, VarName, "__") <> 0 Then ' check if variable is an array
    IsAnArray = True
    TmpArray2 = Split(VarName, "__")
    VarName = TmpArray2(0)
    ArrayIndex = TmpArray2(1)
    End If

    If VarName = "simuperiod" Then
    IsToSkip = True
    End If

    '
    ' Variable treatment
    '

    'check if variable is to be interpreted
    If Not IsToSkip Then
    If IsAnArray Then 'array element variable
    ligne = "ForceArrayElt" & ";" & VarName & "[" & ArrayIndex & "];" & VarVal
    TraitementCB_ForceArrayElt ligne, EnsembleLignes0, EnsembleLignes1, Compteur, VarName, VarPath, ArrayIndex, VarVal

    ElseIf IsATime Then 'time variable
    ligne = "Write" & ";" & VarName & ";" & VarVal
    Call TraitementCB_Write(ligne, EnsembleLignes0, EnsembleLignes1, Compteur, VarName, VarPath, VarVal, Section)

    Else 'normal variable
    Select Case VarType
    Case "B":
    If CBool(VarVal) Then
    VarVal = "true"
    Else
    VarVal = "false"
    End If
    ligne = "Force" & ";" & VarName & ";" & VarVal
    Call TraitementCB_Force(ligne, EnsembleLignes0, EnsembleLignes1, Compteur, VarName, VarPath, VarVal, Section)
    Case "I":
    ligne = "Force" & ";" & VarName & ";" & CLng(VarVal)
    Call TraitementCB_Force(ligne, EnsembleLignes0, EnsembleLignes1, Compteur, VarName, VarPath, CLng(VarVal), Section)
    Case "C":
    ligne = "Force" & ";" & VarName & ";" & Chr(VarVal)
    Call TraitementCB_Force(ligne, EnsembleLignes0, EnsembleLignes1, Compteur, VarName, VarPath, """" & Chr(VarVal) & """", Section)
    Case "F":
    ligne = "Force" & ";" & VarName & ";" & Replace(VarVal, ",", ".")
    Call TraitementCB_Force(ligne, EnsembleLignes0, EnsembleLignes1, Compteur, VarName, VarPath, Replace(VarVal, ",", "."), Section)
    End Select
    End If
    End If
    Wend
    A.Close
    Set A = Nothing
    Set FS = Nothing
    End If

End Sub

Private Sub Traitement_CompareScreenshot(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Compare"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:O20QV1XzbE+dR5V0CpOOeD"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "intScreenshotX =" & Split(localise, ";")(0)
    EnsembleLignes1.Add "intScreenshotY = " & Split(localise, ";")(1)
    EnsembleLignes1.Add "intScreenshotW = " & Split(localise, ";")(2)
    EnsembleLignes1.Add "intScreenshotH = " & Split(localise, ";")(3)
    EnsembleLignes1.Add "stepname = " & strNickName
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""Compare Screenshot"""

    Compteur = Compteur + 1
End Sub

Private Sub Traitement_Initialisation(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

End Sub

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"                               '''fonction déplacée dans le setup
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Init"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "ID = ""ID#:Un729hMWiE6kx4Jk4M7VyC"" "
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "sImagesPath = " & strNickName
    EnsembleLignes1.Add "bReferenceBuild = " & writeValue
    EnsembleLignes1.Add "sShadowFileConf = " & localise
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""DDU Initialisation"""

    Compteur = Compteur + 1

End Sub

Private Sub Traitement_KeyboardControl(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"            '''''''''''''''''''''''''''''''''pas utilisé
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Keyboard"""        '
    '
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"                    '
    EnsembleLignes1.Add "Id = ""ID#:Mb7So24OmkWZfG7pbA2XFB"""                      '
    EnsembleLignes1.Add ""                                                         '
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"                       '
    EnsembleLignes1.Add "key =" & writeValue                                       '
    EnsembleLignes1.Add ""                                                         '
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"                  '
    EnsembleLignes1.Add "%NAME = ""Keyboard Control"""                             '
    '
    Compteur = Compteur + 1                                                        '
    ''''''''''''''''''''
End Sub

Private Sub Traitement_MouseControl(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Mouse"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:tla0zLYnlkyjOt1x7o1R7A"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, DotNetStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "AssemblyPath = ""C:\\Program Files\\Alstom\\TCMS Testbench\\TestHMI\\TestHMI.dll"""
    EnsembleLignes1.Add "ClassVariable = ""FileGlobals.TH"""
    EnsembleLignes1.Add "ClassName = ""TestHMI.TestHMI"""
    EnsembleLignes1.Add "Constructor = ""TestHMI()"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "MemberName = ""MouseControl"""
    EnsembleLignes1.Add "%HI: Params = [8]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params]"
    EnsembleLignes1.Add "%[0] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[6] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[7] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add "%[8] = ""TYPE, DotNetParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[0]]"
    EnsembleLignes1.Add "Name = ""bPassFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 2"
    EnsembleLignes1.Add "Direction = 10"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[1]]"
    EnsembleLignes1.Add "Name = ""sReportText"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 1"
    EnsembleLignes1.Add "Direction = 10"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[2]]"
    EnsembleLignes1.Add "Name = ""num_DDU"""
    EnsembleLignes1.Add "ArgVal = ""Step.num_DDU"""
    EnsembleLignes1.Add "Type = 6"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[3]]"
    EnsembleLignes1.Add "Name = ""ev"""
    EnsembleLignes1.Add "ArgVal = "" \ ""MouseClick\"""""         '''possibilité
    EnsembleLignes1.Add "Type = 16"
    EnsembleLignes1.Add "DisplayType = ""Enum (TestHMI.EnumMouseEvent)"""
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[4]]"
    EnsembleLignes1.Add "Name = ""but"""
    EnsembleLignes1.Add "ArgVal = "" \ ""MouseLeft\"""""          '''possibilité
    EnsembleLignes1.Add "Type = 16"
    EnsembleLignes1.Add "DisplayType = ""Enum (TestHMI.EnumMouseButton)"""
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[5]]"
    EnsembleLignes1.Add "Name = ""iMouseX"""
    EnsembleLignes1.Add "ArgVal = ""Step.iMouseX"""
    EnsembleLignes1.Add "Type = 6"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[6]]"
    EnsembleLignes1.Add "Name = ""iMouseY"""
    EnsembleLignes1.Add "ArgVal = ""Step.iMouseY"""
    EnsembleLignes1.Add "Type = 6"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[7]]"
    EnsembleLignes1.Add "Name = ""iMouseCount"""
    EnsembleLignes1.Add "ArgVal = ""Step.iMouseCount"""
    EnsembleLignes1.Add "Type = 6"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Params[8]]"
    EnsembleLignes1.Add "Name = ""iMouseWait"""
    EnsembleLignes1.Add "ArgVal = ""Step.iMouseWait"""
    EnsembleLignes1.Add "Type = 6"
    EnsembleLignes1.Add "Dimension = -1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "iMouseX = " & Split(localise, ";")(0)
    EnsembleLignes1.Add "iMouseY = " & Split(localise, ";")(1)
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""Mouse Control"""


    Compteur = Compteur + 1

End Sub

Private Sub Traitement_OpticalCharacterRecognition(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"                    ''''''''''''''''''''''''''pas utilisé
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_OCR"""                     '
    '
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"                            '
    EnsembleLignes1.Add "Id = ""ID#:Jlo1ZqpU+kaeS199ohP32D"""                              '
    EnsembleLignes1.Add ""                                                                 '
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"                               '
    EnsembleLignes1.Add "iOcrX = " & Split(localise, ";")(0)                               '
    EnsembleLignes1.Add "iOcrY = " & Split(localise, ";")(1)                               '
    EnsembleLignes1.Add "iOcrW = " & Split(localise, ";")(2)                               '
    EnsembleLignes1.Add "iOcrH = " & Split(localise, ";")(3)                               '
    EnsembleLignes1.Add "sTextToCompare = " & writeValue                                   '
    EnsembleLignes1.Add ""                                                                 '
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"                          '
    EnsembleLignes1.Add "%NAME = ""Optical Character Recognition"""                        '
    '
    '
    Compteur = Compteur + 1                                             ''''''''''''''''''''

End Sub

Private Sub Traitement_RemoteControlStart(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"                                  '''fonction déplacée dans le setup
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Start"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:ADj0HVcsMUi0X0i6ntHHND"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "sRemoteEqtAddress = " & strNickName
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""Remote Control Start"""

    Compteur = Compteur + 1

End Sub

Private Sub Traitement_RemoteControlStop(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"                                  '''fonction déplacée dans le cleanup
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Stop"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:PJPMaXLACkmkneVke5M50B"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""Remote Control Stop"""

    Compteur = Compteur + 1

End Sub

Private Sub Traitement_SendText(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, localise As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"                          ''''''''''''''''''''pas utilisée
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """TH_Text"""                     '
    '
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"                             '
    EnsembleLignes1.Add "Id = ""ID#:6R05Sk2J0E6Md7BvcVHXyA"""                               '
    EnsembleLignes1.Add ""                                                                  '
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"                                '
    EnsembleLignes1.Add "text = " & writeValue                                              '
    EnsembleLignes1.Add ""                                                                  '
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"                           '
    EnsembleLignes1.Add "%NAME = ""Send Text"""                                             '
    '
    Compteur = Compteur + 1                                                   '''''''''''''''

End Sub
Private Sub TraitementLabel(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur)

    If sLigne <> "" Then
    ' traitement du nom avec instance
    Do While sLigne Like "*<*>*"
    temp = Mid(sLigne, InStr(1, sLigne, "<"), InStr(1, sLigne, ">") + 1 - InStr(1, sLigne, "<"))
    X = 0
    Do While instance(X, 0) <> "" Or instance(X + 1, 0) <> ""
    If temp = instance(X, 0) Then
    If Not temp Like "*_*" Then instance(X, 2) = 1
    If temp Like "*_*" Then instance(X, 2) = 2
    sLigne = Replace(sLigne, instance(X, 0), instance(X, 1))
    Exit Do
    End If
    X = X + 1
    Loop
    If sLigne Like "*<*>*" And instance(X, 0) = "" Then
    MsgBox ("Instance non trouvée pour la variable " & sLigne & "   ligne " & i)
    End
    End If
    Loop

    'Pour les Labels
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """Label"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
    End If
End Sub

Private Sub TraitementCB_Force(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, vForcedValue As Variant, Section As String)

    'Pour Forcer les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_Force"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    'EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{60CA140F-FBED-44D2-A0DF-DBCB2D65E7C0}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809363"
    EnsembleLignes1.Add "MemberName = ""CB_Force"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    'EnsembleLignes1.Add "TypeLibVersion = ""453.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    'EnsembleLignes1.Add "VTableOffset = 204"
    EnsembleLignes1.Add "%HI: Parameters = [4]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""vForcedValue"""
    EnsembleLignes1.Add "ArgVal =""" & vForcedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vForcedValue & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    '    If Quick_acces = True Then EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    'EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    'EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    'EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub
Private Sub TraitementCB_ForceArrayElt(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, Index As Integer, vForcedValue As Variant)

    'Pour Forcer les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_ForceArrayElement"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb"""
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{39FB1F8C-2512-42A1-9E37-D22340C540B9}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809398"
    EnsembleLignes1.Add "MemberName = ""CB_ForceArrayElement"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""358.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 340"
    EnsembleLignes1.Add "%HI: Parameters = [5]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""nIndex"""
    EnsembleLignes1.Add "ArgVal =""" & Index & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & Index & "\"""""
    EnsembleLignes1.Add "Type = 2"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""vForcedValue"""
    EnsembleLignes1.Add "ArgVal = ""\""" & vForcedValue & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vForcedValue & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1
End Sub
Private Sub TraitementCB_UnForce(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, Section As String)

    'Pour de-Forcer les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_UnForce"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809365"

    EnsembleLignes1.Add "MemberName = ""CB_UnForce"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 140"
    EnsembleLignes1.Add "%HI: Parameters = [3]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1
End Sub
Private Sub TraitementCB_Write(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, nValue As Variant, Section As String)

    'Pour écrire les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_Write"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{95CEEF41-10E4-4731-BE74-6928FB00A801}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809347"

    EnsembleLignes1.Add "MemberName = ""CB_Write"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""26d.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 88"
    EnsembleLignes1.Add "%HI: Parameters = [4]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""nValue"""
    EnsembleLignes1.Add "ArgVal =""" & nValue & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & nValue & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_WriteArrayElt(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, nValue As Variant, nIndex As Integer)

    'Pour écrire les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_WriteArrayElement"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb"""
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{95CEEF41-10E4-4731-BE74-6928FB00A801}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809397"
    EnsembleLignes1.Add "MemberName = ""CB_WriteArrayElement"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    'EnsembleLignes1.Add "TypeLibVersion = ""26d.0"""
    EnsembleLignes1.Add "TypeLibVersion = ""34d.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 336"
    EnsembleLignes1.Add "%HI: Parameters = [5]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""nIndex"""
    EnsembleLignes1.Add "ArgVal =""" & nIndex & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & nIndex & "\"""""
    EnsembleLignes1.Add "Type = 2"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""nValue"""
    EnsembleLignes1.Add "ArgVal =""" & nValue & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & nValue & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_Read(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, vReadValue As Variant, Section As String)

    'Pour lire les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_Read"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{65D7DD1D-CEED-49EA-A31F-2A4F70D9A107}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809346"

    EnsembleLignes1.Add "MemberName = ""CB_Read"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""454.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 156"
    EnsembleLignes1.Add "%HI: Parameters = [4]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""vReadValue"""
    EnsembleLignes1.Add "ArgVal =""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal =""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"

    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_Test(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, AwaitedValue As Variant, Section As String, Optional lTimeout_ms As Variant)

    'Pour tester les variables sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_Test"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{65D7DD1D-CEED-49EA-A31F-2A4F70D9A107}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809350"

    EnsembleLignes1.Add "MemberName = ""CB_Test"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""454.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 164"
    EnsembleLignes1.Add "%HI: Parameters = [5]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""awaitedValue"""
    EnsembleLignes1.Add "ArgVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_TestDT(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, vReadValue As Variant, AwaitedValue As Variant, Section As String)

    Call TraitementCB_Read(sLigne, EnsembleLignes0, EnsembleLignes1, Compteur, strVariableName, strInstanceName, vReadValue, Section)
    ' FIN DU READ

    'DT formatting

    'Add start/end quotes if necessary
    If Left(AwaitedValue, 1) = """" Then
    AwaitedValue = Right(AwaitedValue, Len(AwaitedValue) - 1)
    End If
    If Right(AwaitedValue, 1) = """" Then
    AwaitedValue = Left(AwaitedValue, Len(AwaitedValue) - 1)
    End If

    DTError = False

    'Check if prefix is "DT#"
    tmp = Split(AwaitedValue, "#")
    If tmp(0) <> "DT" Then
    DTError = True
    Else
    'Check Date formatting
    DateTab = Split(tmp(1), "-")

    If UBound(DateTab) < 2 Then
    DTError = True
    Else
    tYear = CInt(DateTab(0))
    tMonth = CInt(DateTab(1))
    tDay = CInt(DateTab(2))
    End If

    'Check Time formatting
    TimeTab = Split(DateTab(3), ":")
    If UBound(TimeTab) < 2 Then
    DTError = True
    Else
    tHour = CInt(TimeTab(0))
    tMin = CInt(TimeTab(1))
    tSec = CInt(TimeTab(2))
    End If

    AwaitedValue = "DT#" & tYear & "-" & tMonth & "-" & tDay & "-" & tHour & ":" & tMin & ":" & tSec

    End If

    If DTError Then
    MsgBox "TestDT value is not a DateTime value", vbCritical
    End If

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = ""Statement"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "StatusExpr =""" & vReadValue & " == \""" & AwaitedValue & "\""? \""Passed\"": \""Failed\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""Statement"""

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_TestBool(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, AwaitedValue As Boolean, Section As String, Optional lTimeout_ms As Variant)

    'Pour tester les variables booléenne sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_TestBool"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809350"
    EnsembleLignes1.Add "MemberName = ""CB_TestBool"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 96"
    EnsembleLignes1.Add "%HI: Parameters = [5]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""awaitedValue"""
    EnsembleLignes1.Add "ArgVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & AwaitedValue & "\"""""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1



End Sub
Private Sub TraitementCB_TestAna(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, AwaitedValue As Variant, Optional maxValue, Optional nTolerance, Optional lTimeout_ms As Variant)

    'Pour tester les variables analogique sans mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_TestAna"""


    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb"""
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809352"
    EnsembleLignes1.Add "MemberName = ""CB_TestAna"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 100"
    EnsembleLignes1.Add "%HI: Parameters = [7]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[7] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""awaitedValue"""
    EnsembleLignes1.Add "ArgVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & AwaitedValue & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""maxValue"""
    EnsembleLignes1.Add "ArgVal =""" & IIf(IsMissing(maxValue), 0, maxValue) & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(IsMissing(maxValue), 0, maxValue) & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""nTolerance"""
    EnsembleLignes1.Add "ArgVal =""" & IIf(IsMissing(nTolerance), 0, nTolerance) & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(IsMissing(nTolerance), 0, nTolerance) & "\"""""
    EnsembleLignes1.Add "Type = 4"
    EnsembleLignes1.Add "DisplayType = ""Number (32-bit Floating Point)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[6]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[7]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_ForceNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, vForcedValue As Variant, Section As String)
    'Pour Forcer les variables reseaux (avec Mnémonique)
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_ForceNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{1BC9ACCA-0930-40DD-99EB-531946BC8902}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809365"

    EnsembleLignes1.Add "MemberName = ""CB_ForceNN"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""3d6.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 208"
    EnsembleLignes1.Add "%HI: Parameters = [3]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"

    EnsembleLignes1.Add "Name = ""strNickName"""

    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"

    EnsembleLignes1.Add "Name = ""vForcedValue"""

    EnsembleLignes1.Add "ArgVal =""" & vForcedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & vForcedValue & """"
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"

    EnsembleLignes1.Add "Direction = 1"

    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1
End Sub
Private Sub TraitementCB_UnForceNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, Section As String)

    'Pour de-forcer les variables avec mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_UnForceNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809366"

    EnsembleLignes1.Add "MemberName = ""CB_UnForceNN"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 144"
    EnsembleLignes1.Add "%HI: Parameters = [2]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"

    EnsembleLignes1.Add "Name = ""strNickName"""

    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_SetNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName, nValue, Section As String)
    'Pour Ecrire les variables reseaux (avec Mnémonique)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_SetNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{60CA140F-FBED-44D2-A0DF-DBCB2D65E7C0}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809369"

    EnsembleLignes1.Add "MemberName = ""CB_WriteNN"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""453.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 224"
    EnsembleLignes1.Add "%HI: Parameters = [3]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"

    EnsembleLignes1.Add "Name = ""strNickName"""

    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""nValue"""
    EnsembleLignes1.Add "ArgVal =""" & nValue & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & nValue & """"
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"

    EnsembleLignes1.Add "Direction = 1"

    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1
End Sub
Private Sub TraitementCB_ReadNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, vReadValue, Section As String)

    'Pour lire les variables avec mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_ReadNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809367"

    EnsembleLignes1.Add "MemberName = ""CB_ReadNN"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 148"
    EnsembleLignes1.Add "%HI: Parameters = [3]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"

    EnsembleLignes1.Add "Name = ""strNickName"""

    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""vReadValue"""
    EnsembleLignes1.Add "ArgVal =""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal =""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef  = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1

End Sub
Private Sub TraitementCB_TestNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, AwaitedValue As Variant, Section As String, Optional lTimeout_ms As Variant)

    'Pour tester les variables avec mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_TestNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809371"

    EnsembleLignes1.Add "MemberName = ""CB_TestNN"""

    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 164"
    EnsembleLignes1.Add "%HI: Parameters = [4]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"

    EnsembleLignes1.Add "Name = ""strNickName"""

    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""awaitedValue"""

    If UCase(AwaitedValue) = "FALSE" Or UCase(AwaitedValue) = "TRUE" Or IsNumeric(Replace(AwaitedValue, ".", ",")) Then
    EnsembleLignes1.Add "ArgVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & AwaitedValue & """"
    ElseIf InStr(1, AwaitedValue, ".") <> 0 Then
    EnsembleLignes1.Add "ArgVal = " & AwaitedValue
    EnsembleLignes1.Add "ArgDisplayVal = " & AwaitedValue
    Else
    EnsembleLignes1.Add "ArgVal = ""\""" & AwaitedValue & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & AwaitedValue & "\"""""
    End If

    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"

    Compteur = Compteur + 1


End Sub
Private Sub TraitementCB_TestBoolNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, AwaitedValue As Variant, Section As String, Optional lTimeout_ms As Variant)

    'Pour tester les variables boolenne avec mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_TestBoolNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809369"
    EnsembleLignes1.Add "MemberName = ""CB_TestBoolNN"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 156"
    EnsembleLignes1.Add "%HI: Parameters = [4]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strNickName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""awaitedValue"""
    EnsembleLignes1.Add "ArgVal =  ""\""" & AwaitedValue & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & AwaitedValue & "\"""""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    Compteur = Compteur + 1


End Sub
Private Sub TraitementCB_TestAnaNN(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As String, AwaitedValue As Variant, Section As String, Optional maxValue, Optional nTolerance, Optional lTimeout_ms As Variant)

    'Pour tester les variables analogique avec mnémonique
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_TestAnaNN"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{0B54A075-8ABC-42CD-9BB7-7C9CC23D01F1}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809370"
    EnsembleLignes1.Add "MemberName = ""CB_TestAnaNN"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""22b.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 160"
    EnsembleLignes1.Add "%HI: Parameters = [6]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strNickName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""awaitedValue"""
    EnsembleLignes1.Add "ArgVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""maxValue"""
    EnsembleLignes1.Add "ArgVal =""" & IIf(maxValue = "", 0, maxValue) & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(maxValue = "", 0, maxValue) & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""nTolerance"""
    EnsembleLignes1.Add "ArgVal =""" & IIf(nTolerance = "", 0, nTolerance) & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(nTolerance = "", 0, nTolerance) & "\"""""
    EnsembleLignes1.Add "Type = 4"
    EnsembleLignes1.Add "DisplayType = ""Number (32-bit Floating Point)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[6]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strNickName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    Compteur = Compteur + 1


End Sub
Private Sub TraitementCB_TestArrayElt(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strVariableName As String, strInstanceName As String, nIndex As Variant, AwaitedValue As Variant, Optional nVariation, Optional nVariationPercent, Optional lTimeout_ms As Variant)

    'Pour tester un élement de tableau
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_TestArrayElement"""
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb"""
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{A1C02E91-B920-4049-9832-51D70EE7E43C}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809400"
    EnsembleLignes1.Add "MemberName = ""CB_TestArrayElement"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""348.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 344"
    EnsembleLignes1.Add "%HI: Parameters = [8]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[7] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[8] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""strInstanceName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""strVariableName"""
    EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[2]]"
    EnsembleLignes1.Add "Name = ""nIndex"""
    EnsembleLignes1.Add "ArgVal =""" & nIndex & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & nIndex & "\"""""
    EnsembleLignes1.Add "Type = 2"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[3]]"
    EnsembleLignes1.Add "Name = ""AwaitedValue"""
    EnsembleLignes1.Add "ArgVal =""" & AwaitedValue & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & AwaitedValue & "\"""""
    EnsembleLignes1.Add "Type = 12"
    EnsembleLignes1.Add "DisplayType = ""Variant"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[4]]"
    EnsembleLignes1.Add "Name = ""nVariation"""
    EnsembleLignes1.Add "ArgVal =""" & IIf(IsMissing(nVariation), 0, nVariation) & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(IsMissing(nVariation), 0, nVariation) & "\"""""
    EnsembleLignes1.Add "Type = 5"
    EnsembleLignes1.Add "DisplayType = ""Number (64-bit Floating Point)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[5]]"
    EnsembleLignes1.Add "Name = ""nVariationPercent"""
    EnsembleLignes1.Add "ArgVal =""" & IIf(IsMissing(nVariationPercent), 100, nVariationPercent) & """"
    EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(IsMissing(nVariationPercent), 100, nVariationPercent) & "\"""""
    EnsembleLignes1.Add "Type = 5"
    EnsembleLignes1.Add "DisplayType = ""Number (64-bit Floating Point)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[6]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[7]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[8]]"
    EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    EnsembleLignes1.Add "ArgVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "ArgDisplayVal =""" & lTimeout_ms & """"
    EnsembleLignes1.Add "Type = 3"
    EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 1"
    EnsembleLignes1.Add "IsUserOptional = True"
    EnsembleLignes1.Add "IsServerOptional = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1

End Sub
Private Sub TraitementSequenceCall(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, NameSequence, SequenceFile, proj_filglo, proj_conf, NSection)

    'Pour les Appels de sequence
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """SequenceCall"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add "LoadOpt = ""DynamicLoad"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, SeqCallStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData]"
    EnsembleLignes1.Add "SFPath =""" & SequenceFile & """"
    EnsembleLignes1.Add "SeqName =""" & NameSequence & """"
    EnsembleLignes1.Add "SpecifyHostByExpr = True"
    EnsembleLignes1.Add "RemoteExecution = True"
    EnsembleLignes1.Add "RemoteHostExpr = ""FileGlobals.PC_Env" & Right(NSection, 1) & "_ip"""
    'If NSection Like "_*" Then
    '    If NSection Like "*2" Then
    '        EnsembleLignes1.Add "SpecifyHostByExpr = True"
    '        EnsembleLignes1.Add "RemoteExecution = True"
    '        EnsembleLignes1.Add "RemoteHostExpr = ""FileGlobals.PC_Env" & Right(NSection, 1) & "_ip"""
    '    End If
    'Else
    '    EnsembleLignes1.Add "SpecifyHostByExpr = True"
    '    EnsembleLignes1.Add "RemoteExecution = True"
    '    EnsembleLignes1.Add "RemoteHostExpr = ""FileGlobals.PC_Emb" & NSection & "_ip"""
    'End If
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub
Private Sub TraitementSequenceExe(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, NameSequence, SequenceFile)

    'Pour les Appels de sequence
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """SequenceCall"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, SeqCallStepAdditions"""
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData]"
    EnsembleLignes1.Add "SFPath =""" & SequenceFile & """"
    EnsembleLignes1.Add "SeqName =""" & "Single Pass" & """"
    EnsembleLignes1.Add "%FLG: Prototype = 262144"
    EnsembleLignes1.Add "AutoWaitAsync = False"
    EnsembleLignes1.Add "ThreadOpt = 2"
    EnsembleLignes1.Add "ExecTypeMask = 16"
    EnsembleLignes1.Add "ExecSync = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData]"
    EnsembleLignes1.Add "ActualArgs = Arguments"
    EnsembleLignes1.Add "Prototype = Obj"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs]"
    EnsembleLignes1.Add "sequence = """ & "TYPE, SequenceArgument" & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Prototype]"
    EnsembleLignes1.Add "sequence = Str"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Prototype]"
    EnsembleLignes1.Add "sequence = """ & "MainSequence" & """"
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub
Private Sub Gen_SequenceArguments(ByRef EnsembleLignes1, ByRef Compteur, proj_filglo, proj_conf, Section As String)

    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs]"

    EnsembleLignes1.Add "CB_Simu" & " = ""TYPE, SequenceArgument"""
    EnsembleLignes1.Add "Embedded" & " = ""TYPE, SequenceArgument"""
    EnsembleLignes1.Add "Embedded1" & " = ""TYPE, SequenceArgument"""
    EnsembleLignes1.Add "Embedded2" & " = ""TYPE, SequenceArgument"""
    '    EnsembleLignes1.Add "QuickAccess" & " = ""TYPE, SequenceArgument"""
    '   EnsembleLignes1.Add "QuickForceCheck" & " = ""TYPE, SequenceArgument"""
    EnsembleLignes1.Add "InitDone" & " = ""TYPE, SequenceArgument"""

    For i = 0 To UBound(proj_conf, 2)
    EnsembleLignes1.Add proj_conf(0, i) & " = ""TYPE, SequenceArgument"""
    Next
    'FileGlobals variables
    For i = 0 To UBound(proj_filglo, 2)
    EnsembleLignes1.Add proj_filglo(0, i) & " = ""TYPE, SequenceArgument"""
    Next

    EnsembleLignes1.Add ""
    'declaration du nom
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "CB_Simu" & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals.cb" & Section & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "Embedded" & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals." & "Embedded" & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "Embedded1" & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals." & "Embedded1" & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "Embedded2" & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals." & "Embedded2" & """"
    EnsembleLignes1.Add ""
    '      EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "QuickAccess" & "]"
    '      EnsembleLignes1.Add "UseDef = False"
    '      EnsembleLignes1.Add "Expr = ""FileGlobals." & "QuickAccess" & """"
    '      EnsembleLignes1.Add ""
    '      EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "QuickForceCheck" & "]"
    '      EnsembleLignes1.Add "UseDef = False"
    '      EnsembleLignes1.Add "Expr = ""FileGlobals." & "QuickForceCheck" & """"
    '      EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & "InitDone" & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals." & "InitDone" & """"
    EnsembleLignes1.Add ""

    For i = 0 To UBound(proj_conf, 2)
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & proj_conf(0, i) & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals." & proj_conf(0, i) & """"
    EnsembleLignes1.Add ""
    Next

    For i = 0 To UBound(proj_filglo, 2)
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.ActualArgs." & proj_filglo(0, i) & "]"
    EnsembleLignes1.Add "UseDef = False"
    EnsembleLignes1.Add "Expr = ""FileGlobals." & proj_filglo(0, i) & """"
    EnsembleLignes1.Add ""
    Next

    EnsembleLignes1.Add ""
    'definition du type
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Prototype]"
    EnsembleLignes1.Add "CB_Simu" & " = " & ""
    EnsembleLignes1.Add "Embedded" & " = " & ""
    EnsembleLignes1.Add "Embedded1" & " = " & ""
    EnsembleLignes1.Add "Embedded2" & " = " & ""
    '       EnsembleLignes1.Add "QuickAccess" & " = " & ""
    '        EnsembleLignes1.Add "QuickForceCheck" & " = " & ""
    EnsembleLignes1.Add "InitDone" & " = " & ""

    For i = 0 To UBound(proj_conf, 2)
    EnsembleLignes1.Add proj_conf(0, i) & " = " & proj_conf(2, i)
    Next
    For i = 0 To UBound(proj_filglo, 2)
    EnsembleLignes1.Add proj_filglo(0, i) & " = " & proj_filglo(2, i)
    Next

    EnsembleLignes1.Add ""
    'attribution d'une valeur
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Prototype]"
    EnsembleLignes1.Add "%FLG: " & "CB_Simu" & " = 4"
    EnsembleLignes1.Add "%FLG: " & "Embedded" & " = 4"
    EnsembleLignes1.Add "%FLG: " & "Embedded1" & " = 4"
    EnsembleLignes1.Add "%FLG: " & "Embedded2" & " = 4"
    EnsembleLignes1.Add "%FLG: " & "InitDone" & " = 4"

    For i = 0 To UBound(proj_conf, 2)
    EnsembleLignes1.Add "%FLG: " & proj_conf(0, i) & " = 4"
    Next
    For i = 0 To UBound(proj_filglo, 2)
    EnsembleLignes1.Add "%FLG: " & proj_filglo(0, i) & " = 4"
    Next

    EnsembleLignes1.Add ""

End Sub

Private Sub TraitementNI_Wait(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, TimeValue)

    'Pour les Wait
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """NI_Wait"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "SeqCallStepGroupIdx = -1"
    EnsembleLignes1.Add "TimeExpr =""" & TimeValue & """"
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub TraitementCB_StopTask(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur)

    'Pour les Initialisation de control build
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_StopTask"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call]"
    EnsembleLignes1.Add "ObjectVariable = ""FileGlobals.cb"""
    EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
    EnsembleLignes1.Add "CreateObject = 0"
    EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    EnsembleLignes1.Add "Interface = ""{65D7DD1D-CEED-49EA-A31F-2A4F70D9A107}"""
    EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    EnsembleLignes1.Add "MemberType = 1"
    EnsembleLignes1.Add "Member = 1610809361"
    EnsembleLignes1.Add "MemberName = ""CB_StopTask"""
    EnsembleLignes1.Add "HasMemberInfo = True"
    EnsembleLignes1.Add "Locale = 1036"
    EnsembleLignes1.Add "TypeLibVersion = ""454.0"""
    EnsembleLignes1.Add "InterfaceType = 1"
    EnsembleLignes1.Add "VTableOffset = 196"
    EnsembleLignes1.Add "%HI: Parameters = [1]"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters]"
    EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[0]]"
    EnsembleLignes1.Add "Name = ""passFail"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    EnsembleLignes1.Add "Type = 11"
    EnsembleLignes1.Add "DisplayType = ""Boolean"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS.SData.Call.Parameters[1]]"
    EnsembleLignes1.Add "Name = ""errorMsg"""
    EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    EnsembleLignes1.Add "Type = 8"
    EnsembleLignes1.Add "DisplayType = ""String"""
    EnsembleLignes1.Add "TypeValid = True"
    EnsembleLignes1.Add "Direction = 3"
    EnsembleLignes1.Add "IsByRef = True"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub TraitementCB_StartTask(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur)

    'Pour les Initialisation de control build
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_StartTask"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub


Private Sub TraitementCB_Init(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur)

    'Pour l'Initialisation de control build
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_Init"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub TraitementCB_MsgPopUp(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, vMessage As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """MessagePopup"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add "StatusExpr = ""Step.Result.ButtonHit == 1 ? \""Passed\"": Step.Result.ButtonHit == 2 ? \""Failed\"": \""Done\"""""
    EnsembleLignes1.Add ""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "TitleExpr = ""\""Pop Up Message\"""""
    EnsembleLignes1.Add "MessageExpr = ""\""" & vMessage & "\"""""
    EnsembleLignes1.Add "Button1Label = ""\""OK\"""""
    EnsembleLignes1.Add "Button2Label = ""\""NOK\"""""
    EnsembleLignes1.Add "Button3Label = ""\""SKIP\"""""
    EnsembleLignes1.Add "CancelButton = 0"

    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = ""Message Popup"""
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub TraitementCB_Close(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur)

    'Pour la fermeture de l'arbre control build
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """CB_Close"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub TraitementTestDDU(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, readValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """MessagePopup"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add "ConnectionLifetime = 4"
    'EnsembleLignes1.Add "PreCond = ""FileGlobals.Type_simu == 2"""
    EnsembleLignes1.Add "StatusExpr = ""Step.Result.ButtonHit ==1 ? \""Passed\"": Step.Result.ButtonHit ==2 ? \""Failed\"":\""Done\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "TitleExpr = ""\""Read DDU\"""""
    If strNickName <> "" Then
    EnsembleLignes1.Add "MessageExpr = ""\""Check the variable of nickname " & strNickName & " at " & readValue & " on the DDU\"""""
    Else
    EnsembleLignes1.Add "MessageExpr = """"" & sLigne & """"""
    End If
    EnsembleLignes1.Add "Button1Label = ""\""OK\"""""
    EnsembleLignes1.Add "Button2Label = ""\""NOK\"""""
    EnsembleLignes1.Add "Button3Label = ""\""SKIP\"""""
    EnsembleLignes1.Add "CancelButton = 0"
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & "; " & strNickName & "; " & readValue & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub TraitementWriteDDU(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, writeValue As Variant)

    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """MessagePopup"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add "ConnectionLifetime = 4"
    EnsembleLignes1.Add "PreCond = ""FileGlobals.Type_simu == 2"""
    EnsembleLignes1.Add "StatusExpr = ""Step.Result.ButtonHit <2 ? \""Passed\"": \""Failed\"""""
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "TitleExpr = ""\""Write DDU\"""""
    EnsembleLignes1.Add "MessageExpr = ""\""Set the variable of nickname " & strNickName & " at " & writeValue & " on the DDU\"""""
    EnsembleLignes1.Add "Button1Label = ""\""OK\"""""
    EnsembleLignes1.Add "Button2Label = ""\""NOK\"""""
    EnsembleLignes1.Add "CancelButton = 0"
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME =""" & sLigne & "; " & strNickName & "; " & writeValue & """"
    EnsembleLignes1.Add ""

    Compteur = Compteur + 1
End Sub

Private Sub Traitement_Statement(sLigne As String, EnsembleLignes0, EnsembleLignes1, ByRef Compteur, strNickName As Variant, vForcedValue As Variant)
    'test statement
    EnsembleLignes0.Add "%[" & Compteur & "] = Step"
    EnsembleLignes0.Add "%TYPE: %[" & Compteur & "] = " & """Statement"""

    EnsembleLignes1.Add "[SF.Seq[0].Main[" & Compteur & "].TS]"
    EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(Compteur)))), "0") & Trim(Str$(Compteur)) & """"
    EnsembleLignes1.Add "PostExpr = """ & strNickName & " = " & vForcedValue & """"
    EnsembleLignes1.Add ""
    EnsembleLignes1.Add "[DEF, SF.Seq[0].Main[" & Compteur & "]]"
    EnsembleLignes1.Add "%NAME = """ & sLigne & """"

    Compteur = Compteur + 1
End Sub

