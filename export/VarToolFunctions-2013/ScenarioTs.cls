VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScenarioTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim i, j, l, y, x, z, m As Long
Dim sFileseq As String
Dim compteur As Long
Dim Seq() As SubSequence
Dim nb_seq, seq_title_sup, seq_mode_sup, seq_chapter_sup As Long
Dim MODE As String
Dim fic As Integer

Const MAX_COUNTER As Long = 65536

Const ACTION_INIT_TASK As String = "Init_Task"
Const ACTION_POPUP As String = "Popup"
Const ACTION_FORCE As String = "Force"
Const ACTION_FORCE_NN As String = "ForceNN"
Const ACTION_UNFORCE As String = "UnForce"
Const ACTION_UNFORCE_NN As String = "UnForceNN"
Const ACTION_WRITE As String = "Write"
Const ACTION_WRITE_NN As String = "WriteNN"
Const ACTION_READ As String = "Read"
Const ACTION_READ_NN As String = "ReadNN"
Const ACTION_TEST As String = "Test"
Const ACTION_TEST_NN As String = "TestNN"
Const ACTION_TEST_ANA As String = "TestAna"
Const ACTION_TEST_ANA_NN As String = "TestAnaNN"
Const ACTION_CALL As String = "Call"
Const ACTION_WAIT As String = "Wait"
Const ACTION_UNFORCE_ARRAY As String = "UnForceArray"
Const ACTION_UNFORCE_ARRAY_NN As String = "UnForceArrayNN"
Const ACTION_FORCE_ARRAY_ALL As String = "ForceArrayAll"
Const ACTION_FORCE_ARRAY_ALL_NN As String = "ForceArrayAllNN"
Const ACTION_FORCE_ARRAY_ELT As String = "ForceArrayElt"
Const ACTION_FORCE_ARRAY_ELT_NN As String = "ForceArrayEltNN"
Const ACTION_QA_RESET_ALL As String = "QA_ResetAll"
Const ACTION_QA_UNFORCE_ALL As String = "QA_ResetAll"
Const ACTION_QA_FORCE_VAR As String = "QA_ForceVar"
Const ACTION_QA_UNFORCE_VAR As String = "QA_UnForceVar"
Const ACTION_TEST_MMI_150 As String = "Test_MMI_150"
Const ACTION_TEST_CMX_EVOL As String = "Test_CMxEvol"
Const ACTION_HMI_START_TEST As String = "HMI_Start_Test"
Const ACTION_HMI_STOP_TEST As String = "HMI_Stop_Test"
Const ACTION_HMI_SEND_KEY As String = "HMI_Send_Key"
Const ACTION_STATEMENT As String = "Statement"

Type SubSequence
    cpt As Long
    Host As String
    EnsembleLignes0 As New Collection
    EnsembleLignes1 As New Collection
    Comment As String
    name As String
End Type

Public Sub writeHeader()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Public Sub writeSheet(Sheet As Worksheet)
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Public Sub writeFooter()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Private Sub Class_Initialize()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Private Sub Class_Terminate()
    MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
End Sub

Sub ProcessingExcel(file As String)
    nb_seq = 1
    compteur = 0

Dim Type_Mode As String
Type_Mode = ThisWorkbook.Worksheets("Parameters").Range("Type_Mode").Value

If Type_Mode Like "Mode Dégradé" Then
    MODE = Type_Mode
Else
    MODE = "Mode Normal"
End If

restart:

ThisWorkbook.Worksheets(MODE).Activate
i = 4
j = 2



For l = i To MAX_COUNTER
go_T:
    

    If Worksheets(MODE).Cells(l, 2) <> "" Then
    
        If Worksheets(MODE).Cells(l, 2) Like "END*" Then
            MODE = ""
            GoTo fin
        Else
            If Worksheets(MODE).Cells(l, 2) Like ">>>*" Then
                MODE = "Mode Dégradé"
                GoTo restart
            Else
    
    
            Call Processing_SubSequence(Worksheets(MODE).Cells(l, 2), nb_seq, Seq, compteur, 1)
            seq_mode_sup = nb_seq

Retour_Chapter:
            
            For x = l + 2 To MAX_COUNTER
                
                If Worksheets(MODE).Cells(x, 3) <> "" Then
                    Call Processing_SubSequence(Worksheets(MODE).Cells(x, 3), nb_seq, Seq, compteur, seq_mode_sup)
                    seq_chapter_sup = nb_seq
                
Retour_Title:
                                       
                    For y = x + 2 To MAX_COUNTER
                        
                        If Worksheets(MODE).Cells(y, 4) <> "" And Worksheets(MODE).Cells(y, 4).Interior.ColorIndex <> 36 Then
                            Call Processing_SubSequence(Worksheets(MODE).Cells(y, 4), nb_seq, Seq, compteur, seq_chapter_sup)
                            seq_title_sup = nb_seq
                            
Retour_step:
                            
                            For z = y + 2 To MAX_COUNTER
                            
                        
                                If Worksheets(MODE).Cells(z, 5) Like "*STEP*" Then
                                     Call Processing_SubSequence(Worksheets(MODE).Cells(z, 5) & " " & Worksheets(MODE).Cells(z, 6), nb_seq, Seq, compteur, seq_title_sup, Worksheets(MODE).Cells(z, 4), Worksheets(MODE).Cells(z, 6))
                                    For m = z + 1 To MAX_COUNTER
                                    
                                        If Worksheets(MODE).Cells(m, 5).Interior.ColorIndex <> xlNone Then
                                            y = m - 2
                                            GoTo Retour_step
                                        Else
                                            Call ProcessingLabel(Worksheets(MODE).Cells(m, 6), Seq, nb_seq, compteur)
                                            Call ProcessingAction(Seq, nb_seq, compteur, Worksheets(MODE).Cells(m, 7), m, 7)
                                            Call ProcessingLabel(Worksheets(MODE).Cells(m, 11), Seq, nb_seq, compteur)
                                            Call ProcessingAction(Seq, nb_seq, compteur, Worksheets(MODE).Cells(m, 12), m, 12)
                                        End If
                                    Next m
                                Else
                                    If Worksheets(MODE).Cells(z, 5) Like "END" Then
                                        x = z
                                        GoTo Retour_Title
                                    End If
                                End If
                            Next z
                        Else
                            If Worksheets(MODE).Cells(y, 3) <> "" Then
                                l = y - 2
                                GoTo Retour_Chapter
                            End If
                        End If
                        If Worksheets(MODE).Cells(y, 2).Interior.ColorIndex = 11 Then
                            l = y
                            GoTo go_T
                        End If
                    Next y
                Else
                    If Worksheets(MODE).Cells(x, 2) <> "" Then
                        l = x - 1
                        GoTo go_T
                    End If
                End If
            Next x
            End If
        End If
    End If
Next l


fin:


For i = 1 To nb_seq
    If Seq(i).cpt <= 0 Then
        Seq(i).cpt = 1
End If
Next i
    
fic = FreeFile
Open file For Output As #fic
Call CreationFicSeq(fic, sFileseq, compteur, nb_seq, Seq)
Close #fic



End Sub

Sub Processing_SubSequence(sLigne As String, ByRef nb_seq, ByRef Seq() As SubSequence, compteur, seq_sup, Optional HostName As String, Optional ByVal Comment As String = "")
    
    If sLigne <> "" Then
        nb_seq = nb_seq + 1
        ReDim Preserve Seq(nb_seq)
        
        If nb_seq > 1 Then
            Seq(nb_seq).Host = HostName
            Seq(nb_seq).name = sLigne
            Seq(nb_seq).Comment = sLigne
        End If

        Seq(seq_sup).EnsembleLignes0.Add "%[" & Seq(seq_sup).cpt & "] = Step"
        Seq(seq_sup).EnsembleLignes0.Add "%TYPE: %[" & Seq(seq_sup).cpt & "] = " & """SequenceCall"""
    
        Seq(seq_sup).EnsembleLignes1.Add "[SF.Seq[" & seq_sup - 1 & "].Main[" & Seq(seq_sup).cpt & "].TS]"
        Seq(seq_sup).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
        Seq(seq_sup).EnsembleLignes1.Add ""
        Seq(seq_sup).EnsembleLignes1.Add "[DEF, SF.Seq[" & seq_sup - 1 & "].Main[" & Seq(seq_sup).cpt & "].TS]"
        Seq(seq_sup).EnsembleLignes1.Add "SData = ""TYPE, SeqCallStepAdditions"""
        Seq(seq_sup).EnsembleLignes1.Add ""
        Seq(seq_sup).EnsembleLignes1.Add "[SF.Seq[" & seq_sup - 1 & "].Main[" & Seq(seq_sup).cpt & "].TS.SData]"
        Seq(seq_sup).EnsembleLignes1.Add ("SeqName = """ & sLigne & """") '""" & Name & """"
        Seq(seq_sup).EnsembleLignes1.Add "UseCurFile = True"
        'Seq(seq_sup).EnsembleLignes1.Add "Trace = ""Don\t Change"""
        If HostName Like "" Or HostName Like "Local" Then
        Else
            Seq(seq_sup).EnsembleLignes1.Add "SpecifyHostByExpr = True"
            Seq(seq_sup).EnsembleLignes1.Add "RemoteExecution = True"
            Seq(seq_sup).EnsembleLignes1.Add "RemoteHostExpr = """ & HostName & """"
        End If
        Seq(seq_sup).EnsembleLignes1.Add ""
        Seq(seq_sup).EnsembleLignes1.Add "[DEF, SF.Seq[" & seq_sup - 1 & "].Main[" & Seq(seq_sup).cpt & "]]"
        If sLigne Like ("*STEP*") Then
            Comment = ": " & Comment
            Seq(seq_sup).EnsembleLignes1.Add ("%NAME = """ & sLigne & Comment & """")
        Else
            Seq(seq_sup).EnsembleLignes1.Add ("%NAME = """ & sLigne & """")
        End If
        Seq(seq_sup).EnsembleLignes1.Add ""

        Seq(seq_sup).cpt = Seq(seq_sup).cpt + 1
        compteur = compteur + 1
    
    End If
    
End Sub

Sub ProcessingAction(ByRef Seq() As SubSequence, nb_seq, ByRef compteur, MotAction As String, ByVal i As Long, ByVal j As Long)
Dim strLib As String
Dim Name_Var As String
Dim Precon As String
Dim timeout As String

If j = 7 Then
    Precon = Cells(i, j + 11)
Else
    Precon = Cells(i, j + 6)
End If

Name_Var = Cells(i, j + 1)
If Name_Var Like "Entrée_*" Then
    Name_Var = Replace(Name_Var, "Entrée_", "")
Else
    If Name_Var Like "Sortie_*" Then
        Name_Var = Replace(Name_Var, "Sortie_", "")
    End If
End If

If Precon Like "" Then
    strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & "        ***"
Else
    If Precon Like "EMB" Then
        Precon = "StationGlobals.Emb==1"
        strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & "     ***  =>Embarqué"
    ElseIf Precon Like "FCT" Then
        Precon = "StationGlobals.Emb==0"
        strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & "     ***  =>Fonctionnel"
    Else
        strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & " (if" & Precon & ") ***  "
    End If
End If

'vérification du timeout
If j + 4 = 16 Then
    If Cells(i, j + 4) <> "" Then
        timeout = Cells(i, j + 4)
    Else
        timeout = "StationGlobals.TimeOut"
    End If
End If

Select Case UCase(MotAction)
    Case UCase("Init_Task")
        Call ProcessingCB_Init_Task(Seq, nb_seq, compteur, Precon)
    Case UCase("Popup")
        Call ProcessingCB_Popup(Name_Var, Seq, nb_seq, compteur, Precon)
    Case UCase("Force")
        Call ProcessingCB_Force(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    Case UCase("UnForce")
        Call ProcessingCB_UnForce(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Precon)
    Case UCase("Write")
        Call ProcessingCB_Write(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    Case UCase("WriteNN")
        Call ProcessingCB_WriteNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
    Case UCase("Read")
        Call ProcessingCB_Read(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    Case UCase("Test")
        Call ProcessingCB_Test(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), timeout, Precon)
    Case UCase("ForceNN")
        Call ProcessingCB_ForceNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
    Case UCase("UnForceNN")
        Call ProcessingCB_UnForceNN(strLib, Seq, nb_seq, compteur, Name_Var, Precon)
    Case UCase("ReadNN")
        Call ProcessingCB_ReadNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
    Case UCase("TestNN")
        Call ProcessingCB_TestNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), timeout, Precon)
    Case UCase("TestAnaNN")
        Call ProcessingCB_TestAnaNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Cells(i, j + 3), Cells(i, j + 6), timeout, Precon)
    Case UCase("TestAna")
        Call ProcessingCB_TestAna(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Cells(i, j + 3), Cells(i, j + 6), timeout, Precon)
    Case UCase("Call")
        Call ProcessingSequenceCall(Cells(i, j) & " => " & Name_Var, Seq, nb_seq, compteur, Cells(i, j + 1), Cells(i, j + 2), Precon, Cells(i, j + 3))
    Case UCase("Wait")
        Call ProcessingNI_Wait(Cells(i, j), Seq, nb_seq, compteur, Cells(i, j + 3), Precon)
    Case UCase("Tempo_System")
        Call ProcessingNI_Wait(Cells(i, j), Seq, nb_seq, compteur, Cells(i, j + 3), Precon)
    'TEST ARRAY
    Case UCase("UnForceArray")
        Call ProcessingCB_UnForceArray(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Precon)
    Case UCase("UnForceArrayNN")
       Call ProcessingCB_UnForceArrayNN(strLib, Seq, nb_seq, compteur, Name_Var, Precon)
    Case UCase("ForceArrayAll")
        Call ProcessingCB_ForceArrayAll(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    Case UCase("ForceArrayElt")
        Call ProcessingCB_ForceArrayElt(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    Case UCase("ForceArrayAllNN")
        Call ProcessingCB_ForceArrayAllNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
    Case UCase("ForceArrayEltNN")
        Call ProcessingCB_ForceArrayEltNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
    Case UCase("TestArrayElt")
        Call ProcessingCB_TestArrayElt(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), timeout, Precon)
    Case UCase("TestArrayEltNN")
        Call ProcessingCB_TestArrayEltNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), timeout, Precon)
    'CB QUICK_ACCESS
    Case UCase("QA_ResetAll")
        Call ProcessingCB_QA_ResetAll(strLib, Seq, nb_seq, compteur, Precon)
    Case UCase("QA_UnForceAll")
        Call ProcessingCB_QA_UnForceAll(strLib, Seq, nb_seq, compteur, Precon)
    Case UCase("QA_ForceVar")
        Call ProcessingCB_QA_ForceVar(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    Case UCase("QA_UnForceVar")
        Call ProcessingCB_QA_UnForceVar(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Precon)
    'TEST MMI 150
    Case UCase("Test_MMI_150")
        Call ProcessingCallExecutable(Cells(i, j) & " => """ & Cells(i, j + 1) & """", Seq, nb_seq, compteur, "MMI_150 " & Cells(i, j + 1) & " " & Replace(Cells(i, j + 2), " ", "_"), Replace(ThisWorkbook.Worksheets("Parameters").Range("TestIHM_Exe_Path").Value, "\", "\\"), Precon)
    'TEST CMxEvol
    Case UCase("Test_CMxEvol")
        Call ProcessingCallExecutable(Cells(i, j) & " => """ & Cells(i, j + 1) & """", Seq, nb_seq, compteur, "CMxEvol " & Cells(i, j + 1) & " " & Replace(Cells(i, j + 2), " ", "_"), Replace(ThisWorkbook.Worksheets("Parameters").Range("TestIHM_Exe_Path").Value, "\", "\\"), Precon)
    'TEST DDU
    Case UCase("HMI_Start_Test")
        Call ProcessingTL_Init("", Seq, nb_seq, compteur, Name_Var, Precon) 'initialise putty
        Call ProcessingTL_SendCmd("", Seq, nb_seq, compteur, Name_Var, Precon) 'lance vnc depuis putty
        Call ProcessingTH_Init("", Seq, nb_seq, compteur, Name_Var, Precon) 'intialise vnc
        Call ProcessingTH_Start("", Seq, nb_seq, compteur, Name_Var, Precon) 'lance connexion DDU
    Case UCase("HMI_Stop_Test")
        Call ProcessingTH_Stop("", Seq, nb_seq, compteur, Name_Var, Precon)
    Case UCase("HMI_Send_Key")
        Call ProcessingTH_Keyboard("", Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
    'Pas Statement
    Case UCase("Statement")
        Call ProcessingStatement(Cells(i, j), Seq, nb_seq, compteur, Cells(i, j + 3), Precon)
    Case Else
        Call ProcessingLabel(Cells(i, j), Seq, nb_seq, compteur)
End Select

End Sub

'#####################################################      CallExe_HMI     ############################################################

Sub ProcessingCallExecutable(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, parametres As String, executablePath As String, Precon As String)
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CallExecutable"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "Executable = """ & executablePath & """"
    Seq(nb_seq).EnsembleLignes1.Add "Arguments = ""\""" & parametres & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "WaitCondition = ""WAIT_FOR_SPECIFIED_TIME""" 'spécifique TEST_MMI_150
    Seq(nb_seq).EnsembleLignes1.Add "TimeToWait = 240" 'spécifique TEST_MMI_150
    Seq(nb_seq).EnsembleLignes1.Add "TerminateOnAbort = True" 'spécifique TEST_MMI_150
    Seq(nb_seq).EnsembleLignes1.Add "ExitCodeStatusAction = ""EXIT_CODE_EQ_ZERO""" 'spécifique TEST_MMI_150
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'######################################################         CB_QA_ResetAll      #####################################################


Sub ProcessingCB_QA_ResetAll(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, Precon As String)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_QuickReset"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809432"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_QuickReset"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 464"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [1]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add "    "
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
    
End Sub

Sub ProcessingCB_QA_UnForceAll(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, Precon As String)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_QuickUnforceAll"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809433"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_QuickUnforceAll"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 468"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [1]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
    
End Sub

'#####################################################        CB_QuickForceProducer       #############################################

Sub ProcessingCB_QA_ForceVar(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, vForcedValue As Variant, Precon As String)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_QuickForceProducer"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
            Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809453"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_QuickForceProducer"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 548"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & vForcedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'#####################################################      CB_QuickUnForceProducer     ###################################################

Sub ProcessingCB_QA_UnForceVar(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, Precon As String)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_QuickUnForceProducer"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809454"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_QuickUnForceProducer"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 552"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
    
End Sub

'################################################       CB_UnForceArrayElement      ##########################################################

Sub ProcessingCB_UnForceArray(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, Precon As String)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_UnForceArrayElement"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
            Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809412"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_UnForceArrayElement"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 416"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    'Seq(nb_seq).EnsembleLignes1.ADD ""
    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    'Seq(nb_seq).EnsembleLignes1.ADD "InBuf = ""\""" & strInstanceName & "\"""""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'################################################       CB_UnForceArrayElementNN        #############################################

Sub ProcessingCB_UnForceArrayNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, Precon As String)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_UnForceArrayElementNN"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
            Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809413"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_UnForceArrayElementNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 420"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [2]"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'###############################################        CB_ForceAllElements        ####################################################

Sub ProcessingCB_ForceArrayAll(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, vForcedValue As Variant, Precon As String)
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_ForceAllElements"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809419"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_ForceAllElements"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 440"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & vForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""

    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    'Seq(nb_seq).EnsembleLignes1.ADD "InBuf = ""\""" & strInstanceName & "\"""""
    'Seq(nb_seq).EnsembleLignes1.ADD ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'###############################################################        CB_ForceArrayElement        ##########################################################

Sub ProcessingCB_ForceArrayElt(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, valueField As Variant, Precon As String)

    Dim ForcedValue As String
    Dim index As String
    Dim sep As Long
    sep = InStr(1, valueField, "/", vbBinaryCompare)
    
    index = Trim(Left(valueField, sep - 1))
    ForcedValue = Trim(Right(valueField, Len(valueField) - sep))
    ForcedValue = Mid(ForcedValue, 2, Len(ForcedValue) - 2) 'retire les crochets
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_ForceArrayElement"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809409"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_ForceArrayElement"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 408"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [5]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nIndex"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & ForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & ForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""

    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    'Seq(nb_seq).EnsembleLignes1.ADD "InBuf = ""\""" & strInstanceName & "\"""""
    'Seq(nb_seq).EnsembleLignes1.ADD ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'#############################################      CB_ForceAllElementsNN       ######################################################

Sub ProcessingCB_ForceArrayAllNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, vForcedValue As Variant, Precon As String)
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_ForceAllElementsNN"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
            Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809420"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_ForceAllElementsNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 444"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & vForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'##########################################################     CB_ForceArrayElementNN      #######################################################


Sub ProcessingCB_ForceArrayEltNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, valueField As Variant, Precon As String)
    
    Dim ForcedValue As String
    Dim index As String
    Dim sep As Long
    sep = InStr(1, valueField, "/", vbBinaryCompare)
    
    index = Trim(Left(valueField, sep - 1))
    ForcedValue = Trim(Right(valueField, Len(valueField) - sep))
    ForcedValue = Mid(ForcedValue, 2, Len(ForcedValue) - 2) 'retire les crochets
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_ForceArrayElementNN"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
            Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809410"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_ForceArrayElementNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 412"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nIndex"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & ForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & ForcedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'#########################################################      CB_TestArrayElement     #########################################################

Sub ProcessingCB_TestArrayElt(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, valueField As Variant, Optional lTimeout_ms As Variant, Optional Precon As String)

    Dim awaitedValue As String
    Dim index As String
    Dim sep As Long
    sep = InStr(1, valueField, "/", vbBinaryCompare)
    
    index = Trim(Left(valueField, sep - 1))
    awaitedValue = Trim(Right(valueField, Len(valueField) - sep))
    awaitedValue = Mid(awaitedValue, 2, Len(awaitedValue) - 2) 'retire les crochets
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_TestArrayElement"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809414"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_TestArrayElement"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 424"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [8]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[7] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[8] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nIndex"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""awaitedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & awaitedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & awaitedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nVariation"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 5"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (64-bit Floating Point)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nVariationPercent"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""100"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""100"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 5"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (64-bit Floating Point)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[6]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[7]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[8]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add "IsUserOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add "IsServerOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    'Seq(nb_seq).EnsembleLignes1.ADD "InBuf = ""\""" & strInstanceName & "\"""""
    'Seq(nb_seq).EnsembleLignes1.ADD ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'#######################################################        CB_TestArrayElementNN       #######################################################

Sub ProcessingCB_TestArrayEltNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, valueField As Variant, Optional lTimeout_ms As Variant, Optional Precon As String)

    Dim awaitedValue As String
    Dim index As String
    Dim sep As Long
    sep = InStr(1, valueField, "/", vbBinaryCompare)
    
    index = Trim(Left(valueField, sep - 1))
    awaitedValue = Trim(Right(valueField, Len(valueField) - sep))
    awaitedValue = Mid(awaitedValue, 2, Len(awaitedValue) - 2) 'retire les crochets

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_TestArrayElementNN"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
            Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809415"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_TestArrayElementNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 428"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [7]"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[7] = ""TYPE, AutomationParameter"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nIndex"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & index & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 16-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""awaitedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & awaitedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & awaitedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nVariation"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 5"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (64-bit Floating Point)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nVariationPercent"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""100"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""100"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 5"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (64-bit Floating Point)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"

    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[6]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[7]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add "IsUserOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add "IsServerOptional = True"
    
    'Seq(nb_seq).EnsembleLignes1.ADD ""
    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    'Seq(nb_seq).EnsembleLignes1.ADD "InBuf = ""\""" & strNickName & "\"""""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'####################################################       LABEL       ###########################################################

Sub ProcessingLabel(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur)
If sLigne <> "" Then
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """Label"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    Seq(nb_seq).EnsembleLignes1.Add "NoResult = False"
    Seq(nb_seq).EnsembleLignes1.Add "ConnectionLifetime = 4"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    If sLigne Like "MODE*" Then
        Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "--------------------      " & sLigne & "      --------------------" & """"
    Else
        Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "------------------------------------------------------------------" & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
    '--------------------------------------------------------------------------------------------------------------------------------
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """Label"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    Seq(nb_seq).EnsembleLignes1.Add "NoResult = False"
    Seq(nb_seq).EnsembleLignes1.Add "ConnectionLifetime = 4"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    If sLigne Like "MODE*" Then
        Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "--------------------      " & sLigne & "      --------------------" & """"
    Else
        Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "=>      " & sLigne & "      " & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
    '---------------------------------------------------------------------------------------------------------------------------------
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """Label"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    Seq(nb_seq).EnsembleLignes1.Add "NoResult = False"
    Seq(nb_seq).EnsembleLignes1.Add "ConnectionLifetime = 4"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    If sLigne Like "MODE*" Then
        Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "--------------------      " & sLigne & "      --------------------" & """"
    Else
        Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "------------------------------------------------------------------" & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End If
End Sub

Sub ProcessingCB_Init_Task(Seq() As SubSequence, nb_seq, compteur, Precon)

    Call ProcessingCB_Stop_Task(Seq, nb_seq, compteur, Precon)
    Call ProcessingCB_Start_Task(Seq, nb_seq, compteur, Precon)
    
End Sub

'###################################################        CB_Start_Task       #####################################################


Sub ProcessingCB_Start_Task(Seq() As SubSequence, nb_seq, compteur, Precon)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_StartTask"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
        If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    'Seq(nb_seq).EnsembleLignes1.ADD "StatusExpr Line0001 = ""Step.DataSource != \""Step.Result.PassFail\"" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"""
    'Seq(nb_seq).EnsembleLignes1.ADD "StatusExpr Line0002 = ""Fail ? \""Passed\"" : \""Failed\"" "" "
    
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809360"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_StartTask"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 220"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [1]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters"
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = ""Launch or resume a CB Functional Task"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'###########################################        CB_Stop_Task            ########################################################

Sub ProcessingCB_Stop_Task(Seq() As SubSequence, nb_seq, compteur, Precon)

    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_StopTask"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
        If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
        
    'Seq(nb_seq).EnsembleLignes1.ADD "StatusExpr Line0001 = ""Step.DataSource != \""Step.Result.PassFail\"" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"""
    'Seq(nb_seq).EnsembleLignes1.ADD "StatusExpr Line0002 = ""Fail ? \""Passed\"" : \""Failed\"""""
    'Seq(nb_seq).EnsembleLignes1.ADD ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809362"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_StopTask"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 228"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [1]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    'Seq(nb_seq).EnsembleLignes1.ADD "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    'Seq(nb_seq).EnsembleLignes1.ADD "%COMMENT = ""\n """
    'Seq(nb_seq).EnsembleLignes1.ADD ""
        
        
        
        
        
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = ""Stop CB Functional Task"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'###########################################        POPUP       ####################################################


Sub ProcessingCB_Popup(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, Precon As String)

Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """MessagePopup"""

Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
End If
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "TitleExpr = ""\""Intervention Manuelle\"""""
Seq(nb_seq).EnsembleLignes1.Add "MessageExpr = ""\""" & sLigne & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Modal = True"
Seq(nb_seq).EnsembleLignes1.Add ""

Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "%NAME = ""Message Popup"""
Seq(nb_seq).EnsembleLignes1.Add ""

'Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].RTS]"
'Seq(nb_seq).EnsembleLignes1.Add "Type = ""Num"""
'Seq(nb_seq).EnsembleLignes1.Add "OptimizeNonReentrantCalls = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "EPNameExpr = ""str"""
'Seq(nb_seq).EnsembleLignes1.Add "EPEnabledExpr = ""str"""
'Seq(nb_seq).EnsembleLignes1.Add "EPMenuHint = ""str"""
'Seq(nb_seq).EnsembleLignes1.Add "EPIgnoreClient = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "EPInitiallyHidden = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "EPCheckToSaveTitledFile = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "ShowEPAlways = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "ShowEPForFileWin = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "ShowEPForExeWin = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "ShowEPForEditorOnly = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "AllowIntExeOfEP = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "CopyStepsOnOverriding = ""Bool"""
'Seq(nb_seq).EnsembleLignes1.Add "Priority = ""Num"""
'Seq(nb_seq).EnsembleLignes1.Add ""

Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
compteur = compteur + 1

End Sub

Sub ProcessingCB_Force(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, vForcedValue As Variant, Precon As String)

      'Pour Forcer les variables sans mnémonique
Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_Force"""

Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
End If
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 et CB3.32"""
Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{60CA140F-FBED-44D2-A0DF-DBCB2D65E7C0}"""
Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809364"
Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_Force"""
Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""453.0"""
Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 204"
Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & vForcedValue & """"
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vForcedValue & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
Seq(nb_seq).EnsembleLignes1.Add ""

Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
compteur = compteur + 1
End Sub

'#################################################      CB_UnForce      ###################################################################

Sub ProcessingCB_UnForce(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, Precon As String)

      'Pour de-Forcer les variables sans mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_UnForce"""

    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809369"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_UnForce"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 252"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'##########################################################     CB_Write        #####################################################

Sub ProcessingCB_Write(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, nValue As Variant, Precon As String)

      'Pour écrire les variables sans mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_Write"""

    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809347"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_Write"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 188"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & nValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & nValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""

    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'#################################################      CB_WriteNN      ##################################################################

Sub ProcessingCB_WriteNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, nValue As Variant, Optional Precon As String)
    'Pour écrire une valeur dans une mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_SetNN"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809372"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_WriteNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 264"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nValue"""
    If UCase(nValue) = "FALSE" Or UCase(nValue) = "TRUE" Or IsNumeric(nValue) Then
        Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & nValue & """"
        Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & nValue & """"
    Else
        Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & nValue & "\"""""
        Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & nValue & "\"""""
    End If
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
End Sub

'###################################################################        CB_Read     ###################################################################

Sub ProcessingCB_Read(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, vReadValue As Variant, Precon)

    'Pour lire les variables sans mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_Read"""
        
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809346"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_Read"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 184"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vReadValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & vReadValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & vReadValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
        
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'###################################################        CB_Test     #########################################################

Sub ProcessingCB_Test(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, awaitedValue As Variant, Optional lTimeout_ms As Variant, Optional Precon As String)

    'Pour tester les variables sans mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_Test"""
        
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809350"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_Test"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 192"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [5]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""awaitedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & awaitedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & awaitedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add "IsUserOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add "IsServerOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
        
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'#######################################################        CB_TestAna      #################################################

Sub ProcessingCB_TestAna(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, awaitedValue As Variant, Optional maxValue, Optional nTolerance, Optional lTimeout_ms As Variant, Optional Precon As String)

    'Pour tester les variables analogique sans mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_TestAna"""
        
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809353"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_TestAna"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 200"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [7]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[7] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strInstanceName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strVariableName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strVariableName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""awaitedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & awaitedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & awaitedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""maxValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & IIf(IsMissing(maxValue), 0, maxValue) & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(IsMissing(maxValue), 0, maxValue) & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""nTolerance"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Stationglobals.Tolerance"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Stationglobals.Tolerance"""
    'Seq(nb_seq).EnsembleLignes1.ADD "ArgVal = """ & IIf(nTolerance Like "", "Stationglobals.Tolerance", nTolerance) & """"
    'Seq(nb_seq).EnsembleLignes1.ADD "ArgDisplayVal = ""\""" & IIf(nTolerance Like "", "Stationglobals.Tolerance", nTolerance) & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 4"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (32-bit Floating Point)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[6]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[7]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""lTimeout_ms"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & lTimeout_ms & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add "IsUserOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add "IsServerOptional = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strInstanceName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
        
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'######################################################     CB_ForceNN      ############################################################

Sub ProcessingCB_ForceNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, vForcedValue As Variant, Precon As String)
    
    'Pour Forcer les variables reseaux (avec Mnémonique)
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_ForceNN"""
        
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809368"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_ForceNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 248"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vForcedValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & vForcedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & vForcedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
        
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
    
End Sub

'##############################################     CB_UnForceNN        ########################################################

Sub ProcessingCB_UnForceNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, Precon As String)

    'Pour de-forcer les variables avec mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_UnForceNN"""
        
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809370"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_UnForceNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 256"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [2]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1

End Sub

'################################################       CB_ReadNN       ####################################################

Sub ProcessingCB_ReadNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, vReadValue, Precon As String)

    'Pour lire les variables avec mnémonique
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_ReadNN"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
    Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
    Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
    Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
    Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809371"
    Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_ReadNN"""
    Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
    Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
    Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
    Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
    Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 260"
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [3]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""vReadValue"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & vReadValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & vReadValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef  = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
    Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
    Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
    Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
    Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    compteur = compteur + 1
        
End Sub

'##################################################     CB_TestNN       ##########################################################

Sub ProcessingCB_TestNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As Variant, awaitedValue As Variant, Optional lTimeout_ms As Variant, Optional Precon As String)

      'Pour tester les variables avec mnémonique
Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_TestNN"""

Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
End If
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809375"
Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_TestNN"""
Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 276"
Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [4]"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""awaitedValue"""
If UCase(awaitedValue) = "FALSE" Or UCase(awaitedValue) = "TRUE" Or IsNumeric(awaitedValue) Then
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & awaitedValue & """"
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & awaitedValue & """"
Else
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & awaitedValue & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & awaitedValue & "\"""""
End If
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""lTimeout_ms"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & lTimeout_ms & """"
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & lTimeout_ms & """"
Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add "IsUserOptional = True"
Seq(nb_seq).EnsembleLignes1.Add "IsServerOptional = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strNickName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
Seq(nb_seq).EnsembleLignes1.Add ""


Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
compteur = compteur + 1

End Sub

'#########################################################      CB_TestAnaNN        ################################################################

Sub ProcessingCB_TestAnaNN(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strNickName As String, awaitedValue As Variant, Optional maxValue, Optional nTolerance, Optional lTimeout_ms As Variant, Optional Precon As String)

      'Pour tester les variables analogique avec mnémonique
Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """CB_TestAnaNN"""

Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
End If
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, AutomationStepAdditions"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
Seq(nb_seq).EnsembleLignes1.Add "ObjectVariable = ""Stationglobals.cb"""
Seq(nb_seq).EnsembleLignes1.Add "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
Seq(nb_seq).EnsembleLignes1.Add "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
Seq(nb_seq).EnsembleLignes1.Add "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
Seq(nb_seq).EnsembleLignes1.Add "CoClassName = ""ControlBuild"""
Seq(nb_seq).EnsembleLignes1.Add "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
Seq(nb_seq).EnsembleLignes1.Add "InterfaceName = ""_ControlBuild"""
Seq(nb_seq).EnsembleLignes1.Add "MemberType = 1"
Seq(nb_seq).EnsembleLignes1.Add "Member = 1610809374"
Seq(nb_seq).EnsembleLignes1.Add "MemberName = ""CB_TestAnaNN"""
Seq(nb_seq).EnsembleLignes1.Add "HasMemberInfo = True"
Seq(nb_seq).EnsembleLignes1.Add "Locale = 1036"
Seq(nb_seq).EnsembleLignes1.Add "TypeLibVersion = ""4ba.0"""
Seq(nb_seq).EnsembleLignes1.Add "InterfaceType = 1"
Seq(nb_seq).EnsembleLignes1.Add "VTableOffset = 272"
Seq(nb_seq).EnsembleLignes1.Add "%HI: Parameters = [6]"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters]"
Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add "%[6] = ""TYPE, AutomationParameter"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[0]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""strNickName"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & strNickName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & strNickName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[1]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""awaitedValue"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & awaitedValue & """"
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & awaitedValue & """"
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[2]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""maxValue"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & IIf(IsMissing(maxValue), 0, maxValue) & """"
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""\""" & IIf(IsMissing(maxValue), 0, maxValue) & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 12"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Variant"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[3]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""nTolerance"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Stationglobals.Tolerance"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Stationglobals.Tolerance"""
'Seq(nb_seq).EnsembleLignes1.ADD "ArgVal = """ & IIf(nTolerance Like "", "Stationglobals.Tolerance", nTolerance) & """"
'Seq(nb_seq).EnsembleLignes1.ADD "ArgDisplayVal = ""\""" & IIf(nTolerance Like "", "Stationglobals.Tolerance", nTolerance) & "\"""""
Seq(nb_seq).EnsembleLignes1.Add "Type = 4"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (32-bit Floating Point)"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[4]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""passFail"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.PassFail"""
Seq(nb_seq).EnsembleLignes1.Add "Type = 11"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Boolean"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[5]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""errorMsg"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = ""Step.Result.ReportText"""
Seq(nb_seq).EnsembleLignes1.Add "Type = 8"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""String"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 3"
Seq(nb_seq).EnsembleLignes1.Add "IsByRef = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parameters[6]]"
Seq(nb_seq).EnsembleLignes1.Add "Name = ""lTimeout_ms"""
Seq(nb_seq).EnsembleLignes1.Add "ArgVal = """ & lTimeout_ms & """"
Seq(nb_seq).EnsembleLignes1.Add "ArgDisplayVal = """ & lTimeout_ms & """"
Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
Seq(nb_seq).EnsembleLignes1.Add "DisplayType = ""Number (Signed 32-bit Integer)"""
Seq(nb_seq).EnsembleLignes1.Add "TypeValid = True"
Seq(nb_seq).EnsembleLignes1.Add "Direction = 1"
Seq(nb_seq).EnsembleLignes1.Add "IsUserOptional = True"
Seq(nb_seq).EnsembleLignes1.Add "IsServerOptional = True"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "InBuf = ""\""" & strNickName & "\"""""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
Seq(nb_seq).EnsembleLignes1.Add ""

Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
compteur = compteur + 1

End Sub

Sub ProcessingSequenceCall(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, NameSequence, SequenceFile, Precon As String, Optional parameters = "")

    'Pour les Appels de sequence
    Dim params() As String
    Dim i As Integer
    Dim param As Variant
          
If parameters <> "" Then
    params = Split(parameters, "][")
    For i = 0 To UBound(params, 1)
        If InStr(1, params(i), "[") Then
        'If param Like "[*" Then
            params(i) = Replace(params(i), "[", "")
        End If
        If InStr(1, params(i), "]") Then
            params(i) = Replace(params(i), "]", "")
        End If
    Next
End If

Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """SequenceCall"""

Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
End If
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, SeqCallStepAdditions"""
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData]"
Seq(nb_seq).EnsembleLignes1.Add "SFPath = """ & SequenceFile & """"
Seq(nb_seq).EnsembleLignes1.Add "SeqName = """ & NameSequence & """"
Seq(nb_seq).EnsembleLignes1.Add "%FLG: Prototype = 262144"
Seq(nb_seq).EnsembleLignes1.Add ""
Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData]"
Seq(nb_seq).EnsembleLignes1.Add "ActualArgs = Arguments"
Seq(nb_seq).EnsembleLignes1.Add "Prototype = Obj"
'Seq(nb_seq).EnsembleLignes1.Add "Trace = ""Don\t Change"""
Seq(nb_seq).EnsembleLignes1.Add ""
If parameters <> "" Then
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.ActualArgs]"
    i = 1
    For Each param In params
        Seq(nb_seq).EnsembleLignes1.Add "param" & i & " = ""TYPE, SequenceArgument"""
        i = i + 1
    Next
    Seq(nb_seq).EnsembleLignes1.Add ""
    i = 1
    For Each param In params
        Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.ActualArgs.param" & i & "]"
        Seq(nb_seq).EnsembleLignes1.Add "UseDef = False"
        Seq(nb_seq).EnsembleLignes1.Add "Expr = """ & param & """"
        Seq(nb_seq).EnsembleLignes1.Add ""
        i = i + 1
    Next
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Prototype]"
    i = 1
    For Each param In params
        Seq(nb_seq).EnsembleLignes1.Add "param" & i & " ="
        i = i + 1
    Next
    Seq(nb_seq).EnsembleLignes1.Add ""
    i = 1
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Prototype]"
    For Each param In params
        Seq(nb_seq).EnsembleLignes1.Add "%FLG: param" & i & " = 4"
        i = i + 1
    Next
    Seq(nb_seq).EnsembleLignes1.Add ""
End If


Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
Seq(nb_seq).EnsembleLignes1.Add ""

Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
compteur = compteur + 1

End Sub

'##################################################     NI_Wait     ####################################################

Sub ProcessingNI_Wait(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, TimeValue, Precon As String)
    
        'Pour les Wait
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """NI_Wait"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "SeqCallStepGroupIdx = -1"
    Seq(nb_seq).EnsembleLignes1.Add "TimeExpr = """ & TimeValue & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub

'#################################################          Init_HMI_VNC        ###################################################

Sub ProcessingTH_Init(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, Precon As String)
    
    Dim DDU_num As String
    Dim img_path As String
    Dim ref_build As String
    Dim shadowFile As String
    
    If Right(strVariableName, 1) >= 0 And Right(strVariableName, 1) <= 3 Then
        DDU_num = Right(strVariableName, 1)
    Else
        DDU_num = 0
    End If
    img_path = Replace(ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_Picture_Folder_Path").Value, "\", "\\")
    ref_build = ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_Ref_Mode").Value
    shadowFile = Replace(ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_Shadow_File_Path").Value, "\", "\\")
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """TH_Init"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "num_DDU = " & DDU_num
    Seq(nb_seq).EnsembleLignes1.Add "sImagesPath = """ & img_path & """"
    Seq(nb_seq).EnsembleLignes1.Add "bReferenceBuild = " & ref_build
    Seq(nb_seq).EnsembleLignes1.Add "sShadowFileConf = """ & shadowFile & """"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "DDU Initialisation : DDU" & DDU_num & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub

Sub ProcessingTH_Start(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, Precon As String)
    
    Dim DDU_num As String
    Dim remoteAdd As String
    Dim remotePass As String
    
    If Right(strVariableName, 1) >= 0 And Right(strVariableName, 1) <= 3 Then
        DDU_num = Right(strVariableName, 1)
    Else
        DDU_num = 0
    End If
    remoteAdd = ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_IP_Target").Value
    remotePass = ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_VNC_Password").Value
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """TH_Start"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "num_DDU = " & DDU_num
    Seq(nb_seq).EnsembleLignes1.Add "sRemoteEqtAddress = """ & remoteAdd & """"
    Seq(nb_seq).EnsembleLignes1.Add "sRemoteEqtPassword = """ & remotePass & """"
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "DDU Start : DDU" & DDU_num & "@" & remoteAdd & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub

'####################################################       Init_Putty      ####################################################


Sub ProcessingTL_Init(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, Precon As String)

    Dim DDU_num As String
    Dim remoteAdd As String
    Dim remotePass As String
    
    If Right(strVariableName, 1) >= 0 And Right(strVariableName, 1) <= 3 Then
        DDU_num = Right(strVariableName, 1)
    Else
        DDU_num = 0
    End If
    remoteAdd = ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_IP_Target").Value
    remotePass = ThisWorkbook.Worksheets("Parameters").Range("DDU" & DDU_num & "_VNC_Password").Value
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """TL_Init"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, FlexCStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "LibPath = ""C:\\Program Files\\ExtraPuTTY\\TestStand\\ExtraPuTTY.dll"""
    Seq(nb_seq).EnsembleLignes1.Add "Func = ""Connexion_F"""
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parms = [16]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[6] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[7] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[8] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[9] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[10] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[11] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[12] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[13] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[14] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[15] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[16] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Return Value"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""TargetName"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & remoteAdd & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""ConnexionId"""
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 5"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""1"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 16"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Login"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""root\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Password"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""ShowPuTTY"""
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 0"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""true"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 4"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[6]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Protocol"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[7]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""PortNumber"""
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 5"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[8]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""GenerateReport"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[9]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""CRLF"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[10]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""NewCRLF"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[11]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""ReportFileData"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""ThisContext.RunState.SequenceFile.Path"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[12]]"
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 0"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 20"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[13]]"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "StrPass = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[14]]"
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 0"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.Error.Occurred"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 20"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[15]]"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.Error.Code"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 16"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[16]]"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "StrPass = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.Error.Msg"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = ""Initialize Putty Connexion with " & remoteAdd & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub

'##################################################     SendCmd_Putty       ###################################################


Sub ProcessingTL_SendCmd(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, Precon As String)
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """TL_SendCmd"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If

    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "SData = ""TYPE, FlexCStepAdditions"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call]"
    Seq(nb_seq).EnsembleLignes1.Add "LibPath = ""C:\\Program Files\\ExtraPuTTY\\TestStand\\ExtraPuTTY.dll"""
    Seq(nb_seq).EnsembleLignes1.Add "Func = ""SendRcvData_F"""
    Seq(nb_seq).EnsembleLignes1.Add "%HI: Parms = [13]"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms]"
    Seq(nb_seq).EnsembleLignes1.Add "%[0] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[1] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[2] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[3] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[4] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[5] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[6] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[7] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[8] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[9] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[10] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[11] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[12] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add "%[13] = ""TYPE, FCParameter"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[0]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Return Value"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 3"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[1]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""ConnexionId"""
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 5"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""1"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[2]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Command"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""" & Replace(ThisWorkbook.Worksheets("Parameters").Range("Start_VNC_Exe_Path").Value, "\", "\\") & "\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[3]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Title"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[4]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Comments"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""\""\"""""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[5]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""TimeCapture"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[6]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""DataRcv"""
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "StrPass = 2"
    Seq(nb_seq).EnsembleLignes1.Add "StrSize = ""1000000"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.DataReceived"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[7]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""MaxSizeofData"""
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""1000000"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[8]]"
    Seq(nb_seq).EnsembleLignes1.Add "Name = ""Settings"""
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 5"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""0"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[9]]"
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 0"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.PassFail"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 20"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[10]]"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "StrPass = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.ReportText"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[11]]"
    Seq(nb_seq).EnsembleLignes1.Add "NumType = 0"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.Error.Occurred"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 20"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[12]]"
    Seq(nb_seq).EnsembleLignes1.Add "NumPass = 1"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.Error.Code"""
    Seq(nb_seq).EnsembleLignes1.Add "Flags = 16"
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS.SData.Call.Parms[13]]"
    Seq(nb_seq).EnsembleLignes1.Add "Type = 2"
    Seq(nb_seq).EnsembleLignes1.Add "StrPass = 2"
    Seq(nb_seq).EnsembleLignes1.Add "ArgVal = ""Step.Result.Error.Msg"""
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = ""Putty => Send command : start_vnc.sh"""
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1
    
End Sub

Sub ProcessingTH_Stop(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, Precon As String)
    
    Dim DDU_num As String
    
    If Right(strVariableName, 1) >= 0 And Right(strVariableName, 1) <= 3 Then
        DDU_num = Right(strVariableName, 1)
    Else
        DDU_num = 0
    End If
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """TH_Stop"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "num_DDU = " & DDU_num
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "DDU Stop : DDU" & DDU_num & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub

'#####################################################      Send_Cmd_HMI_VNC        ###########################################################

Sub ProcessingTH_Keyboard(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, strVariableName As String, strInstanceName As String, nValue As Variant, Precon As String)

    Dim DDU_num As String
    Dim key_Modifiers As String
    Dim key As String
    Dim key_Count As String
    Dim key_Wait As String
    
    Dim reg As RegExp
    Set reg = CreateObject("VBScript.RegExp")
    
    If Right(strVariableName, 1) >= 0 And Right(strVariableName, 1) <= 3 Then
        DDU_num = Right(strVariableName, 1)
    Else
        DDU_num = 0
    End If
        
    key_Modifiers = ""
    If InStr(1, strInstanceName, "[Alt]", vbTextCompare) > 0 Then
        key_Modifiers = key_Modifiers & "Alt"
    End If
    If InStr(1, strInstanceName, "[Shift]", vbTextCompare) > 0 Then
        key_Modifiers = key_Modifiers & "Shift"
    End If
    If InStr(1, strInstanceName, "[Ctrl]", vbTextCompare) > 0 Then
        key_Modifiers = key_Modifiers & "Ctrl"
    End If
    'key
    reg.Pattern = ".*\[Key=(.*)\].*"
    If reg.test(strInstanceName) Then
        key = reg.Replace(strInstanceName, "$1")
    Else
        key = ""
    End If
    'count
    reg.Pattern = ".*\[Rep_Count=(\d+)\].*"
    If reg.test(nValue) Then
        key_Count = reg.Replace(nValue, "$1")
    Else
        key_Count = 0
    End If
    'delay
    reg.Pattern = ".*\[Rep_Delay=(\d+)\].*"
    If reg.test(nValue) Then
        key_Wait = reg.Replace(nValue, "$1")
    Else
        key_Wait = 0
    End If
    
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """TH_Keyboard"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "num_DDU = " & DDU_num
    Seq(nb_seq).EnsembleLignes1.Add "sKeyModifiers = """ & key_Modifiers & """"
    Seq(nb_seq).EnsembleLignes1.Add "key = """ & key & """"
    Seq(nb_seq).EnsembleLignes1.Add "iKeyCount = " & key_Count
    Seq(nb_seq).EnsembleLignes1.Add "iKeyWait = " & key_Wait
    
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & "DDU send key " & key & " (" & key_Modifiers & ")" & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub

Sub ProcessingStatement(sLigne As String, ByRef Seq() As SubSequence, nb_seq, ByRef compteur, exprValue As String, Precon As String)
    
    'Pour le pas Statement
    Seq(nb_seq).EnsembleLignes0.Add "%[" & Seq(nb_seq).cpt & "] = Step"
    Seq(nb_seq).EnsembleLignes0.Add "%TYPE: %[" & Seq(nb_seq).cpt & "] = " & """Statement"""
    
    Seq(nb_seq).EnsembleLignes1.Add "[SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "].TS]"
    Seq(nb_seq).EnsembleLignes1.Add "Id = ""ID#:" & String$(22 - Len(Trim((Str$(compteur)))), "0") & Trim(Str$(compteur)) & """"
    Seq(nb_seq).EnsembleLignes1.Add "PostExpr = """ & exprValue & """"
    
    If Precon <> "" Then
        Seq(nb_seq).EnsembleLignes1.Add "PreCond = """ & Precon & """"
    End If
    Seq(nb_seq).EnsembleLignes1.Add ""
    Seq(nb_seq).EnsembleLignes1.Add "[DEF, SF.Seq[" & nb_seq - 1 & "].Main[" & Seq(nb_seq).cpt & "]]"
    Seq(nb_seq).EnsembleLignes1.Add "%NAME = """ & sLigne & """"
    Seq(nb_seq).EnsembleLignes1.Add ""
    
    compteur = compteur + 1
    Seq(nb_seq).cpt = Seq(nb_seq).cpt + 1

End Sub
