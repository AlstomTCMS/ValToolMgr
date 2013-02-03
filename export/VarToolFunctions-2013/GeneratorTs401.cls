VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneratorTs401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private filename As String
Private filePtr As Integer
Public testsCont As CTestContainer

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

Private Sub Class_Initialize()
    Debug.Print "Class Generator is initialised"
End Sub


Public Sub writeScenario(file As String, o_testContainer As CTestContainer)
    filename = file
    Set testsCont = o_testContainer
    
    Call OpenScenario
    Call WriteFileHeader
    
    ' TestStand Needs to predeclare sequences that we want to describe later
    Dim SequenceIdx As Integer
    For SequenceIdx = 0 To testsCont.getCount() - 1
        Print #filePtr, "%[" & SequenceIdx & "] = Sequence"
    Next SequenceIdx
    Print #filePtr, ""
    
    ' Declare High level steps
    Print #filePtr, "[SF.Seq[0]]"
    Print #filePtr, "%HI: Main = [" & testsCont.getCount() - 1 & "]"
    Print #filePtr, ""
    Print #filePtr, "[DEF, SF.Seq[0].Main]"

    'Implementing call to sub-sequences
    For SequenceIdx = 0 To testsCont.getCount() - 1
       Print #filePtr, "%[" & SequenceIdx & "] = Step"
       Print #filePtr, "%TYPE: %[" & SequenceIdx & "] = " & """SequenceCall"""
    Next SequenceIdx
    
    Print #filePtr, ""
    'For i = 1 To Seq(1).EnsembleLignes1.Count
    '   Print #fic, Seq(1).EnsembleLignes1(i)
    'Next i

    Print #filePtr, "[DEF, SF.Seq[0].RTS]"
    
    
    
    
    Call WriteFileFooter
    
    Call CloseScenario
End Sub

Private Sub OpenScenario()
    'If File Exists, its is moved to the same name with .old extension
    If Dir(filename) <> vbNullString Then
        FileSystem.FileCopy filename, filename & ".old"
        FileSystem.Kill filename
    End If

    filePtr = FreeFile
    Open filename For Output As #filePtr
End Sub

Private Sub WriteFileHeader()
    Print #filePtr, "[__Header__]"
    Print #filePtr, "Path = """ & filename & """"
    Print #filePtr, ""
    Print #filePtr, "[DEF, SF.Seq]"
End Sub

Private Sub InstrWriteSequence()

End Sub

Private Sub WriteActions()
    For j = 2 To nb_seq

            Print #filePtr, "[DEF, SF.Seq[" & j - 1 & "]]"
            Print #filePtr, "Parameters = Obj"
            Print #filePtr, "Locals = Obj"
            Print #filePtr, "Main = Objs"
            Print #filePtr, "Setup = Objs"
            Print #filePtr, "Cleanup = Objs"
            Print #filePtr, "GotoCleanupOnFail = Bool"
            Print #filePtr, "RecordResults = Bool"
            Print #filePtr, "RTS = Obj"
            Print #filePtr, "Requirements = Obj"
            Print #filePtr, "%NAME = """ & Seq(j).name & """"
            Print #filePtr, ""

            Print #filePtr, "[SF.Seq[" & j - 1 & "]]"
            Print #filePtr, "%FLG: Parameters = 4456448"
            Print #filePtr, "%FLG: Locals = 4194304"
            Print #filePtr, "%HI: Main = [" & Seq(j).cpt - 1 & "]"
            Print #filePtr, "%FLG: Main = 4194304"
            Print #filePtr, "%FLG: Setup = 4194304"
            Print #filePtr, "%FLG: Cleanup = 4194304"
            Print #filePtr, "%FLG: GotoCleanupOnFail = 4194312"
            Print #filePtr, "RecordResults = True"
            Print #filePtr, "%FLG: RecordResults = 4194312"
            Print #filePtr, "%FLG: RTS = 4456456"
            Print #filePtr, "%FLG: Requirements = 4456456"
            Print #filePtr, ""

            Print #filePtr, "[DEF, SF.Seq[" & j - 1 & "].Locals]"
            Print #filePtr, "ResultList = Objs"
            Print #filePtr, ""
            
            Print #filePtr, "[DEF, SF.Seq[" & j - 1 & "].Locals.ResultList]"
            Print #filePtr, "%EPTYPE = TEResult"
            Print #filePtr, ""

            Print #filePtr, "[DEF, SF.Seq[" & j - 1 & "].Main]"
            For i = 1 To Seq(j).EnsembleLignes0.Count
                    Print #filePtr, Seq(j).EnsembleLignes0(i)
            Next i
            Print #filePtr, ""
            For i = 1 To Seq(j).EnsembleLignes1.Count
                    Print #filePtr, Seq(j).EnsembleLignes1(i)
            Next i

            Print #filePtr, "[DEF, SF.Seq[" & j - 1 & "].RTS]"
            Print #filePtr, "Type = Num"
            Print #filePtr, "OptimizeNonReentrantCalls = Bool"
            Print #filePtr, "EPNameExpr = Str"
            Print #filePtr, "EPEnabledExpr = Str"
            Print #filePtr, "EPMenuHint = Str"
            Print #filePtr, "EPIgnoreClient = Bool"
            Print #filePtr, "EPInitiallyHidden = Bool"
            Print #filePtr, "EPCheckToSaveTitledFile = Bool"
            Print #filePtr, "ShowEPAlways = Bool"
            Print #filePtr, "ShowEPForFileWin = Bool"
            Print #filePtr, "ShowEPForExeWin = Bool"
            Print #filePtr, "ShowEPForEditorOnly = Bool"
            Print #filePtr, "AllowIntExeOfEP = Bool"
            Print #filePtr, "CopyStepsOnOverriding = Bool"
            Print #filePtr, "Priority = Num"
            Print #filePtr, ""

            Print #filePtr, "[SF.Seq[" & j - 1 & "].RTS]"
            Print #filePtr, "%FLG: Type = 4194304"
            Print #filePtr, "OptimizeNonReentrantCalls = True"
            Print #filePtr, "%FLG: OptimizeNonReentrantCalls = 4194304"
            Print #filePtr, "EPNameExpr = ""\""Unnamed Entry Point\"""""
            Print #filePtr, "%FLG: EPNameExpr = 4194304"
            Print #filePtr, "EPEnabledExpr = ""True"""
            Print #filePtr, "%FLG: EPEnabledExpr = 4194304"
            Print #filePtr, "%FLG: EPMenuHint = 4194304"
            Print #filePtr, "%FLG: EPIgnoreClient = 4194304"
            Print #filePtr, "%FLG: EPInitiallyHidden = 4194304"
            Print #filePtr, "EPCheckToSaveTitledFile = True"
            Print #filePtr, "%FLG: EPCheckToSaveTitledFile = 4194304"
            Print #filePtr, "%FLG: ShowEPAlways = 4194304"
            Print #filePtr, "ShowEPForFileWin = True"
            Print #filePtr, "%FLG: ShowEPForFileWin = 4194304"
            Print #filePtr, "%FLG: ShowEPForExeWin = 4194304"
            Print #filePtr, "%FLG: ShowEPForEditorOnly = 4194304"
            Print #filePtr, "%FLG: AllowIntExeOfEP = 4194304"
            Print #filePtr, "CopyStepsOnOverriding = True"
            Print #filePtr, "%FLG: CopyStepsOnOverriding = 4194304"
            Print #filePtr, "Priority = 2953567917"
            Print #filePtr, "%FLG: Priority = 4194304"
            Print #filePtr, ""
            
            Print #filePtr, "[DEF, SF.Seq[" & j - 1 & "].Requirements]"
            Print #filePtr, "Links = Strs"
            Print #filePtr, ""
            
            Print #filePtr, "[SF.Seq[" & j - 1 & "].Requirements]"
            Print #filePtr, "%FLG: Links = 71303168"
            Print #filePtr, ""
            
    Next j
End Sub

Private Sub WriteFileFooter()
    Print #filePtr, "[DEF, SF.FileglobalDefaults]"
End Sub

Private Sub CloseScenario()
    Close #filePtr
    Debug.Print "Scenario is finished"
End Sub
