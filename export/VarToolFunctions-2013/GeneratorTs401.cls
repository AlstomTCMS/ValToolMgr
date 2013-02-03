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
Private idSalt As Integer ' Used only to generate an ID which is always different
Private alreadyGenerated As Boolean
Private Const TEMPLATE_DIR = "C:\\macros_alstom\\templates\\TestStand3\\"

Public MainSeq As CTestStandSeq
Public SeqList As CTestStandSeqContainer


Private Sub Class_Initialize()
    Debug.Print "Class Generator is initialised"
    alreadyGenerated = False
    idSalt = 1
End Sub


Public Sub writeScenario(file As String, o_testContainer As CTestContainer)
    filename = file
    
    If (alreadyGenerated) Then
        Call Class_Initialize
    End If
    
    Call genTsStructFromTestContainer(o_testContainer)

    Call OpenScenario
    Call WriteFileHeader
    
    
    
    Call WriteFileFooter
    
    Call CloseScenario
    alreadyGenerated = True
End Sub

' Converts a potentially complex tree structure to a standardized, linear TestStand sequence list
Private Sub genTsStructFromTestContainer(o_testContainer As CTestContainer)
    
    Set MainSeq = New CTestStandSeq
    Set SeqList = New CTestStandSeqContainer
    
    Dim SeqIdx As Integer
    
    For SeqIdx = 1 To o_testContainer.getCount
        Dim TestContainer As CTest
        Set TestContainer = o_testContainer.getTest(SeqIdx)
        
        Dim SubSeq As CTestStandSeq
        Set SubSeq = New CTestStandSeq
        
        Call genInstrListFromTest(SubSeq, TestContainer)
        
        Dim TsInstr As CTestStandInstr
        Set TsInstr = New CTestStandInstr
        
        TsInstr.category = TS_CALL
        TsInstr.Data = SubSeq.identifier
        
        Call MainSeq.AddInstr(TsInstr)
        Call SeqList.AddSequence(SubSeq)
    Next SeqIdx
End Sub

Private Sub genInstrListFromTest(ByRef SubSeq As CTestStandSeq, TestContainer As CTest)
    SubSeq.identifier = TestContainer.title
    
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

    Call AppendTemplateFile("0_Header.txt")
    
    Dim SeqIdx As Integer
    For SeqIdx = 1 To SeqList.getCount
        Print #filePtr, "%[" & SeqIdx & "] = Sequence"
    Next SeqIdx
    
    
    
    Call AppendTemplateFile("1_MainSequence_Locals.txt")
    
    Call AppendTemplateFile("1_MainSequence_Setup.txt")
    
    Call AppendTemplateFile("1_MainSequence_cleanup.txt")
    
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
    Call AppendTemplateFile("99_Footer.txt")
End Sub

Private Sub CloseScenario()
    Close #filePtr
    Debug.Print "Scenario is finished"
End Sub

Private Sub AppendTemplateFile(sInputFile As String)
    sInputFile = TEMPLATE_DIR & sInputFile
    
    'Declare the variables
    Dim sLineOfText As String
    Dim SourceNum As Long
    Dim DestNum As Long
    
    'If an error occurs, close the files and exit the sub
    On Error GoTo ErrHandler
    
    'Open the source text file
    SourceNum = FreeFile()
    Open sInputFile For Input As SourceNum
    
    'Include the following line if the first line of the source
    'text file is a header row and you don't want to append it to
    'the destination text file:
    'Line Input #SourceNum, sLineOfText
    
    'Read each line of the source file and append it to the destination file
    Do Until EOF(SourceNum)
        Line Input #SourceNum, sLineOfText
        Print #filePtr, sLineOfText
    Loop
    
CloseFile:
    'Close the source file
    Close #SourceNum
    Exit Sub
ErrHandler:
      MsgBox "Error # " & Err & ": " & Error(Err)
      Resume CloseFile
End Sub

'Function used only to generate unique ID's for TestStand
Private Function genIdentifier() As String
    genIdentifier = String$(22 - Len(Trim((Str$(idSalt)))), "0") & Trim(Str$(idSalt))
    idSalt = idSalt + 1
End Function
