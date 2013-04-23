using System;


using ValToolMgrDna.Interface;

namespace TestStandGen
{
    using Antlr.Runtime;
    using Antlr4.StringTemplate;
    using Antlr4.StringTemplate.Compiler;
    using Antlr4.StringTemplate.Misc;
    using System.IO;

    public class TestStandGen
    {
        private CTestContainer sequence;
        private string outFile;
        private string templatePath;

        private CTestStandSeq MainSeq;
        private CTestStandSeqContainer SeqList;
        private bool alreadyGenerated;
        private int idsalt;
        private TemplateGroup group;

        public static void genSequence(CTestContainer sequence, string outFile, string templatePath)
        {
            TestStandGen test = new TestStandGen(sequence, outFile, templatePath);
            
            test.writeScenario();
        }

        private TestStandGen(CTestContainer sequence, string outFile, string templatePath)
        {
            this.sequence = sequence;
            this.outFile = outFile;
            this.templatePath = templatePath;

            
            initialize();
        }

        private void initialize()
        {
            MainSeq = new CTestStandSeq();

            CTestStandInstr instr = new CTestStandInstr();
            instr.category = CTestStandInstr.categoryList.TS_CALL;

            MainSeq.Add(instr);

            SeqList = new CTestStandSeqContainer();
            alreadyGenerated = false;
            idsalt = 1;
        }

        private void writeScenario()
        {
            if (alreadyGenerated) initialize();

            genTsStructFromTestContainer(sequence);

            TemplateGroup group = new TemplateGroupDirectory(this.templatePath, '$', '$');
            try
            {
                ErrorBuffer errors = new ErrorBuffer();
                group.Listener = errors;
                group.Load();

                Template st = group.GetInstanceOf("MainTemplate");
                System.Collections.Generic.Dictionary<string, object> test = (System.Collections.Generic.Dictionary<string, object>)st.GetAttributes();
                st.Add("filename", TestStandAdapter.protectBackslashes(this.outFile));
                st.Add("nbOfSequences", this.SeqList.Count);

                st.Add("SequenceDeclare", this.MainSeq.ToArray());

                st.Add("SequenceList", this.SeqList.ToArray());

                string result = st.Render();

                if (errors.Errors.Count > 0)
                {
                    foreach (TemplateMessage m in errors.Errors)
                    {
                        Console.WriteLine(m.ToString());
                    }
                    Console.ReadLine();
                }

                StreamWriter output = new StreamWriter(this.outFile);

                output.Write(result);
                output.Close();
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
                Console.ReadLine();
            }

            this.alreadyGenerated = true;
        }

        /// <summary>
        /// Converts a potentially complex tree structure to a standardized, linear TestStand sequence list
        /// </summary>
        /// <param name="sequence">Sequence to convert</param>
        private void genTsStructFromTestContainer(CTestContainer sequence)
        {
            foreach(CTest test in sequence)
            {
                CTestStandSeq SubSeq = new CTestStandSeq();

                genInstrListFromTest(SubSeq, test);

                SeqList.Add(SubSeq);

                CTestStandInstr TsInstr = new CTestStandInstr();

                TsInstr.category = CTestStandInstr.categoryList.TS_CALL;
                TsInstr.Data = SubSeq.identifier;

                MainSeq.Add(TsInstr);

            }
        }

        private void genInstrListFromTest(CTestStandSeq SubSeq, CTest TestContainer)
        {
            SubSeq.identifier = TestContainer.title;
            SubSeq.Title = TestContainer.title;

            foreach (CStep step in TestContainer)
            {
                //SubSeq.Add(genTsLabel("===================================="));
            }

            //    SubSeq.identifier = TestContainer.title
            //    SubSeq.title = TestContainer.title

            //    Dim StepIdx As Integer
            //    For StepIdx = 1 To TestContainer.getCount
            //        Dim o_step As CStep
            //        Set o_step = TestContainer.getStep(StepIdx)
            //        SubSeq.AddInstr genTsLabel("====================================")
            //        SubSeq.AddInstr genTsLabel("======== Step " & StepIdx & " : " & o_step.title)
            //        SubSeq.AddInstr genTsLabel("====================================")

            //        Dim InstrIdx As Integer

            //        If (o_step.DescAction <> "") Then SubSeq.AddInstr genTsLabel("== Actions : " & o_step.DescAction)
            //        For InstrIdx = 1 To o_step.getActionCount
            //            SubSeq.AddInstr getTsEquivFromInstr(o_step.getAction(InstrIdx))
            //        Next InstrIdx

            //        If (o_step.DescCheck <> "") Then SubSeq.AddInstr genTsLabel("== Checks : " & o_step.DescCheck)
            //        For InstrIdx = 1 To o_step.getcheckCount
            //            SubSeq.AddInstr getTsEquivFromInstr(o_step.getCheck(InstrIdx))
            //        Next InstrIdx
            //    Next StepIdx
        }

        private CTestStandInstr genTsLabel(string p)
        {
            CTestStandInstr label = new CTestStandInstr();
            label.category = CTestStandInstr.categoryList.TS_LABEL;
            label.Data = p;
            return label;
        }

        private CTestStandInstr getTsEquivFromInstr(CInstruction o_inst)
        {
            return null;
            //        Select Case o_inst.category
            //          Case A_FORCE
            //              getTsEquivFromInstr.category = TS_FORCE
            //              Set getTsEquivFromInstr.Data = o_inst.Data
            //          Case A_UNFORCE
            //              getTsEquivFromInstr.category = TS_UNFORCE
            //              Set getTsEquivFromInstr.Data = o_inst.Data
            //          Case A_TEST
            //              getTsEquivFromInstr.category = TS_TEST
            //              Set getTsEquivFromInstr.Data = o_inst.Data
            //          Case A_WAIT
            //              getTsEquivFromInstr.category = TS_WAIT
            //              getTsEquivFromInstr.Data = o_inst.Data / 1000
            //          Case Else
            //              getTsEquivFromInstr.category = categoryList.UNKNOWN
            //    End Select
        }
    }

//Private Sub OpenScenario()
//    'If File Exists, its is moved to the same name with .old extension
//    If Dir(filename) <> vbNullString Then
//        FileSystem.FileCopy filename, filename & ".old"
//        FileSystem.Kill filename
//    End If

//    filePtr = FreeFile
//    Open filename For Output As #filePtr
//End Sub

//Private Sub WriteFileHeader()

//    Call AppendTemplateFile("0_1Header.txt")
//    Print #filePtr, "%HI: Seq = [" & SeqList.getCount & "]"
//    Call AppendTemplateFile("0_2Header.txt")
    
//     Dim SeqIdx As Integer
//    For SeqIdx = 1 To SeqList.getCount
//        Print #filePtr, "%[" & SeqIdx & "] = Sequence"
//    Next SeqIdx
//    Print #filePtr, ""
    
//    Call AppendTemplateFile("0_3Header.txt")
//    Print #filePtr, "%HI: Main = [" & SeqList.getCount - 1 & "]"
    
//    Call AppendTemplateFile("1_MainSequence_Locals.txt")
    
//    Call writeInstructions(0, MainSeq)
    
//    Call AppendTemplateFile("1_MainSequence_Setup.txt")
//    Call AppendTemplateFile("1_MainSequence_cleanup.txt")
    
//End Sub

//Private Sub writeInstructions(SequenceIdx As Integer, Sequence As CTestStandSeq)
//    Dim InstrIdx As Integer
    
//    If (Sequence.getCount > 0) Then Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main]"
//    For InstrIdx = 1 To Sequence.getCount
//        Print #filePtr, "%[" & InstrIdx - 1 & "] = Step"
//        Print #filePtr, "%TYPE: %[" & InstrIdx - 1 & "] = """ & Sequence.getInstr(InstrIdx).getCategoryAsText & """"
//    Next InstrIdx
    
//    If (Sequence.getCount > 0) Then Print #filePtr, ""
    
//    For InstrIdx = 1 To Sequence.getCount
//        Call writeSingleInstruction(SequenceIdx, InstrIdx - 1, Sequence.getInstr(InstrIdx))
//    Next InstrIdx
//End Sub

//Private Sub writeSingleInstruction(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)
//    Dim strLib As String
//    Dim Name_Var As String
//    Dim Precon As String
//    Dim timeout As String
    
//    Precon = ""
//    timeout = "StationGlobals.TimeOut"
    

//    'If j = 7 Then
//    '    Precon = Cells(i, j + 11)
//    'Else
//    '    Precon = Cells(i, j + 6)
//    'End If

//    'Name_Var = Cells(i, j + 1)
//    'If Name_Var Like "Entrée_*" Then
//    '    Name_Var = Replace(Name_Var, "Entrée_", "")
//    'Else
//    '    If Name_Var Like "Sortie_*" Then
//    '        Name_Var = Replace(Name_Var, "Sortie_", "")
//    '    End If
//    'End If

//    'If Precon Like "" Then
//    '    strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & "        ***"
//    'Else
//    '    If Precon Like "EMB" Then
//    '        Precon = "StationGlobals.Emb==1"
//    '        strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & "     ***  =>Embarqué"
//    '    ElseIf Precon Like "FCT" Then
//     '       Precon = "StationGlobals.Emb==0"
//   '         strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & "     ***  =>Fonctionnel"
//   '     Else
//  '         strLib = Cells(i, j) & "    ***     (" & Cells(i, j + 2) & ") " & Name_Var & " = " & Cells(i, j + 3) & " (if" & Precon & ") ***  "
//   '     End If
//   ' End If

//    'vérification du timeout
//    'If j + 4 = 16 Then
//    '    If Cells(i, j + 4) <> "" Then
//    '        timeout = Cells(i, j + 4)
//    '    Else
//    '        timeout = "StationGlobals.TimeOut"
//    '    End If
//    'End If

//    Select Case instruction.category
//        Case TS_LABEL
//            Call ProcessingLabel(SequenceIdx, InstructionIdx, instruction)
//        Case TS_FORCE
//            Call ProcessingCB_Force(SequenceIdx, InstructionIdx, instruction)
//        Case TS_UNFORCE
//            Call ProcessingCB_UnForce(SequenceIdx, InstructionIdx, instruction)
//        Case TS_TEST
//            Call ProcessingCB_Test(SequenceIdx, InstructionIdx, instruction)
//        Case TS_CALL
//            Call ProcessingSequenceCall(SequenceIdx, InstructionIdx, instruction)
//        Case TS_WAIT
//            Call ProcessingNI_Wait(SequenceIdx, InstructionIdx, instruction)
//        'Case UCase("Init_Task")
//        '    Call ProcessingCB_Init_Task(Seq, nb_seq, compteur, Precon)
//        'Case UCase("Popup")
//        '    Call ProcessingCB_Popup(Name_Var, Seq, nb_seq, compteur, Precon)
//        'Case UCase("Write")
//        '    Call ProcessingCB_Write(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
//        'Case UCase("WriteNN")
//        '    Call ProcessingCB_WriteNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
//        'Case UCase("Read")
//        '    Call ProcessingCB_Read(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
//        'Case UCase("ForceNN")
//        '    Call ProcessingCB_ForceNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
//        'Case UCase("UnForceNN")
//        '    Call ProcessingCB_UnForceNN(strLib, Seq, nb_seq, compteur, Name_Var, Precon)
//        'Case UCase("ReadNN")
//        '    Call ProcessingCB_ReadNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
//        'Case UCase("TestNN")
//        '    Call ProcessingCB_TestNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), timeout, Precon)
//        'Case UCase("TestAnaNN")
//        '    Call ProcessingCB_TestAnaNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Cells(i, j + 3), Cells(i, j + 6), timeout, Precon)
//        'Case UCase("TestAna")
//        '    Call ProcessingCB_TestAna(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Cells(i, j + 3), Cells(i, j + 6), timeout, Precon)
//        'Case UCase("Tempo_System")
//        '    Call ProcessingNI_Wait(Cells(i, j), Seq, nb_seq, compteur, Cells(i, j + 3), Precon)
//        'TEST ARRAY
//        'Case UCase("UnForceArray")
//        '    Call ProcessingCB_UnForceArray(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Precon)
//        'Case UCase("UnForceArrayNN")
//        '   Call ProcessingCB_UnForceArrayNN(strLib, Seq, nb_seq, compteur, Name_Var, Precon)
//        'Case UCase("ForceArrayAll")
//        '    Call ProcessingCB_ForceArrayAll(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
//        'Case UCase("ForceArrayElt")
//        '    Call ProcessingCB_ForceArrayElt(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
//        'Case UCase("ForceArrayAllNN")
//        '    Call ProcessingCB_ForceArrayAllNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
//        'Case UCase("ForceArrayEltNN")
//        '    Call ProcessingCB_ForceArrayEltNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), Precon)
//        'Case UCase("TestArrayElt")
//        '    Call ProcessingCB_TestArrayElt(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), timeout, Precon)
//        'Case UCase("TestArrayEltNN")
//        '    Call ProcessingCB_TestArrayEltNN(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 3), timeout, Precon)
//        'CB QUICK_ACCESS
//        'Case UCase("QA_ResetAll")
//        '    Call ProcessingCB_QA_ResetAll(strLib, Seq, nb_seq, compteur, Precon)
//        'Case UCase("QA_UnForceAll")
//        '    Call ProcessingCB_QA_UnForceAll(strLib, Seq, nb_seq, compteur, Precon)
//        'Case UCase("QA_ForceVar")
//        '    Call ProcessingCB_QA_ForceVar(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
//        'Case UCase("QA_UnForceVar")
//        '    Call ProcessingCB_QA_UnForceVar(strLib, Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Precon)
//        'TEST MMI 150
//        'Case UCase("Test_MMI_150")
//        '    Call ProcessingCallExecutable(Cells(i, j) & " => """ & Cells(i, j + 1) & """", Seq, nb_seq, compteur, "MMI_150 " & Cells(i, j + 1) & " " & Replace(Cells(i, j + 2), " ", "_"), Replace(ThisWorkbook.Worksheets("Parameters").Range("TestIHM_Exe_Path").value, "\", "\\"), Precon)
//        'TEST CMxEvol
//        'Case UCase("Test_CMxEvol")
//        '    Call ProcessingCallExecutable(Cells(i, j) & " => """ & Cells(i, j + 1) & """", Seq, nb_seq, compteur, "CMxEvol " & Cells(i, j + 1) & " " & Replace(Cells(i, j + 2), " ", "_"), Replace(ThisWorkbook.Worksheets("Parameters").Range("TestIHM_Exe_Path").value, "\", "\\"), Precon)
//        'TEST DDU
//        'Case UCase("HMI_Start_Test")
//        '    Call ProcessingTL_Init("", Seq, nb_seq, compteur, Name_Var, Precon) 'initialise putty
//        '    Call ProcessingTL_SendCmd("", Seq, nb_seq, compteur, Name_Var, Precon) 'lance vnc depuis putty
//        '    Call ProcessingTH_Init("", Seq, nb_seq, compteur, Name_Var, Precon) 'intialise vnc
//        '    Call ProcessingTH_Start("", Seq, nb_seq, compteur, Name_Var, Precon) 'lance connexion DDU
//        'Case UCase("HMI_Stop_Test")
//        '    Call ProcessingTH_Stop("", Seq, nb_seq, compteur, Name_Var, Precon)
//        'Case UCase("HMI_Send_Key")
//        '    Call ProcessingTH_Keyboard("", Seq, nb_seq, compteur, Name_Var, Cells(i, j + 2), Cells(i, j + 3), Precon)
//        'Pas Statement
//        'Case UCase("Statement")
//        '    Call ProcessingStatement(Cells(i, j), Seq, nb_seq, compteur, Cells(i, j + 3), Precon)
//        Case Else
//            MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
//    End Select
//End Sub

//Private Sub WriteFileFooter()
//    Call AppendTemplateFile("99_Footer.txt")
//End Sub

//Private Sub CloseScenario()
//    Close #filePtr
//    Debug.Print "Scenario is finished"
//End Sub

//Private Sub AppendTemplateFile(sInputFile As String)
//    sInputFile = TEMPLATE_DIR & sInputFile
    
//    'Declare the variables
//    Dim sLineOfText As String
//    Dim SourceNum As Long
//    Dim DestNum As Long
    
//    'If an error occurs, close the files and exit the sub
//    On Error GoTo ErrHandler
    
//    'Open the source text file
//    SourceNum = FreeFile()
//    Open sInputFile For Input As SourceNum
    
//    'Include the following line if the first line of the source
//    'text file is a header row and you don't want to append it to
//    'the destination text file:
//    'Line Input #SourceNum, sLineOfText
    
//    'Read each line of the source file and append it to the destination file
//    Do Until EOF(SourceNum)
//        Line Input #SourceNum, sLineOfText
//        Print #filePtr, sLineOfText
//    Loop
    
//CloseFile:
//    'Close the source file
//    Close #SourceNum
//    Exit Sub
//ErrHandler:
//      MsgBox "Error # " & Err & ": " & Error(Err)
//      Resume CloseFile
//End Sub

//'Function used only to generate unique ID's for TestStand
//Private Function genIdentifier() As String
//    genIdentifier = String$(22 - Len(Trim((Str$(idSalt)))), "0") & Trim(Str$(idSalt))
//    idSalt = idSalt + 1
//End Function

//'********************************************************************************************************************
//'                   Librairies séquence TS
//'********************************************************************************************************************
//'role = Ce module contient le code générique utilisées pour créer le fichier séquence TS
//'date_modification = 02/02/2013
//'remarque = Module Générique (fonctionne avec les FDT Fc et NT)
//'********************************************************************************************************************

//Private Sub WriteSequence(SequenceIdx As Integer, Sequence As CTestStandSeq)
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "]]"
//    Print #filePtr, "Parameters = Obj"
//    Print #filePtr, "Locals = Obj"
//    Print #filePtr, "Main = Objs"
//    Print #filePtr, "Setup = Objs"
//    Print #filePtr, "Cleanup = Objs"
//    Print #filePtr, "GotoCleanupOnFail = Bool"
//    Print #filePtr, "RecordResults = Bool"
//    Print #filePtr, "RTS = Obj"
//    'Print #filePtr, "Requirements = Obj"
//    Print #filePtr, "%NAME = """ & Sequence.identifier & """"
//    Print #filePtr, ""
    
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "]]"
//    Print #filePtr, "%FLG: Parameters = 4456448"
//    Print #filePtr, "%FLG: Locals = 4194304"
//    If (Sequence.getCount > 0) Then Print #filePtr, "%HI: Main = [" & Sequence.getCount - 1 & "]"
//    Print #filePtr, "%FLG: Main = 4194304"
//    Print #filePtr, "%FLG: Setup = 4194304"
//    Print #filePtr, "%FLG: Cleanup = 4194304"
//    Print #filePtr, "%FLG: GotoCleanupOnFail = 4194312"
//    Print #filePtr, "RecordResults = True"
//    Print #filePtr, "%FLG: RecordResults = 4194312"
//    Print #filePtr, "%FLG: RTS = 4456456"
//    'Print #filePtr, "%FLG: Requirements = 4456456"
//    Print #filePtr, ""

//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Locals]"
//    Print #filePtr, "ResultList = Objs"
//    Print #filePtr, ""
            
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Locals.ResultList]"
//    Print #filePtr, "%EPTYPE = TEResult"
//    Print #filePtr, ""

//    Call writeInstructions(SequenceIdx, Sequence)

//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].RTS]"
//    Print #filePtr, "Type = Num"
//    Print #filePtr, "OptimizeNonReentrantCalls = Bool"
//    Print #filePtr, "EPNameExpr = Str"
//    Print #filePtr, "EPEnabledExpr = Str"
//    Print #filePtr, "EPMenuHint = Str"
//    Print #filePtr, "EPIgnoreClient = Bool"
//    Print #filePtr, "EPInitiallyHidden = Bool"
//    Print #filePtr, "EPCheckToSaveTitledFile = Bool"
//    Print #filePtr, "ShowEPAlways = Bool"
//    Print #filePtr, "ShowEPForFileWin = Bool"
//    Print #filePtr, "ShowEPForExeWin = Bool"
//    Print #filePtr, "ShowEPForEditorOnly = Bool"
//    Print #filePtr, "AllowIntExeOfEP = Bool"
//    Print #filePtr, "CopyStepsOnOverriding = Bool"
//    Print #filePtr, "Priority = Num"
//    Print #filePtr, ""

//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].RTS]"
//    Print #filePtr, "%FLG: Type = 4194304"
//    Print #filePtr, "OptimizeNonReentrantCalls = True"
//    Print #filePtr, "%FLG: OptimizeNonReentrantCalls = 4194304"
//    Print #filePtr, "EPNameExpr = ""\""Unnamed Entry Point\"""""
//    Print #filePtr, "%FLG: EPNameExpr = 4194304"
//    Print #filePtr, "EPEnabledExpr = ""True"""
//    Print #filePtr, "%FLG: EPEnabledExpr = 4194304"
//    Print #filePtr, "%FLG: EPMenuHint = 4194304"
//    Print #filePtr, "%FLG: EPIgnoreClient = 4194304"
//    Print #filePtr, "%FLG: EPInitiallyHidden = 4194304"
//    Print #filePtr, "EPCheckToSaveTitledFile = True"
//    Print #filePtr, "%FLG: EPCheckToSaveTitledFile = 4194304"
//    Print #filePtr, "%FLG: ShowEPAlways = 4194304"
//    Print #filePtr, "ShowEPForFileWin = True"
//    Print #filePtr, "%FLG: ShowEPForFileWin = 4194304"
//    Print #filePtr, "%FLG: ShowEPForExeWin = 4194304"
//    Print #filePtr, "%FLG: ShowEPForEditorOnly = 4194304"
//    Print #filePtr, "%FLG: AllowIntExeOfEP = 4194304"
//    Print #filePtr, "CopyStepsOnOverriding = True"
//    Print #filePtr, "%FLG: CopyStepsOnOverriding = 4194304"
//    Print #filePtr, "Priority = 2953567917"
//    Print #filePtr, "%FLG: Priority = 4194304"
//    Print #filePtr, ""
            
//    'Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Requirements]"
//    'Print #filePtr, "Links = Strs"
//    'Print #filePtr, ""
            
//    'Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Requirements]"
//    'Print #filePtr, "%FLG: Links = 71303168"
//    'Print #filePtr, ""
//End Sub

//Private Sub ProcessingCB_Force(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)

//    Dim variable As CVariable
//    Set variable = instruction.Data

//    'Pour Forcer les variables sans mnémonique
//    Call printInstructionHeader(SequenceIdx, InstructionIdx, instruction)
    
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS]"
//    Print #filePtr, "SData = ""TYPE, AutomationStepAdditions"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call]"
//    Print #filePtr, "ObjectVariable = ""FileGlobals.cb1"""
//    Print #filePtr, "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
//    Print #filePtr, "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
//    Print #filePtr, "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
//    Print #filePtr, "CoClassName = ""ControlBuild"""
//    Print #filePtr, "Interface = ""{60CA140F-FBED-44D2-A0DF-DBCB2D65E7C0}"""
//    Print #filePtr, "InterfaceName = ""_ControlBuild"""
//    Print #filePtr, "MemberType = 1"
//    Print #filePtr, "Member = 1610809363"
//    Print #filePtr, "MemberName = ""CB_Force"""
//    Print #filePtr, "HasMemberInfo = True"
//    Print #filePtr, "Locale = 1036"
//    Print #filePtr, "TypeLibVersion = ""22b.0"""
//    Print #filePtr, "InterfaceType = 1"
//    'Print #filePtr, "VTableOffset = 204"
//    Print #filePtr, "%HI: Parameters = [4]"
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters]"
//    Print #filePtr, "%[0] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[1] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[2] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[3] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[4] = ""TYPE, AutomationParameter"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[0]]"
//    Print #filePtr, "Name = ""strInstanceName"""
//    Print #filePtr, "ArgVal = ""\""" & variable.path & "\"""""
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.path & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[1]]"
//    Print #filePtr, "Name = ""strVariableName"""
//    Print #filePtr, "ArgVal = ""\""" & variable.name & "\"""""
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.name & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[2]]"
//    Print #filePtr, "Name = ""vForcedValue"""
//    Print #filePtr, "ArgVal = """ & variable.getStringValue & """"
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.getStringValue & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[3]]"
//    Print #filePtr, "Name = ""passFail"""
//    Print #filePtr, "ArgVal = ""Step.Result.PassFail"""
//    Print #filePtr, "ArgDisplayVal = ""Step.Result.PassFail"""
//    Print #filePtr, "Type = 11"
//    Print #filePtr, "DisplayType = ""Boolean"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 3"
//    Print #filePtr, "IsByRef = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[4]]"
//    Print #filePtr, "Name = ""errorMsg"""
//    Print #filePtr, "ArgVal = ""Step.Result.ReportText"""
//    Print #filePtr, "ArgDisplayVal = ""Step.Result.ReportText"""
//    Print #filePtr, "Type = 8"
//    Print #filePtr, "DisplayType = ""String"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 3"
//    Print #filePtr, "IsByRef = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%COMMENT = ""\n """
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%NAME = ""Force " & variable.name & " at " & variable.value & """"
//    Print #filePtr, ""
//End Sub

//'#################################################      CB_UnForce      ###################################################################
//'Pour de-Forcer les variables sans mnémonique
//Private Sub ProcessingCB_UnForce(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)

//    Dim variable As CVariable
//    Set variable = instruction.Data

//    Call printInstructionHeader(SequenceIdx, InstructionIdx, instruction)
    
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS]"
//    Print #filePtr, "SData = ""TYPE, AutomationStepAdditions"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call]"
//    Print #filePtr, "ObjectVariable = ""Stationglobals.cb"""
//    Print #filePtr, "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
//    Print #filePtr, "ServerName = ""Interface between TestStand3.0 - CB4.0 - RP"""
//    Print #filePtr, "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
//    Print #filePtr, "CoClassName = ""ControlBuild"""
//    Print #filePtr, "Interface = ""{C60955F5-C36C-4941-8201-5DE04CB2FAD9}"""
//    Print #filePtr, "InterfaceName = ""_ControlBuild"""
//    Print #filePtr, "MemberType = 1"
//    Print #filePtr, "Member = 1610809369"
//    Print #filePtr, "MemberName = ""CB_UnForce"""
//    Print #filePtr, "HasMemberInfo = True"
//    Print #filePtr, "Locale = 1036"
//    Print #filePtr, "TypeLibVersion = ""4ba.0"""
//    Print #filePtr, "InterfaceType = 1"
//    Print #filePtr, "VTableOffset = 252"
//    Print #filePtr, "%HI: Parameters = [3]"
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters]"
//    Print #filePtr, "%[0] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[1] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[2] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[3] = ""TYPE, AutomationParameter"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[0]]"
//    Print #filePtr, "Name = ""strInstanceName"""
//    Print #filePtr, "ArgVal = ""\""" & variable.path & "\"""""
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.path & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[1]]"
//    Print #filePtr, "Name = ""strVariableName"""
//    Print #filePtr, "ArgVal = ""\""" & variable.name & "\"""""
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.name & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[2]]"
//    Print #filePtr, "Name = ""passFail"""
//    Print #filePtr, "ArgVal = ""Step.Result.PassFail"""
//    Print #filePtr, "ArgDisplayVal = ""Step.Result.PassFail"""
//    Print #filePtr, "Type = 11"
//    Print #filePtr, "DisplayType = ""Boolean"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 3"
//    Print #filePtr, "IsByRef = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[3]]"
//    Print #filePtr, "Name = ""errorMsg"""
//    Print #filePtr, "ArgVal = ""Step.Result.ReportText"""
//    Print #filePtr, "ArgDisplayVal = ""Step.Result.ReportText"""
//    Print #filePtr, "Type = 8"
//    Print #filePtr, "DisplayType = ""String"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 3"
//    Print #filePtr, "IsByRef = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "InBuf = ""\""" & variable.name & "\"""""
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%NAME = ""Test_" & idSalt & """"
//    Print #filePtr, ""
//End Sub

//'###################################################        CB_Test     #########################################################

//Private Sub ProcessingCB_Test(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)

//    Dim lTimeout_ms As String
//    lTimeout_ms = "FileGlobals.TimeOut"
//    Dim sLigne As String

//    'Pour tester les variables sans mnémonique
//    Dim variable As CVariable
//    Set variable = instruction.Data
    
//    Call printInstructionHeader(SequenceIdx, InstructionIdx, instruction)
    
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS]"
//    Print #filePtr, "SData = ""TYPE, AutomationStepAdditions"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call]"
//    Print #filePtr, "ObjectVariable = ""FileGlobals.cb1"""
//    Print #filePtr, "Server = ""{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"""
//    Print #filePtr, "ServerName = ""Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"""
//    Print #filePtr, "CoClass = ""{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"""
//    Print #filePtr, "CoClassName = ""ControlBuild"""
//    Print #filePtr, "Interface = ""{65D7DD1D-CEED-49EA-A31F-2A4F70D9A107}"""
//    Print #filePtr, "InterfaceName = ""_ControlBuild"""
//    Print #filePtr, "MemberType = 1"
//    Print #filePtr, "Member = 1610809350"
//    Print #filePtr, "MemberName = ""CB_Test"""
//    Print #filePtr, "HasMemberInfo = True"
//    Print #filePtr, "Locale = 1036"
//    Print #filePtr, "TypeLibVersion = ""454.0"""
//    Print #filePtr, "InterfaceType = 1"
//    Print #filePtr, "VTableOffset = 164"
//    Print #filePtr, "%HI: Parameters = [5]"
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters]"
//    Print #filePtr, "%[0] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[1] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[2] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[3] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[4] = ""TYPE, AutomationParameter"""
//    Print #filePtr, "%[5] = ""TYPE, AutomationParameter"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[0]]"
//    Print #filePtr, "Name = ""strInstanceName"""
//    Print #filePtr, "ArgVal = ""\""" & variable.path & "\"""""
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.path & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[1]]"
//    Print #filePtr, "Name = ""strVariableName"""
//    Print #filePtr, "ArgVal = ""\""" & variable.name & "\"""""
//    Print #filePtr, "ArgDisplayVal = ""\""" & variable.name & "\"""""
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[2]]"
//    Print #filePtr, "Name = ""awaitedValue"""
//    Print #filePtr, "ArgVal = """ & variable.getStringValue & """"
//    Print #filePtr, "ArgDisplayVal = """ & variable.getStringValue & """"
//    Print #filePtr, "Type = 12"
//    Print #filePtr, "DisplayType = ""Variant"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[3]]"
//    Print #filePtr, "Name = ""passFail"""
//    Print #filePtr, "ArgVal = ""Step.Result.PassFail"""
//    Print #filePtr, "ArgDisplayVal = ""Step.Result.PassFail"""
//    Print #filePtr, "Type = 11"
//    Print #filePtr, "DisplayType = ""Boolean"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 3"
//    Print #filePtr, "IsByRef = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[4]]"
//    Print #filePtr, "Name = ""errorMsg"""
//    Print #filePtr, "ArgVal = ""Step.Result.ReportText"""
//    Print #filePtr, "ArgDisplayVal = ""Step.Result.ReportText"""
//    Print #filePtr, "Type = 8"
//    Print #filePtr, "DisplayType = ""String"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 3"
//    Print #filePtr, "IsByRef = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Call.Parameters[5]]"
//    Print #filePtr, "Name = ""lTimeout_ms"""
//    Print #filePtr, "ArgVal = """ & lTimeout_ms & """"
//    Print #filePtr, "ArgDisplayVal = """ & lTimeout_ms & """"
//    Print #filePtr, "Type = 3"
//    Print #filePtr, "DisplayType = ""Number (Signed 32-bit Integer)"""
//    Print #filePtr, "TypeValid = True"
//    Print #filePtr, "Direction = 1"
//    Print #filePtr, "IsUserOptional = True"
//    Print #filePtr, "IsServerOptional = True"
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "InBuf = ""\""" & variable.path & "\"""""
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%NAME = ""Test " & variable.name & " at " & variable.getStringValue & """"
//    Print #filePtr, ""
//End Sub

//Private Sub ProcessingSequenceCall(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)

//    'Pour les Appels de sequence
//    'Dim params() As String
//    'Dim i As Integer
//    'Dim param As Variant
          
//    'If parameters <> "" Then
//    '    params = Split(parameters, "][")
//    '    For i = 0 To UBound(params, 1)
//    '        If InStr(1, params(i), "[") Then
//    '        'If param Like "[*" Then
//    '            params(i) = Replace(params(i), "[", "")
//    '        End If
//    '        If InStr(1, params(i), "]") Then
//    '            params(i) = Replace(params(i), "]", "")
//    '        End If
//    '    Next
//    'End If

//    Call printInstructionHeader(SequenceIdx, InstructionIdx, instruction, True)
    
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS]"
//    Print #filePtr, "SData = ""TYPE, SeqCallStepAdditions"""
//    Print #filePtr, ""
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData]"
//    'Print #filePtr, "SFPath = """ & SequenceFile & """"
    
//    Print #filePtr, "SeqName = """ & instruction.Data & """"
//    Print #filePtr, "UseCurFile = True"
//    'Print #filePtr, "%FLG: Prototype = 262144"
//    'Print #filePtr, "UsePrototype = True"
//    'Print #filePtr, ""
//    'Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData]"
//    'Print #filePtr, "ActualArgs = Arguments"
//    'Print #filePtr, "Prototype = Obj"
//    'Print #filePtr, "Trace = ""Don\t Change"""
//    Print #filePtr, ""
    
//    'If parameters <> "" Then
//    '
//    '    i = 1
//    '    For Each param In params
//     '       Print #filePtr, "param" & i & " = ""TYPE, SequenceArgument"""
//    '        i = i + 1
//    '    Next
//    '    Print #filePtr, ""
//    '    i = 1
//    '    For Each param In params
//    '        Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.ActualArgs.param" & i & "]"
//    '        Print #filePtr, "UseDef = False"
//    '        Print #filePtr, "Expr = """ & param & """"
//    '        Print #filePtr, ""
//    '        i = i + 1
//    '    Next
//    '    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Prototype]"
//    '    i = 1
//    '    For Each param In params
//    '        Print #filePtr, "param" & i & " ="
//    '        i = i + 1
//    '    Next
//    '    Print #filePtr, ""
//    '    i = 1
//    '    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS.SData.Prototype]"
//    '    For Each param In params
//    '        Print #filePtr, "%FLG: param" & i & " = 4"
//    '        i = i + 1
//    '    Next
//    '    Print #filePtr, ""
//    'End If


//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%NAME = """ & instruction.Data & """"
//    Print #filePtr, ""
//End Sub

//'##################################################     NI_Wait     ####################################################
//'Pour les Wait
//Private Sub ProcessingNI_Wait(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)
    
        
//    Call printInstructionHeader(SequenceIdx, InstructionIdx, instruction, True)
    
//    Dim timeStr As String
//    timeStr = instruction.Data
//    timeStr = Replace(timeStr, ",", ".")

//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "SeqCallStepGroupIdx = -1"
//    Print #filePtr, "TimeExpr = """ & timeStr & """"
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%NAME = ""Wait " & timeStr & " seconds"""
//    Print #filePtr, ""
//End Sub

//Private Sub ProcessingLabel(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr)
    
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS]"
//    Print #filePtr, "Id = ""ID#:" & genIdentifier() & """"
//    'Print #filePtr, "NoResult = False"
//    'Print #filePtr, "ConnectionLifetime = 4"
//    Print #filePtr, ""
//    Print #filePtr, "[DEF, SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "]]"
//    Print #filePtr, "%NAME = """ & instruction.Data & """"
//    Print #filePtr, ""
//End Sub

//Private Sub printInstructionHeader(SequenceIdx As Integer, InstructionIdx As Integer, instruction As CTestStandInstr, Optional skipEval As Boolean)
//    Print #filePtr, "[SF.Seq[" & SequenceIdx & "].Main[" & InstructionIdx & "].TS]"
//    Print #filePtr, "Id = ""ID#:" & genIdentifier() & """"
//    If (Not skipEval) Then
//        Print #filePtr, "StatusExpr Line0001 = ""Step.DataSource != \""Step.Result.PassFail\""? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.PassF"""
//        Print #filePtr, "StatusExpr Line0002 = ""ail ? \""Passed\"": \""Failed\"""""
//    End If

//    'If Precon <> "" Then
//    '    Print #filePtr, "PreCond = """ & Precon & """"
//    'End If
//    Print #filePtr, ""

}
