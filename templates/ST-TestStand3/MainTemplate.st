MainTemplate(TestStandFile) ::= <<

[__Header__]
ProductName = "TestStand"
ProductVersion = 3.0.0.12373
CompatibleProdVersion = 3.0.0.0
Version = 172
Type = "SequenceFile"
Path = "C:\\macros_alstom\\test\\genTest-ref.seq"
ProductVersionString = "3.0 (3.0.0.12373)"
CharEncoding = "SBCS"

[DEF, %OBJROOT]
SF = SequenceFileData
Path = PathValue
CB_Init = StepType
CB_StartTask = StepType
StepTypeMenu = Obj
StepTypeSubstepsArray = Objs
EditSubstep = StepType
TEInf = Obj
Error = Obj
CommonResults = Obj
AutomationStepAdditions = AutomationModule
AutomationParameter = AutomationParameter
FlexCStepAdditions = FCModule
FCParameter = FCParameter
Expression = ExprValue
NoneStepAdditions = NoneModule
Label = StepType
SequenceCall = StepType
SeqCallStepAdditions = SeqCallModule
NI_Wait = StepType
PostSubstep = StepType
RP_Init = StepType
RP_Close = StepType
RP_StartServer = StepType
RP_StopServer = StepType
RP_ReadConfig = StepType
RP_ResetConfig = StepType
RP_StartAllEqt = StepType
DotNetStepAdditions = DotNetModule
DotNetParameter = DotNetParameter
Variable_Force = StepType
Variable_Read_Test = StepType
Statement = StepType

[DEF, SF]
Seq = Objs
FileGlobalDefaults = Obj
ModelFile = "TYPE, Path"
LoadOpt = Str
UnloadOpt = Str
Version = Str
BatchSync = Num
SFGlobalsScope = Num
Type = Num
ModelOption = Num
%NAME = "Data"

[SF]
%HI: Seq = [1]
%FLG: Seq = 4194304
%FLG: FileGlobalDefaults = 4194304
%FLG: ModelFile = 4194312
LoadOpt = "UseStepUnloadOpt"
%FLG: LoadOpt = 4194312
UnloadOpt = "UseStepUnloadOpt"
%FLG: UnloadOpt = 4194312
Version = "2.2.1.0"
%FLG: Version = 4194312
BatchSync = 1
%FLG: BatchSync = 4194312
%FLG: SFGlobalsScope = 4194312
%FLG: Type = 4194312
%FLG: ModelOption = 4194312

[DEF, SF.Seq]
%[0] = Sequence
%[1] = Sequence

[DEF, SF.Seq[0]]
Parameters = Obj
Locals = Obj
Main = Objs
Setup = Objs
Cleanup = Objs
GotoCleanupOnFail = Bool
RecordResults = Bool
RTS = Obj
%NAME = "MainSequence"

[SF.Seq[0]]
%FLG: Parameters = 4456448
%FLG: Locals = 4194304
%HI: Main = [0]
%FLG: Main = 4194304
%HI: Setup = [8]
%FLG: Setup = 4194304
%FLG: Cleanup = 4194304
%FLG: GotoCleanupOnFail = 4194312
RecordResults = True
%FLG: RecordResults = 4194312
%FLG: RTS = 4456456

[DEF, SF.Seq[0].Locals]
ResultList = Objs
cycle_TESTSTAND = Num
Report = Ref
lue = Bool
EACB_File = Str
HARD = Str

[DEF, SF.Seq[0].Locals.ResultList]
%EPTYPE = TEResult

[SF.Seq[0].Locals.cycle_TESTSTAND]
%NUMFMT = "%g"

[SF.Seq[0].Locals]
EACB_File = "C:\\Russian_2ES5_EnvTest_1_0_0\\Workbench\\build\\conf_banc_env\\embedded\\target_testbench\\EACB.xml"
HARD = "C:\\Russian_2ES5_EnvTest_1_0_0\\Workbench\\build\\conf_banc_env\\embedded\\target_testbench\\HARD.xml"

[DEF, SF.Seq[0].Main]
%[0] = Step
%TYPE: %[0] = "SequenceCall"

[SF.Seq[0].Main[0].TS]
Id = "ID#:0000000000000000000543"

[DEF, SF.Seq[0].Main[0].TS]
SData = "TYPE, SeqCallStepAdditions"

[SF.Seq[0].Main[0].TS.SData]
SeqName = "Test_1.1"
UseCurFile = True

[DEF, SF.Seq[0].Main[0]]
%NAME = "Call to subsequence Test_1.1"

[DEF, SF.Seq[0].Setup]
%[0] = Step
%TYPE: %[0] = "Statement"
%[1] = Step
%TYPE: %[1] = "Statement"
%[2] = Step
%TYPE: %[2] = "Statement"
%[3] = Step
%TYPE: %[3] = "Statement"
%[4] = Step
%TYPE: %[4] = "Statement"
%[5] = Step
%TYPE: %[5] = "Statement"
%[6] = Step
%TYPE: %[6] = "Statement"
%[7] = Step
%TYPE: %[7] = "Statement"
%[8] = Step
%TYPE: %[8] = "Statement"

[SF.Seq[0].Setup[0].TS]
Id = "ID#:sFBjEI3Am0+svYDDcGLg0C"
PostExpr = "FileGlobals.WebMaintenance = StationGlobals.TrainTracer.WebMaintenance"

[DEF, SF.Seq[0].Setup[0]]
%NAME = "Copy Traintracer project references"

[SF.Seq[0].Setup[1].TS]
Id = "ID#:OZLBiCRcQUO/lDbyxcTkVB"
PostExpr = "FileGlobals.CbEmb1= StationGlobals.ControlBuildPtr.CbPtr_emb1"

[DEF, SF.Seq[0].Setup[1]]
%NAME = "Copy Embedded1 project references"

[SF.Seq[0].Setup[2].TS]
Id = "ID#:zqkp4ZN6ZUSItjai6kkc6C"
PostExpr = "FileGlobals.CbEmb2= StationGlobals.ControlBuildPtr.CbPtr_emb2"

[DEF, SF.Seq[0].Setup[2]]
%NAME = "Copy Embedded2 project references"

[SF.Seq[0].Setup[3].TS]
Id = "ID#:uG7yCx6kg0idI/WnRVHhIC"
PostExpr = "FileGlobals.CbEmb3= StationGlobals.ControlBuildPtr.CbPtr_emb3"

[DEF, SF.Seq[0].Setup[3]]
%NAME = "Copy Embedded3 project references"

[SF.Seq[0].Setup[4].TS]
Id = "ID#:feYbtVm5t0qrooqZV1PeIC"
PostExpr = "FileGlobals.CbEmb4= StationGlobals.ControlBuildPtr.CbPtr_emb4"

[DEF, SF.Seq[0].Setup[4]]
%NAME = "Copy Embedded4 project references"

[SF.Seq[0].Setup[5].TS]
Id = "ID#:YjKdiRHgQE6rJO/BcUC4XC"
PostExpr = "FileGlobals.CbEmb5= StationGlobals.ControlBuildPtr.CbPtr_emb5"

[DEF, SF.Seq[0].Setup[5]]
%NAME = "Copy Embedded5 project references"

[SF.Seq[0].Setup[6].TS]
Id = "ID#:Vfo+S7kSHE2oSV5skDRwLA"
PostExpr = "FileGlobals.CbEmb6= StationGlobals.ControlBuildPtr.CbPtr_emb6"

[DEF, SF.Seq[0].Setup[6]]
%NAME = "Copy Embedded6 project references"

[SF.Seq[0].Setup[7].TS]
Id = "ID#:KdoNygbjhkyh4Yeihz2OMC"
PostExpr = "FileGlobals.CbEmb7= StationGlobals.ControlBuildPtr.CbPtr_emb7"

[DEF, SF.Seq[0].Setup[7]]
%NAME = "Copy Embedded7 project references"

[SF.Seq[0].Setup[8].TS]
Id = "ID#:Dz0YUMfUbkajteWnP2+7fB"
PostExpr = "FileGlobals.CbEmb8= StationGlobals.ControlBuildPtr.CbPtr_emb8"

[DEF, SF.Seq[0].Setup[8]]
%NAME = "Copy Embedded8 project references"

[DEF, SF.Seq[0].RTS]
Type = Num
OptimizeNonReentrantCalls = Bool
EPNameExpr = Str
EPEnabledExpr = Str
EPMenuHint = Str
EPIgnoreClient = Bool
EPInitiallyHidden = Bool
EPCheckToSaveTitledFile = Bool
ShowEPAlways = Bool
ShowEPForFileWin = Bool
ShowEPForExeWin = Bool
ShowEPForEditorOnly = Bool
AllowIntExeOfEP = Bool
CopyStepsOnOverriding = Bool
Priority = Num

[SF.Seq[0].RTS]
%FLG: Type = 4194304
OptimizeNonReentrantCalls = True
%FLG: OptimizeNonReentrantCalls = 4194304
EPNameExpr = "\"Unnamed Entry Point\""
%FLG: EPNameExpr = 4194304
EPEnabledExpr = "True"
%FLG: EPEnabledExpr = 4194304
%FLG: EPMenuHint = 4194304
%FLG: EPIgnoreClient = 4194304
%FLG: EPInitiallyHidden = 4194304
EPCheckToSaveTitledFile = True
%FLG: EPCheckToSaveTitledFile = 4194304
%FLG: ShowEPAlways = 4194304
ShowEPForFileWin = True
%FLG: ShowEPForFileWin = 4194304
%FLG: ShowEPForExeWin = 4194304
%FLG: ShowEPForEditorOnly = 4194304
%FLG: AllowIntExeOfEP = 4194304
CopyStepsOnOverriding = True
%FLG: CopyStepsOnOverriding = 4194304
Priority = 2953567917
%FLG: Priority = 4194304

$TestStandFile.Sequences:{s |$SequenceBody(i, s)$}$
[DEF, SF.FileGlobalDefaults]
Perf = Ref
FIP = Ref
TH = Ref
TH_ED = Ref
HPTS = Ref
opc = Ref
outils = Ref
Pilotage = Ref
Mmi = Ref
Matrix = Ref
Power = Ref
Analyze = Ref
IHMConsole = Ref
IHMAcquisition = Ref
LS = Ref
cb1 = Ref
cb_1 = Ref
trameParam = Str
LectureBool = Bool
Type_simu = Num
Flange_Clock_Time = Bool
Result_Button = Num
InitDone = Bool
DDU_ScreenShot = Bool
DDU_ip = Str
cycle_CB = Num
TimeOut = Num
Var1 = Str
Var2 = Str
Var3 = Str
Var4 = Str
Var5 = Str
Var6 = Str
Var7 = Str
Var8 = Str
Var9 = Str
Var10 = Str
Var11 = Str
WebMaintenance = Ref
CbEnv = Ref
CbEnv1 = Ref
CbEnv2 = Ref
CbEnv3 = Ref
CbEnv4 = Ref
CbEnv5 = Ref
CbEnv6 = Ref
CbEnv7 = Ref
CbEnv8 = Ref
CbEmb = Ref
CbEmb1 = Ref
CbEmb2 = Ref
CbEmb3 = Ref
CbEmb4 = Ref
CbEmb5 = Ref
CbEmb6 = Ref
CbEmb7 = Ref
CbEmb8 = Ref

[SF.FileGlobalDefaults]
DDU_ScreenShot = True
DDU_ip = "10.22.98.17"
cycle_CB = 0.05
TimeOut = 5000

[%TYPES]
Path = "Path"
CB_Init = "CB_Init"
CB_StartTask = "CB_StartTask"
StepTypeMenu = "StepTypeMenu"
StepTypeSubstepsArray = "StepTypeSubstepsArray"
EditSubstep = "EditSubstep"
TEInf = "TEInf"
Error = "Error"
CommonResults = "CommonResults"
AutomationStepAdditions = "AutomationStepAdditions"
AutomationParameter = "AutomationParameter"
FlexCStepAdditions = "FlexCStepAdditions"
FCParameter = "FCParameter"
Expression = "Expression"
NoneStepAdditions = "NoneStepAdditions"
Label = "Label"
SequenceCall = "SequenceCall"
SeqCallStepAdditions = "SeqCallStepAdditions"
NI_Wait = "NI_Wait"
PostSubstep = "PostSubstep"
RP_Init = "RP_Init"
RP_Close = "RP_Close"
RP_StartServer = "RP_StartServer"
RP_StopServer = "RP_StopServer"
RP_ReadConfig = "RP_ReadConfig"
RP_ResetConfig = "RP_ResetConfig"
RP_StartAllEqt = "RP_StartAllEqt"
DotNetStepAdditions = "DotNetStepAdditions"
DotNetParameter = "DotNetParameter"
Variable_Force = "Variable_Force"
Variable_Read_Test = "Variable_Read_Test"
Statement = "Statement"

[DEF, Path]
%LOCATION = "StdTypes"
%ROOT_TYPE = True

[Path]
%TIMESTAMP = 1060111435
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194304
%TYPE_FLG = 33554436

[DEF, CB_Init]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[CB_Init]
%HI: Substeps = [0]
DescriptionFormat Line0001 = "\"Initialize CB with the project \" + Step.TS.SData.Call.Parameters[1].ArgVal + \" of path \"  + Step.TS.SData.Call.Para"
DescriptionFormat Line0002 = "meters[0].ArgVal + \", and with the application \"  + Step.TS.SData.Call.Parameters[2].ArgVal + \" and the functional ta"
DescriptionFormat Line0003 = "sk \"   + Step.TS.SData.Call.Parameters[3].ArgVal "
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Initialize CB\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1201268994
%VERSION = "0.0.0.12"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, CB_Init.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, CB_Init.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[CB_Init.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, CB_Init.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[CB_Init.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[CB_Init.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[CB_Init.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[CB_Init.Substeps[0]]
%INSTFLG: TS = 0

[CB_Init.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, CB_Init.Substeps[0]]
%NAME = "Edit"

[CB_Init.Menu]
ItemName = "\"Initialize CB\""
Group = "Start\\\\Stop CB"

[CB_Init.TS]
Icon = "ExprBrowse.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, CB_Init.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[CB_Init.TS.SData.Call]
ObjectVariable = "FileGlobals.cb"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"
CoClassName = "ControlBuild"
Interface = "{BAB4E7D4-DE84-4A87-8E5E-EC7ABEC2963C}"
InterfaceName = "_ControlBuild"
MemberType = 1
Member = 1610809344
MemberName = "CB_Init"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "30f.0"
InterfaceType = 1
VTableOffset = 148
%HI: Parameters = [5]

[DEF, CB_Init.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"
%[3] = "TYPE, AutomationParameter"
%[4] = "TYPE, AutomationParameter"
%[5] = "TYPE, AutomationParameter"

[CB_Init.TS.SData.Call.Parameters[0]]
Name = "strPathName"
ArgVal = "FileGlobals.CB_Path_Name"
ArgDisplayVal = "FileGlobals.CB_Path_Name"
Type = 12
DisplayType = "Variant"
TypeValid = True
Direction = 1

[CB_Init.TS.SData.Call.Parameters[1]]
Name = "strProjectName"
ArgVal = "FileGlobals.CB_Project_Name"
ArgDisplayVal = "FileGlobals.CB_Project_Name"
Type = 12
DisplayType = "Variant"
TypeValid = True
Direction = 1

[CB_Init.TS.SData.Call.Parameters[2]]
Name = "strApplicationName"
ArgVal = "FileGlobals.CB_Application_Name"
ArgDisplayVal = "FileGlobals.CB_Application_Name"
Type = 12
DisplayType = "Variant"
TypeValid = True
Direction = 1

[CB_Init.TS.SData.Call.Parameters[3]]
Name = "strFunctionalTaskName"
ArgVal = "FileGlobals.CB_Task_Name"
ArgDisplayVal = "FileGlobals.CB_Task_Name"
Type = 12
DisplayType = "Variant"
TypeValid = True
Direction = 1

[CB_Init.TS.SData.Call.Parameters[4]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[CB_Init.TS.SData.Call.Parameters[5]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, CB_Init.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[CB_Init.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, CB_StartTask]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[CB_StartTask]
%HI: Substeps = [0]
DescriptionFormat = "\"Launch or resume a CB Functional Task\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Launch or resume a CB Functional Task\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1220858404
%VERSION = "2.8.0.0"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, CB_StartTask.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, CB_StartTask.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[CB_StartTask.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, CB_StartTask.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[CB_StartTask.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[CB_StartTask.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[CB_StartTask.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[CB_StartTask.Substeps[0]]
%INSTFLG: TS = 0

[CB_StartTask.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, CB_StartTask.Substeps[0]]
%NAME = "Edit"

[CB_StartTask.Menu]
ItemName = "\"Launch or resume a CB Functional Task\""
Group = "Tasks"

[CB_StartTask.TS]
Icon = "ExprBrowse.ico"
%INSTOVRD: SData = 5177368
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, CB_StartTask.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[CB_StartTask.TS.SData.Call]
ObjectVariable = "FileGlobals.cb"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand et CB, Fip, Hpts, Mmi, Matrix"
CoClass = "{B51A0060-36D9-4E5C-AB1A-65720FD2E9CA}"
CoClassName = "ControlBuild"
Interface = "{F90181E8-DF9C-4EF3-BB69-04AB9FD5BD3C}"
InterfaceName = "_ControlBuild"
MemberType = 1
Member = 1610809355
MemberName = "CB_StartTask"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "1d5.0"
%HI: Parameters = [1]

[DEF, CB_StartTask.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[CB_StartTask.TS.SData.Call.Parameters[0]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[CB_StartTask.TS.SData.Call.Parameters[1]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, CB_StartTask.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[CB_StartTask.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, StepTypeMenu]
%LOCATION = "StdTypes"
CanBeSubstepType = Bool
CanOnlyBeSubstepType = Bool
Category = Str
ItemName = "TYPE, Expression"
SingularItemName = Str
SeparatorBeforeCategory = Bool
SeparatorBeforeItemName = Bool
Group = Str
%ROOT_TYPE = True

[StepTypeMenu]
Category = "\"\""
ItemName = "\"\""
SingularItemName = "\"\""
%TIMESTAMP = 1059680189
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194328
%TYPE_FLG = 33554436
%INSTFLG = 4718616
%INSTOVRD = 4718616

[DEF, StepTypeSubstepsArray]
%LOCATION = "StdTypes"
%ROOT_TYPE = True

[StepTypeSubstepsArray]
%TIMESTAMP = 1059680189
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194328
%TYPE_FLG = 33554436
%INSTFLG = 4718616
%INSTOVRD = 4849688

[DEF, EditSubstep]
%LOCATION = "StepTypes"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
Substeps = "TYPE, StepTypeSubstepsArray"
CodeTemplates = Str
%FLG: CodeTemplates = 524288
TS = "TYPE, TEInf"
Result = Obj
MenuName = "TYPE, Expression"
SupportsReadOnly = Bool
%ROOT_TYPE = True

[EditSubstep]
DescriptionFormat = "\"%ModuleDescription\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "ResStr(\"NI_STEPTYPES\", \"EDIT_DEF_SUBSTEP_NAME\")"
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
CodeTemplates Line0001 = "DefaultLabVIEW|DefaultCVI|DefaultVB.NET|DefaultCSharp.NET|DefaultC++.NET|DefaultVC++_Template|DefaultHTB72_Template|Defa"
CodeTemplates Line0002 = "ultHTB80_Template|Default_Template"
%FLG: CodeTemplates = 4194328
%INSTFLG: TS = 4456472
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%TIMESTAMP = 1063730206
%VERSION = "3.0.0.12374"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%TYPE_FLG = 33554438
%INSTFLG = 4194304
%INSTOVRD = 4194304

[EditSubstep.Menu]
CanBeSubstepType = True
CanOnlyBeSubstepType = True
ItemName = "ResStr(\"NI_STEPTYPES\", \"EDIT_SUBSTEP_MENU_ITEM_NAME\")"
Group = "EditSubsteps"

[DEF, EditSubstep.TS]
SData = Obj
%FLG: SData = 2097152

[EditSubstep.TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7274521

[DEF, EditSubstep.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"

[EditSubstep.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304

[DEF, TEInf]
%LOCATION = "StdTypes"
Id = Str
Icon = Str
SData = Obj
%FLG: SData = 2097152
PreCond = "TYPE, Expression"
LoadOpt = Str
UnloadOpt = Str
Mode = Str
WindowActivation = Str
NoResult = Bool
StepFCSeqF = Bool
IgnoreRTE = Bool
UseMutex = Bool
MutexNameOrRef = "TYPE, Expression"
BatchSyncOpt = Num
SwitchEnabled = Bool
VirtualDeviceName = "TYPE, Expression"
SwitchOperation = Num
RouteGroupConnect = "TYPE, Expression"
RouteGroupDisconnect = "TYPE, Expression"
MulticonnectMode = Num
OperationOrder = Num
ConnectionLifetime = Num
WaitForDebounce = Bool
PassAct = Str
FailAct = Str
PassActTarget = "TYPE, Expression"
FailActTarget = "TYPE, Expression"
CustExpr = "TYPE, Expression"
CustTrueAct = Str
CustFalseAct = Str
CustTrueActTarget = "TYPE, Expression"
CustFalseActTarget = "TYPE, Expression"
LoopType = Str
LoopWhile = "TYPE, Expression"
LoopStatus = "TYPE, Expression"
LoopIncrement = "TYPE, Expression"
LoopInitialize = "TYPE, Expression"
LoopOpt = Num
PreExpr = "TYPE, Expression"
PostExpr = "TYPE, Expression"
StatusExpr = "TYPE, Expression"
CanSpecifyModule = Bool
CanEditCode = Bool
CanEditModulePrototype = Bool
PrecondIntExe = Num
%ROOT_TYPE = True

[TEInf]
%INSTFLG: Id = 1
%INSTOVRD: Id = 5046297
%INSTOVRD: Icon = 5046296
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7143448
%INSTOVRD: PreCond = 5046296
LoadOpt = "PreloadWhenExecuted"
%INSTOVRD: LoadOpt = 5046296
UnloadOpt = "UnloadWithFile"
%INSTOVRD: UnloadOpt = 5046296
Mode = "Normal"
%INSTOVRD: Mode = 5046296
WindowActivation = "None"
%INSTOVRD: WindowActivation = 5046296
%INSTOVRD: NoResult = 5046296
StepFCSeqF = True
%INSTOVRD: StepFCSeqF = 5046296
%INSTOVRD: IgnoreRTE = 5046296
%INSTOVRD: UseMutex = 5046296
%INSTOVRD: MutexNameOrRef = 5046296
%INSTOVRD: BatchSyncOpt = 5046296
%INSTOVRD: SwitchEnabled = 5046296
%INSTOVRD: VirtualDeviceName = 5046296
SwitchOperation = 1
%INSTOVRD: SwitchOperation = 5046296
%INSTOVRD: RouteGroupConnect = 5046296
%INSTOVRD: RouteGroupDisconnect = 5046296
MulticonnectMode = 1
%INSTOVRD: MulticonnectMode = 5046296
OperationOrder = 2
%INSTOVRD: OperationOrder = 5046296
ConnectionLifetime = 4
%INSTOVRD: ConnectionLifetime = 5046296
WaitForDebounce = True
%INSTOVRD: WaitForDebounce = 5046296
PassAct = "Next"
%INSTOVRD: PassAct = 5046296
FailAct = "Next"
%INSTOVRD: FailAct = 5046296
%INSTOVRD: PassActTarget = 5046296
%INSTOVRD: FailActTarget = 5046296
%INSTOVRD: CustExpr = 5046296
CustTrueAct = "Next"
%INSTOVRD: CustTrueAct = 5046296
CustFalseAct = "Next"
%INSTOVRD: CustFalseAct = 5046296
%INSTOVRD: CustTrueActTarget = 5046296
%INSTOVRD: CustFalseActTarget = 5046296
LoopType = "NoLooping"
%INSTOVRD: LoopType = 5046296
%INSTOVRD: LoopWhile = 5046296
%INSTOVRD: LoopStatus = 5046296
LoopIncrement = "RunState.LoopIndex += 1"
%INSTOVRD: LoopIncrement = 5046296
LoopInitialize = "RunState.LoopIndex = 0"
%INSTOVRD: LoopInitialize = 5046296
%INSTOVRD: LoopOpt = 5046296
%INSTOVRD: PreExpr = 5046296
%INSTOVRD: PostExpr = 5046296
%INSTOVRD: StatusExpr = 5046296
CanSpecifyModule = True
%INSTOVRD: CanSpecifyModule = 5046296
CanEditCode = True
%INSTOVRD: CanEditCode = 5046296
CanEditModulePrototype = True
%INSTOVRD: CanEditModulePrototype = 5046296
%INSTOVRD: PrecondIntExe = 5046296
%TIMESTAMP = 1059680189
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194328
%TYPE_FLG = 33554436
%INSTFLG = 4456472
%INSTOVRD = 4456472

[DEF, Error]
%LOCATION = "StdTypes"
Code = Num
Msg = Str
Occurred = Bool
%ROOT_TYPE = True

[Error]
%FLG: Code = 4194304
%INSTFLG: Code = 4194304
%FLG: Msg = 4194304
%INSTFLG: Msg = 4194304
%FLG: Occurred = 4194304
%INSTFLG: Occurred = 4194304
%TIMESTAMP = 1063730206
%VERSION = "3.0.0.12374"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194304
%TYPE_FLG = 33554432
%EXTDATA = True

[Error.Code]
%EXTDATA = True

[EXTDATA, Error.Code, STRUCT]
DataVersion = 1
Type = 6

[EXTDATA, Error.Code, CLUST]
DataVersion = 1
ClusterMemberLabelName = "code"

[EXTDATA, Error.Code, DNSTRUCT]
DataVersion = 1
StructMemberName = "code"

[Error.Msg]
%EXTDATA = True

[EXTDATA, Error.Msg, STRUCT]
DataVersion = 1
Type = 2
StringBufferSize = 1024
StringStorage = 1

[EXTDATA, Error.Msg, CLUST]
DataVersion = 1
ClusterMemberLabelName = "source"

[EXTDATA, Error.Msg, DNSTRUCT]
DataVersion = 1
StructMemberName = "msg"

[Error.Occurred]
%EXTDATA = True

[EXTDATA, Error.Occurred, STRUCT]
DataVersion = 1
Type = 6

[EXTDATA, Error.Occurred, CLUST]
DataVersion = 1
ClusterMemberLabelName = "status"

[EXTDATA, Error.Occurred, DNSTRUCT]
DataVersion = 1
StructMemberName = "occurred"

[EXTDATA, Error, STRUCT]
AllowStructPassing = True
DataVersion = 1
PackingOption = 8
Type = 0

[EXTDATA, Error, CLUST]
AllowClusterPassing = True
DataVersion = 1
ClusterMemberLabelName = ""

[EXTDATA, Error, DNSTRUCT]
AllowStructPassing = True
DataVersion = 1
StructMemberName = ""

[DEF, CommonResults]
%LOCATION = "StdTypes"
%ROOT_TYPE = True

[CommonResults]
%TIMESTAMP = 1059680189
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194304
%TYPE_FLG = 33554432
%INSTFLG = 4194304

[DEF, AutomationStepAdditions]
%LOCATION = "StdTypes"
Call = AutomationCall
%ROOT_TYPE = True

[DEF, AutomationStepAdditions.Call]
ObjectVariable = "TYPE, Expression"
Server = Str
ServerName = Str
CreateObject = Num
CreateFromFilePath = "TYPE, Path"
CoClass = Str
CoClassName = Str
Interface = Str
InterfaceName = Str
RemoteMachine = "TYPE, Expression"
RemoteMachineByExpr = Bool
UseLoadSpec = Bool
MakeCall = Bool
MemberType = Num
Member = Num
MemberName = Str
HasMemberInfo = Bool
Locale = Num
HasReturnValue = Bool
TypeLibLocale = Num
TypeLibVersion = Str
InterfaceType = Num
VTableOffset = Num
Parameters = Objs

[AutomationStepAdditions.Call]
CreateObject = 3
MakeCall = True
MemberType = 2
InterfaceType = 2

[DEF, AutomationStepAdditions.Call.Parameters]
%EPTYPE = "TYPE, AutomationParameter"
%FLG: %EPTYPE = 131072

[AutomationStepAdditions]
%TIMESTAMP = 1059680804
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, AutomationParameter]
%LOCATION = "StdTypes"
Name = Str
ArgVal = "TYPE, Expression"
ArgDisplayVal = "TYPE, Expression"
Type = Num
DisplayType = Str
TypeValid = Bool
Direction = Num
IsUserOptional = Bool
IsServerOptional = Bool
IsByRef = Bool
IsSafeArray = Bool
IID = Str
UseDefault = Bool
%ROOT_TYPE = True

[AutomationParameter]
Name = "_notNamed"
%TIMESTAMP = 1059680804
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, FlexCStepAdditions]
%LOCATION = "StdTypes"
Call = ExternalCall
ModuleSrcPath = "TYPE, Path"
ModulePrjPath = "TYPE, Path"
ModuleWorkspacePath = "TYPE, Path"
ModuleCreateSrcType = Num
%ROOT_TYPE = True

[DEF, FlexCStepAdditions.Call]
LibPath = "TYPE, Path"
Func = Str
Parms = Objs

[DEF, FlexCStepAdditions.Call.Parms]
%EPTYPE = "TYPE, FCParameter"
%FLG: %EPTYPE = 131072
%[0] = "TYPE, FCParameter"

[FlexCStepAdditions.Call]
%HI: Parms = [0]

[FlexCStepAdditions.Call.Parms[0]]
Name = "Return Value"
Type = 3

[FlexCStepAdditions]
%TIMESTAMP = 1059681025
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, FCParameter]
%LOCATION = "StdTypes"
Name = Str
Type = Num
NumType = Num
ObjType = Num
StructType = Str
NumPass = Num
StrPass = Num
ElemPass = Num
ArrayDimensionsSize = Objs
StrSize = "TYPE, Expression"
ResultAct = Num
ArgVal = "TYPE, Expression"
ArgDisplayVal = "TYPE, Expression"
Flags = Num
ArgValImag = "TYPE, Expression"
%ROOT_TYPE = True

[FCParameter]
Name = "_notNamed"
NumType = 4
StrSize = "1024"
%TIMESTAMP = 1059681025
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, FCParameter.ArrayDimensionsSize]
%EPTYPE = "TYPE, Expression"
%FLG: %EPTYPE = 131072

[FCParameter.ArrayDimensionsSize]
%EPTYPE = "1024"

[DEF, Expression]
%LOCATION = "StdTypes"
%ROOT_TYPE = True

[Expression]
%TIMESTAMP = 1060111435
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 4194304
%TYPE_FLG = 33554436

[DEF, NoneStepAdditions]
%LOCATION = "StdTypes"
%ROOT_TYPE = True

[NoneStepAdditions]
%TIMESTAMP = 1059680859
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, Label]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
CodeTemplates = Str
%FLG: CodeTemplates = 524288
Description = Str
%ROOT_TYPE = True

[Label]
%HI: Substeps = [0]
%INSTFLG: Substeps = 524312
%INSTOVRD: Substeps = 655384
DescriptionFormat = "Step.Description"
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "ResStr(\"NI_STEPTYPES\", \"LABEL_DEF_STEP_NAME\")"
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: Menu = 524312
%INSTOVRD: Menu = 524312
%INSTFLG: TS = 4456472
%INSTOVRD: TS = 4456472
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%FLG: CodeTemplates = 4194328
%TIMESTAMP = 1133254865
%VERSION = "3.0.0.12377"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%TYPE_FLG = 33554446

[DEF, Label.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, Label.Substeps[0].TS]
SData = "TYPE, FlexCStepAdditions"
%FLG: SData = 2097152

[Label.Substeps[0].TS.SData.Call]
LibPath = "CommonSubsteps.dll"
Func = "EditLabelStep"
%HI: Parms = [3]

[DEF, Label.Substeps[0].TS.SData.Call.Parms]
%[0] = "TYPE, FCParameter"
%[1] = "TYPE, FCParameter"
%[2] = "TYPE, FCParameter"
%[3] = "TYPE, FCParameter"

[Label.Substeps[0].TS.SData.Call.Parms[0]]
Name = "Return Value"
Type = 3

[Label.Substeps[0].TS.SData.Call.Parms[1]]
Name = "object"
Type = 4
ArgVal = "ThisContext"

[Label.Substeps[0].TS.SData.Call.Parms[2]]
Name = "reserved"
ArgVal = "0"

[Label.Substeps[0].TS.SData.Call.Parms[3]]
Name = "changed"
NumPass = 1
ArgVal = "RunState.InitialSelection.SelectedFile.ChangeCount"
ArgDisplayVal = "RunState.InitialSelection.SelectedFile.ChangeCount"

[Label.Substeps[0].TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 2097152
ConnectionLifetime = 0

[Label.Substeps[0]]
%INSTFLG: TS = 262168
%INSTOVRD: TS = 262168
MenuName = "ResStr(\"NI_SUBSTEPS\", \"LABEL_EDIT_STEP_MENU_NAME\")"
SupportsReadOnly = True

[Label.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0
%INSTOVRD: Common = 4194304

[DEF, Label.Substeps[0]]
%NAME = "Edit"

[Label.Menu]
ItemName = "ResStr(\"NI_STEPTYPES\", \"LABEL_MENU_ITEM_NAME\")"
Group = "NI_Miscellaneous"

[Label.TS]
Icon = "label.ico"
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7143448
%INSTFLG: PassAct = 1
%INSTOVRD: PassAct = 5046297
%INSTFLG: FailAct = 1
%INSTOVRD: FailAct = 5046297
%INSTFLG: PassActTarget = 1
%INSTOVRD: PassActTarget = 5046297
%INSTFLG: FailActTarget = 1
%INSTOVRD: FailActTarget = 5046297
CanSpecifyModule = False

[DEF, Label.TS]
SData = "TYPE, NoneStepAdditions"
%FLG: SData = 2097152

[DEF, Label.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"

[Label.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304
%INSTOVRD: Common = 4194304

[DEF, SequenceCall]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True

[SequenceCall]
%INSTFLG: Substeps = 524312
%INSTOVRD: Substeps = 655384
DescriptionFormat = "\"%ModuleDescription\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "ResStr(\"NI_STEPTYPES\", \"SC_DEF_STEP_NAME\")"
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: Menu = 524312
%INSTOVRD: Menu = 524312
%INSTFLG: TS = 4456472
%INSTOVRD: TS = 4456472
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%FLG: CodeTemplates = 4194328
%TIMESTAMP = 1063730206
%VERSION = "3.0.0.12374"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%TYPE_FLG = 33554438

[SequenceCall.Menu]
ItemName = "ResStr(\"NI_STEPTYPES\", \"SC_MENU_ITEM_NAME\")"
Group = "NI_Miscellaneous"

[SequenceCall.TS]
Icon = "SeqAdp.ico"
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7143448
ConnectionLifetime = 0
StatusExpr Line0001 = "(Step.Result.Status == \"Done\" && (Step.TS.SData.ThreadOpt == 0 || Step.TS.SData.ThreadOpt == 3)) ? \"Passed\" : Step.R"
StatusExpr Line0002 = "esult.Status"
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, SequenceCall.TS]
SData = "TYPE, SeqCallStepAdditions"
%FLG: SData = 2097152

[DEF, SequenceCall.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"

[SequenceCall.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304
%INSTOVRD: Common = 4194304

[DEF, SeqCallStepAdditions]
%LOCATION = "StdTypes"
SFPath = "TYPE, Path"
SFPathExpr = "TYPE, Expression"
SeqName = Str
SeqNameExpr = "TYPE, Expression"
ActualArgs = Arguments
%FLG: ActualArgs = 2097152
UseCurFile = Bool
SpecifyByExpr = Bool
Prototype = Obj
%FLG: Prototype = 2097152
UsePrototype = Bool
SpecifyHostByExpr = Bool
RemoteExecution = Bool
RemoteHost = Str
RemoteHostExpr = "TYPE, Expression"
AutoWaitAsync = Bool
CreateThreadSuspended = Bool
AsyncThreadExpr = "TYPE, Expression"
AsyncApartmentThreaded = Bool
ThreadOpt = Num
ExecModelOpt = Num
ExecTypeMask = Num
ExecTypeMaskExpr = "TYPE, Expression"
ExecBreakOnEntryExpr = "TYPE, Expression"
ExecModelPath = "TYPE, Path"
ExecModelPathExpr = "TYPE, Expression"
ExecSync = Bool
Trace = Str
IgnoreTerminate = Bool
%ROOT_TYPE = True

[SeqCallStepAdditions]
%INSTFLG: ActualArgs = 2097152
%INSTFLG: Prototype = 2097152
AutoWaitAsync = True
ExecModelOpt = 1
ExecBreakOnEntryExpr = "False"
Trace = "Don't Change"
%TIMESTAMP = 1059680765
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, NI_Wait]
%LOCATION = "StepTypes"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
Substeps = "TYPE, StepTypeSubstepsArray"
CodeTemplates = Str
%FLG: CodeTemplates = 524288
TS = "TYPE, TEInf"
Result = Obj
%FLG: Result = 2097152
TimeoutEnabled = Bool
ErrorOnTimeout = Bool
ThreadRefExpr = "TYPE, Expression"
SeqCallName = Str
SeqCallStepGroupIdx = Num
TimeoutExpr = "TYPE, Expression"
WaitForTarget = Num
TimeExpr = "TYPE, Expression"
ExecutionRefExpr = "TYPE, Expression"
SpecifyBySeqCall = Bool
%ROOT_TYPE = True

[NI_Wait]
DescriptionFormat Line0001 = "ResStr(\"NI_WAIT_STEP_TYPE_TARGET_NAME\", Str(Step.WaitForTarget)) +\r\"(\" +\r((Step.WaitForTarget == 0) ? LocalizeExpr"
DescriptionFormat Line0002 = "ession(Step.TimeExpr):\r (Step.WaitForTarget == 1) ? LocalizeExpression(Step.TimeExpr):\r (Step.WaitForTarget == 2) ?\r "
DescriptionFormat Line0003 = "      (Step.SpecifyBySeqCall \r         ? ((Step.SeqCallName == \"\") \r             ? ResStr(\"NI_WAIT_STEP_TYPE\", \"U"
DescriptionFormat Line0004 = "NSPECIFIED\")\r             : TargetName(RunState.Sequence, Step.SeqCallName, false, true, Step.SeqCallStepGroupIdx, Run"
DescriptionFormat Line0005 = "State.StepGroup))\r         : LocalizeExpression(Step.ThreadRefExpr)) :\r (Step.WaitForTarget == 3) ?\r       (Step.Spec"
DescriptionFormat Line0006 = "ifyBySeqCall \r         ? ((Step.SeqCallName == \"\") \r             ? ResStr(\"NI_WAIT_STEP_TYPE\", \"UNSPECIFIED\")\r "
DescriptionFormat Line0007 = "            : TargetName(RunState.Sequence, Step.SeqCallName, false, true, Step.SeqCallStepGroupIdx, RunState.StepGroup)"
DescriptionFormat Line0008 = ")\r         : LocalizeExpression(Step.ExecutionRefExpr)) :\r \"\")\r+ \")\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "ResStr(\"NI_WAIT_STEP_TYPE\", \"DEF_STEP_NAME\")"
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: Menu = 524312
%INSTOVRD: Menu = 524312
%HI: Substeps = [1]
%INSTFLG: Substeps = 524312
%INSTOVRD: Substeps = 655384
CodeTemplates = "Default_Template"
%FLG: CodeTemplates = 4194328
%INSTFLG: TS = 4456472
%INSTOVRD: TS = 4456472
%FLG: Result = 4194304
%INSTFLG: Result = 6291456
ErrorOnTimeout = True
SpecifyBySeqCall = True
%TIMESTAMP = 1032811630
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%TYPE_FLG = 33554446

[NI_Wait.Menu]
ItemName = "ResStr(\"NI_WAIT_STEP_TYPE\", \"MENU_ITEM_NAME\")"
Group = "NI_NonAdvSynchronization"

[DEF, NI_Wait.Substeps]
%[0] = Step
%TYPE: %[0] = "PostSubstep"
%[1] = Step
%TYPE: %[1] = "EditSubstep"

[DEF, NI_Wait.Substeps[0].TS]
SData = "TYPE, FlexCStepAdditions"
%FLG: SData = 2097152

[NI_Wait.Substeps[0].TS.SData.Call]
LibPath = "SyncSteps.dll"
Func = "DoWaitStep"
%HI: Parms = [1]

[DEF, NI_Wait.Substeps[0].TS.SData.Call.Parms]
%[0] = "TYPE, FCParameter"
%[1] = "TYPE, FCParameter"

[NI_Wait.Substeps[0].TS.SData.Call.Parms[0]]
Name = "Return Value"
Type = 3

[NI_Wait.Substeps[0].TS.SData.Call.Parms[1]]
Name = "seqContext"
Type = 4
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"

[NI_Wait.Substeps[0].TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 2097152
ConnectionLifetime = 0

[NI_Wait.Substeps[0]]
%INSTFLG: TS = 262168
%INSTOVRD: TS = 262168

[NI_Wait.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0
%INSTOVRD: Common = 4194304

[DEF, NI_Wait.Substeps[0]]
%NAME = "Post"

[DEF, NI_Wait.Substeps[1].TS]
SData = "TYPE, FlexCStepAdditions"
%FLG: SData = 2097152

[NI_Wait.Substeps[1].TS.SData.Call]
LibPath = "SyncSteps.dll"
Func = "EditWaitStep"
%HI: Parms = [3]

[DEF, NI_Wait.Substeps[1].TS.SData.Call.Parms]
%[0] = "TYPE, FCParameter"
%[1] = "TYPE, FCParameter"
%[2] = "TYPE, FCParameter"
%[3] = "TYPE, FCParameter"

[NI_Wait.Substeps[1].TS.SData.Call.Parms[0]]
Name = "Return Value"
Type = 3

[NI_Wait.Substeps[1].TS.SData.Call.Parms[1]]
Name = "seqContextDisp"
Type = 4
ArgVal = "ThisContext"

[NI_Wait.Substeps[1].TS.SData.Call.Parms[2]]
Name = "reserved"
ArgVal = "0"

[NI_Wait.Substeps[1].TS.SData.Call.Parms[3]]
Name = "changeCount"
NumPass = 1
ArgVal = "RunState.InitialSelection.SelectedFile.ChangeCount"
ArgDisplayVal = "RunState.InitialSelection.SelectedFile.ChangeCount"

[NI_Wait.Substeps[1].TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 2097152
ConnectionLifetime = 0

[NI_Wait.Substeps[1]]
%INSTFLG: TS = 262168
%INSTOVRD: TS = 262168
MenuName = "ResStr(\"NI_WAIT_STEP_TYPE\", \"EDIT_STEP_MENU_NAME\")"
SupportsReadOnly = True

[NI_Wait.Substeps[1].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0
%INSTOVRD: Common = 4194304

[DEF, NI_Wait.Substeps[1]]
%NAME = "Edit"

[NI_Wait.TS]
Icon = "ni_hourglass.ico"
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7143448
BatchSyncOpt = 2
ConnectionLifetime = 0

[DEF, NI_Wait.TS]
SData = "TYPE, NoneStepAdditions"
%FLG: SData = 2097152

[DEF, NI_Wait.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"

[NI_Wait.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304
%INSTOVRD: Common = 4194304

[DEF, PostSubstep]
%LOCATION = "StepTypes"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
Substeps = "TYPE, StepTypeSubstepsArray"
CodeTemplates = Str
%FLG: CodeTemplates = 524288
TS = "TYPE, TEInf"
Result = Obj
%ROOT_TYPE = True

[PostSubstep]
DescriptionFormat = "\"%ModuleDescription\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "ResStr(\"NI_STEPTYPES\", \"POST_DEF_SUBSTEP_NAME\")"
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
CodeTemplates Line0001 = "DefaultLabVIEW|DefaultCVI|DefaultVB.NET|DefaultCSharp.NET|DefaultC++.NET|DefaultVC++_Template|DefaultHTB72_Template|Defa"
CodeTemplates Line0002 = "ultHTB80_Template|Default_Template"
%FLG: CodeTemplates = 4194328
%INSTFLG: TS = 4456472
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%TIMESTAMP = 1059680189
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%TYPE_FLG = 33554438
%INSTFLG = 4194304
%INSTOVRD = 4194304

[PostSubstep.Menu]
CanBeSubstepType = True
CanOnlyBeSubstepType = True
ItemName = "ResStr(\"NI_STEPTYPES\", \"POST_SUBSTEP_MENU_ITEM_NAME\")"
Group = "ExecSubsteps"

[DEF, PostSubstep.TS]
SData = Obj
%FLG: SData = 2097152

[PostSubstep.TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7274521

[DEF, PostSubstep.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"

[PostSubstep.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304

[DEF, RP_Init]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_Init]
DescriptionFormat = "\"Initialize calls to R/P servers\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Initialize calls to R/P servers\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1287753984
%VERSION = "0.0.0.6"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[RP_Init.Menu]
ItemName = "\"Initialize calls to R/P servers\""
Group = "Start\\\\Stop Testbench"

[RP_Init.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_Init.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_Init.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3.0 - CB4.0 - RP"
CreateObject = 0
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{0DF4C4C9-8D92-4AAA-A7A3-AFFBB50D832E}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809344
MemberName = "RP_Initialize"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "4ba.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [3]

[DEF, RP_Init.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"
%[3] = "TYPE, AutomationParameter"

[RP_Init.TS.SData.Call.Parameters[0]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 1
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[RP_Init.TS.SData.Call.Parameters[1]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_Init.TS.SData.Call.Parameters[2]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[RP_Init.TS.SData.Call.Parameters[3]]
Name = "SharedMemoryName"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsUserOptional = True
IsServerOptional = True
IsByRef = True

[DEF, RP_Init.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_Init.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, RP_Close]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_Close]
DescriptionFormat = "\"Terminate R/P calls\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Terminate R/P calls\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1173370942
%VERSION = "0.0.0.4"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[RP_Close.Menu]
ItemName = "\"Terminate R/P calls\""
Group = "Start\\\\Stop Testbench"

[RP_Close.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_Close.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_Close.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{06ECF673-4EE8-4BFB-B24A-F1BA0C42CCF6}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809345
MemberName = "RP_Terminate"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "2cc.0"
InterfaceType = 1
VTableOffset = 32
%HI: Parameters = [1]

[DEF, RP_Close.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[RP_Close.TS.SData.Call.Parameters[0]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_Close.TS.SData.Call.Parameters[1]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, RP_Close.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_Close.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, RP_StartServer]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_StartServer]
%HI: Substeps = [0]
DescriptionFormat = "\"Start R/P Server\" + \" ( IP adress : \" + Step.TS.SData.Call.Parameters[0].ArgVal + \")\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Start R/P Server\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1201268994
%VERSION = "0.0.0.6"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, RP_StartServer.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, RP_StartServer.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_StartServer.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, RP_StartServer.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[RP_StartServer.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[RP_StartServer.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[RP_StartServer.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[RP_StartServer.Substeps[0]]
%INSTFLG: TS = 0

[RP_StartServer.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, RP_StartServer.Substeps[0]]
%NAME = "Edit"

[RP_StartServer.Menu]
ItemName = "\"Start R/P Server\""
Group = "Start\\\\Stop Testbench"

[RP_StartServer.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_StartServer.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_StartServer.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{06ECF673-4EE8-4BFB-B24A-F1BA0C42CCF6}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809346
MemberName = "RP_StartServer"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "2cc.0"
InterfaceType = 1
VTableOffset = 36
%HI: Parameters = [2]

[DEF, RP_StartServer.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"

[RP_StartServer.TS.SData.Call.Parameters[0]]
Name = "strIPadress"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_StartServer.TS.SData.Call.Parameters[1]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_StartServer.TS.SData.Call.Parameters[2]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, RP_StartServer.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_StartServer.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, RP_StopServer]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_StopServer]
%HI: Substeps = [0]
DescriptionFormat = "\"Stop R/P Server\" + \" ( IP adress : \" + Step.TS.SData.Call.Parameters[0].ArgVal + \")\""
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Stop R/P Server\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1201268994
%VERSION = "0.0.0.7"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, RP_StopServer.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, RP_StopServer.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_StopServer.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, RP_StopServer.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[RP_StopServer.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[RP_StopServer.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[RP_StopServer.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[RP_StopServer.Substeps[0]]
%INSTFLG: TS = 0

[RP_StopServer.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, RP_StopServer.Substeps[0]]
%NAME = "Edit"

[RP_StopServer.Menu]
ItemName = "\"Stop R/P Server\""
Group = "Start\\\\Stop Testbench"

[RP_StopServer.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_StopServer.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_StopServer.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{06ECF673-4EE8-4BFB-B24A-F1BA0C42CCF6}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809347
MemberName = "RP_StopServer"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "2cc.0"
InterfaceType = 1
VTableOffset = 40
%HI: Parameters = [2]

[DEF, RP_StopServer.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"

[RP_StopServer.TS.SData.Call.Parameters[0]]
Name = "strIPadress"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_StopServer.TS.SData.Call.Parameters[1]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_StopServer.TS.SData.Call.Parameters[2]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, RP_StopServer.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_StopServer.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, RP_ReadConfig]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_ReadConfig]
%HI: Substeps = [0]
DescriptionFormat Line0001 = "\"Configure a R/P server \" + Step.TS.SData.Call.Parameters[0].ArgVal + \" -> \" + Step.TS.SData.Call.Parameters[1].ArgV"
DescriptionFormat Line0002 = "al "
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Configure a target R/P server\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1201268994
%VERSION = "0.0.0.6"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, RP_ReadConfig.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, RP_ReadConfig.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_ReadConfig.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, RP_ReadConfig.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[RP_ReadConfig.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[RP_ReadConfig.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[RP_ReadConfig.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[RP_ReadConfig.Substeps[0]]
%INSTFLG: TS = 0

[RP_ReadConfig.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, RP_ReadConfig.Substeps[0]]
%NAME = "Edit"

[RP_ReadConfig.Menu]
ItemName = "\"Configure a R/P server\""
Group = "Testbench configuration"

[RP_ReadConfig.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_ReadConfig.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_ReadConfig.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{06ECF673-4EE8-4BFB-B24A-F1BA0C42CCF6}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809348
MemberName = "RP_ReadConfig"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "2cc.0"
InterfaceType = 1
VTableOffset = 44
%HI: Parameters = [3]

[DEF, RP_ReadConfig.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"
%[3] = "TYPE, AutomationParameter"

[RP_ReadConfig.TS.SData.Call.Parameters[0]]
Name = "strIPadress"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_ReadConfig.TS.SData.Call.Parameters[1]]
Name = "strXMLpathname"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_ReadConfig.TS.SData.Call.Parameters[2]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_ReadConfig.TS.SData.Call.Parameters[3]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, RP_ReadConfig.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_ReadConfig.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, RP_ResetConfig]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_ResetConfig]
%HI: Substeps = [0]
DescriptionFormat = "\"Reset R/P server configuration\"+ Step.TS.SData.Call.Parameters[0].ArgVal"
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Reset R/P server configuration\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1201268994
%VERSION = "0.0.0.5"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, RP_ResetConfig.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, RP_ResetConfig.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_ResetConfig.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, RP_ResetConfig.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[RP_ResetConfig.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[RP_ResetConfig.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[RP_ResetConfig.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[RP_ResetConfig.Substeps[0]]
%INSTFLG: TS = 0

[RP_ResetConfig.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, RP_ResetConfig.Substeps[0]]
%NAME = "Edit"

[RP_ResetConfig.Menu]
ItemName = "\"Reset R/P server configuration\""
Group = "Testbench configuration"

[RP_ResetConfig.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_ResetConfig.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_ResetConfig.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{06ECF673-4EE8-4BFB-B24A-F1BA0C42CCF6}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809349
MemberName = "RP_ResetConfig"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "2cc.0"
InterfaceType = 1
VTableOffset = 48
%HI: Parameters = [2]

[DEF, RP_ResetConfig.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"

[RP_ResetConfig.TS.SData.Call.Parameters[0]]
Name = "strIPadress"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_ResetConfig.TS.SData.Call.Parameters[1]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_ResetConfig.TS.SData.Call.Parameters[2]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, RP_ResetConfig.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_ResetConfig.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, RP_StartAllEqt]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
InBuf = Str
DataSource = Str
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True
%ALWAYS_SAVE = True

[RP_StartAllEqt]
%HI: Substeps = [0]
DescriptionFormat Line0001 = "\"Start all the equipments on network \" + Step.TS.SData.Call.Parameters[1].ArgVal + \" from R/P server \" + Step.TS.SDa"
DescriptionFormat Line0002 = "ta.Call.Parameters[0].ArgVal"
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Start all equipments\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: TS = 4194304
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
DataSource = "Step.Result.PassFail"
CodeTemplates = "PassFail_Template|PassFailMFC_Template"
%FLG: CodeTemplates = 24
%COMMENT = "\n  "
%TIMESTAMP = 1201268994
%VERSION = "0.0.0.7"
%TYPELASTMOD = "3.0.0.12373"
%TYPE_FLG = 2

[DEF, RP_StartAllEqt.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, RP_StartAllEqt.Substeps[0].TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_StartAllEqt.Substeps[0].TS.SData.Call]
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CreateObject = 0
CoClass = "{3153F69B-D9BA-4C88-84B5-C6663E5D63E3}"
CoClassName = "affichage"
Interface = "{C6CD374F-AA1B-482C-849A-0BB8E8C9B397}"
InterfaceName = "_affichage"
MemberType = 1
Member = 1610809344
MemberName = "edition"
HasMemberInfo = True
TypeLibVersion = "2f1.0"
InterfaceType = 1
VTableOffset = 28
%HI: Parameters = [1]

[DEF, RP_StartAllEqt.Substeps[0].TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"

[RP_StartAllEqt.Substeps[0].TS.SData.Call.Parameters[0]]
Name = "Step"
ArgVal = "ThisContext.step"
ArgDisplayVal = "ThisContext.step"
Type = 9
DisplayType = "PropertyObject (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{8D8769A0-09C0-11D1-93BD-0020AF68E893}"

[RP_StartAllEqt.Substeps[0].TS.SData.Call.Parameters[1]]
Name = "SeqContext"
ArgVal = "ThisContext"
ArgDisplayVal = "ThisContext"
Type = 9
DisplayType = "SequenceContext (Object Reference)"
TypeValid = True
Direction = 3
IsByRef = True
IID = "{6427D0E0-E410-11D0-93B4-0020AF68E893}"

[RP_StartAllEqt.Substeps[0].TS]
%INSTOVRD: SData = 7274520
ConnectionLifetime = 0

[RP_StartAllEqt.Substeps[0]]
%INSTFLG: TS = 0

[RP_StartAllEqt.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0

[DEF, RP_StartAllEqt.Substeps[0]]
%NAME = "Edit"

[RP_StartAllEqt.Menu]
ItemName = "\"Start all equipments\""
Group = "Equipments"

[RP_StartAllEqt.TS]
Icon = "ni_ThreadPriority.ico"
ConnectionLifetime = 0
StatusExpr Line0001 = "Step.DataSource != \"Step.Result.PassFail\" ? Step.Result.PassFail = Evaluate(Step.DataSource) : False, Step.Result.Pass"
StatusExpr Line0002 = "Fail ? \"Passed\" : \"Failed\" "
%INSTFLG: StatusExpr = 1
%INSTOVRD: StatusExpr = 5046297

[DEF, RP_StartAllEqt.TS]
SData = "TYPE, AutomationStepAdditions"
%FLG: SData = 2097152

[RP_StartAllEqt.TS.SData.Call]
ObjectVariable = "FileGlobals.Remote"
Server = "{1E52CADB-5F9E-4E14-992F-C317D7B79AE2}"
ServerName = "Interface between TestStand3 et CB3.0.1."
CoClass = "{640B1D1E-8DBE-4F30-8988-8E985EF95F9D}"
CoClassName = "RemoteCBTarget"
Interface = "{1035A14C-A552-459B-95B9-CF76AD1AB41A}"
InterfaceName = "_RemoteCBTarget"
MemberType = 1
Member = 1610809364
MemberName = "RP_StartAllEqt"
HasMemberInfo = True
Locale = 1036
TypeLibVersion = "305.0"
InterfaceType = 1
VTableOffset = 108
%HI: Parameters = [3]

[DEF, RP_StartAllEqt.TS.SData.Call.Parameters]
%[0] = "TYPE, AutomationParameter"
%[1] = "TYPE, AutomationParameter"
%[2] = "TYPE, AutomationParameter"
%[3] = "TYPE, AutomationParameter"

[RP_StartAllEqt.TS.SData.Call.Parameters[0]]
Name = "strIPadress"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_StartAllEqt.TS.SData.Call.Parameters[1]]
Name = "strNetwork"
ArgVal = "\"\""
ArgDisplayVal = "\"\""
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 1

[RP_StartAllEqt.TS.SData.Call.Parameters[2]]
Name = "passFail"
ArgVal = "Step.Result.PassFail"
ArgDisplayVal = "Step.Result.PassFail"
Type = 11
DisplayType = "Boolean"
TypeValid = True
Direction = 3
IsByRef = True

[RP_StartAllEqt.TS.SData.Call.Parameters[3]]
Name = "errorMsg"
ArgVal = "Step.Result.ReportText"
ArgDisplayVal = "Step.Result.ReportText"
Type = 8
DisplayType = "String"
TypeValid = True
Direction = 3
IsByRef = True

[DEF, RP_StartAllEqt.Result]
Error = "TYPE, Error"
Status = Str
PassFail = Bool
ReportText = Str
Common = "TYPE, CommonResults"

[RP_StartAllEqt.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
PassFail = True
%INSTOVRD: Common = 4194304

[DEF, DotNetStepAdditions]
%LOCATION = "StdTypes"
Call = DotNetCall
ModuleSrcPath = "TYPE, Path"
ModulePrjPath = "TYPE, Path"
ModuleSlnPath = "TYPE, Path"
FunctionName = Str
%ROOT_TYPE = True

[DEF, DotNetStepAdditions.Call]
AssemblyPath = "TYPE, Path"
AssemblyStrongName = Str
AssemblyLocation = Num
ClassVariable = "TYPE, Expression"
ClassName = Str
IsStruct = Bool
CreateInstance = Bool
DisposeCreatedInstance = Bool
Constructor = Str
RemoteSpecifyExpr = Bool
RemoteURI = Str
UseLoadSpec = Bool
CallMember = Bool
MemberType = Num
Static = Bool
TopLevel = Bool
MemberName = Str
HasReturnValue = Bool
Params = Objs
ConstructorParams = Objs
StructDef = "TYPE, DotNetParameter"

[DotNetStepAdditions.Call]
DisposeCreatedInstance = True
CallMember = True
TopLevel = True

[DEF, DotNetStepAdditions.Call.Params]
%EPTYPE = "TYPE, DotNetParameter"
%FLG: %EPTYPE = 131072

[DEF, DotNetStepAdditions.Call.ConstructorParams]
%EPTYPE = "TYPE, DotNetParameter"
%FLG: %EPTYPE = 131072

[DotNetStepAdditions]
%TIMESTAMP = 1059697789
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, DotNetParameter]
%LOCATION = "StdTypes"
Name = Str
ArgVal = "TYPE, Expression"
ArgDisplayVal = "TYPE, Expression"
Type = Num
DisplayType = Str
Direction = Num
IsOptional = Bool
CallDispose = Bool
IsArray = Bool
Dimension = Num
MultiElement = Objs
%ROOT_TYPE = True

[DotNetParameter]
Name = "_notNamed"
Dimension = 1
%TIMESTAMP = 1059697789
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%FLG = 24
%TYPE_FLG = 33554436

[DEF, Variable_Force]
%LOCATION = "StepTypes"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
BlockStartTypes = Str
%FLG: BlockStartTypes = 524288
BlockEndTypes = Str
%FLG: BlockEndTypes = 524288
AppliesToBlockStructure = Bool
%FLG: AppliesToBlockStructure = 524288
Menu = "TYPE, StepTypeMenu"
Substeps = "TYPE, StepTypeSubstepsArray"
CodeTemplates = Str
%FLG: CodeTemplates = 524288
TS = "TYPE, TEInf"
Result = Obj
VariableName = Str
VariableValue = Num
%ROOT_TYPE = True

[Variable_Force]
DescriptionFormat Line0001 = "(Step.VariableName ==\"\") ? \"Forces variable\" : \"Forces variable \"+Str(Step.VariableName)+((Step.VariableValue !=0)"
DescriptionFormat Line0002 = " ? \" : sets its value to \"+ Str(Step.VariableValue) : \"\")"
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Force Variable\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%FLG: BlockStartTypes = 4194328
%FLG: BlockEndTypes = 4194328
%FLG: AppliesToBlockStructure = 4194328
%HI: Substeps = [0]
CodeTemplates Line0001 = "DefaultLabVIEW|DefaultCVI|DefaultVB.NET|DefaultCSharp.NET|DefaultC++.NET|DefaultVC++_Template|DefaultHTB72_Template|Defa"
CodeTemplates Line0002 = "ultHTB80_Template|Default_Template"
%FLG: CodeTemplates = 4194328
%INSTFLG: TS = 262168
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%TIMESTAMP = 1266242554
%VERSION = "0.0.0.12"
%TYPELASTMOD = "4.0.0.1030"
%TYPE_FLG = 8

[Variable_Force.Menu]
ItemName = "\"Force variable\""
Group = "Variables"

[DEF, Variable_Force.Substeps]
%[0] = Step
%TYPE: %[0] = "PostSubstep"

[Variable_Force.Substeps[0].TS]
Id = "ID#:L06vGxZM10K3NTY1SnrPbA"

[DEF, Variable_Force.Substeps[0].TS]
SData = "TYPE, DotNetStepAdditions"
%FLG: SData = 2097152

[Variable_Force.Substeps[0].TS.SData.Call]
AssemblyPath = "C:\\\\Program Files\\\\Alstom\\\\UTM4TS\\\\bin\\\\TestStand_TrainTracer.dll"
ClassVariable = "FileGlobals.WebMaintenance"
ClassName = "Alstom.eTrain.TestStand.Maintenance"
MemberType = 1
MemberName = "GetMessage"
%HI: Params = [3]

[DEF, Variable_Force.Substeps[0].TS.SData.Call.Params]
%[0] = "TYPE, DotNetParameter"
%[1] = "TYPE, DotNetParameter"
%[2] = "TYPE, DotNetParameter"
%[3] = "TYPE, DotNetParameter"

[Variable_Force.Substeps[0].TS.SData.Call.Params[0]]
Name = "errorOccurred"
ArgVal = "Step.Result.Error.Occurred"
Type = 2
Direction = 10
Dimension = -1

[Variable_Force.Substeps[0].TS.SData.Call.Params[1]]
Name = "errorCode"
ArgVal = "Step.Result.Error.Code"
Type = 6
Direction = 10
Dimension = -1

[Variable_Force.Substeps[0].TS.SData.Call.Params[2]]
Name = "errorMsg"
ArgVal = "Step.Result.Error.Msg"
Type = 1
Direction = 10
Dimension = -1

[Variable_Force.Substeps[0].TS.SData.Call.Params[3]]
Name = "reportText"
ArgVal = "Step.Result.ReportText"
Type = 1
Direction = 10
Dimension = -1

[DEF, Variable_Force.Substeps[0]]
%NAME = "Post"

[Variable_Force.TS]
Icon = "TrainTracer\\\\Force.ico"
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7274521
UnloadOpt = "UnloadAfterSequenceExecution"

[DEF, Variable_Force.TS]
SData = "TYPE, DotNetStepAdditions"
%FLG: SData = 2097152

[Variable_Force.TS.SData.Call]
AssemblyPath = "C:\\\\Program Files\\\\Alstom\\\\UTM4TS\\\\bin\\\\TestStand_TrainTracer.dll"
ClassVariable = "FileGlobals.WebMaintenance"
ClassName = "Alstom.eTrain.TestStand.Maintenance"
DisposeCreatedInstance = False
Constructor = "Maintenance()"
MemberType = 1
MemberName = "ForceVariable"
HasReturnValue = True
%HI: Params = [2]

[DEF, Variable_Force.TS.SData.Call.Params]
%[0] = "TYPE, DotNetParameter"
%[1] = "TYPE, DotNetParameter"
%[2] = "TYPE, DotNetParameter"

[Variable_Force.TS.SData.Call.Params[0]]
Name = "Return Value"
ArgVal = "Step.Result.APICall"
Type = 2
Direction = 6
Dimension = -1

[Variable_Force.TS.SData.Call.Params[1]]
Name = "variableSymbolicPath"
ArgVal = "Step.VariableName"
Type = 1
Direction = 1
Dimension = -1

[Variable_Force.TS.SData.Call.Params[2]]
Name = "value"
ArgVal = "Step.VariableValue"
Type = 12
Direction = 1
Dimension = -1

[DEF, Variable_Force.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"
APICall = Bool

[Variable_Force.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304

[DEF, Variable_Read_Test]
%LOCATION = "StepTypes"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
BlockStartTypes = Str
%FLG: BlockStartTypes = 524288
BlockEndTypes = Str
%FLG: BlockEndTypes = 524288
AppliesToBlockStructure = Bool
%FLG: AppliesToBlockStructure = 524288
Menu = "TYPE, StepTypeMenu"
Substeps = "TYPE, StepTypeSubstepsArray"
CodeTemplates = Str
%FLG: CodeTemplates = 524288
TS = "TYPE, TEInf"
Result = Obj
VariableName = Str
Limits = Obj
Comp = Str
CompExpr = "TYPE, Expression"
UseCompExpr = Bool
InBuf = Str
DataSource = "TYPE, Expression"
%ROOT_TYPE = True

[Variable_Read_Test]
DescriptionFormat = "(Step.VariableName ==\"\") ? \"Reads variable\" : \"Reads variable : \"+Step.VariableName"
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "\"Pass/Fail Read Variable\""
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%FLG: BlockStartTypes = 4194328
%FLG: BlockEndTypes = 4194328
%FLG: AppliesToBlockStructure = 4194328
%HI: Substeps = [1]
CodeTemplates Line0001 = "DefaultLabVIEW|DefaultCVI|DefaultVB.NET|DefaultCSharp.NET|DefaultC++.NET|DefaultVC++_Template|DefaultHTB72_Template|Defa"
CodeTemplates Line0002 = "ultHTB80_Template|Default_Template|NumericLimitLabVIEW|NumericLimitCVI|NumericLimitVB.NET|NumericLimitCSharp.NET|Numeric"
CodeTemplates Line0003 = "LimitC++.NET|NumericLimitVC++_Template|NumericLimitHTB80_Template|NumericLimitHTB72Template|NumericLimit_Template"
%FLG: CodeTemplates = 4194328
%INSTFLG: TS = 262168
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%FLG: Limits = 12288
%INSTFLG: Limits = 12288
Comp = "GELE"
%FLG: Comp = 12288
%INSTFLG: Comp = 12288
%FLG: CompExpr = 6144
DataSource = "Step.Result.Numeric"
%TIMESTAMP = 1266242453
%VERSION = "0.0.0.19"
%TYPELASTMOD = "4.0.0.1030"
%TYPE_FLG = 8

[Variable_Read_Test.Menu]
ItemName = "\"Pass/Fail Read variable\""
Group = "Variables"

[DEF, Variable_Read_Test.Substeps]
%[0] = Step
%TYPE: %[0] = "PostSubstep"
%[1] = Step
%TYPE: %[1] = "EditSubstep"

[Variable_Read_Test.Substeps[0].TS]
Id = "ID#:L06vGxZM10K3NTY1SnrPbA"

[DEF, Variable_Read_Test.Substeps[0].TS]
SData = "TYPE, DotNetStepAdditions"
%FLG: SData = 2097152

[Variable_Read_Test.Substeps[0].TS.SData.Call]
AssemblyPath = "C:\\\\Program Files\\\\Alstom\\\\UTM4TS\\\\bin\\\\TestStand_TrainTracer.dll"
ClassVariable = "FileGlobals.WebMaintenance"
ClassName = "Alstom.eTrain.TestStand.Maintenance"
MemberType = 1
MemberName = "GetMessage"
%HI: Params = [3]

[DEF, Variable_Read_Test.Substeps[0].TS.SData.Call.Params]
%[0] = "TYPE, DotNetParameter"
%[1] = "TYPE, DotNetParameter"
%[2] = "TYPE, DotNetParameter"
%[3] = "TYPE, DotNetParameter"

[Variable_Read_Test.Substeps[0].TS.SData.Call.Params[0]]
Name = "errorOccurred"
ArgVal = "Step.Result.Error.Occurred"
Type = 2
Direction = 10
Dimension = -1

[Variable_Read_Test.Substeps[0].TS.SData.Call.Params[1]]
Name = "errorCode"
ArgVal = "Step.Result.Error.Code"
Type = 6
Direction = 10
Dimension = -1

[Variable_Read_Test.Substeps[0].TS.SData.Call.Params[2]]
Name = "errorMsg"
ArgVal = "Step.Result.Error.Msg"
Type = 1
Direction = 10
Dimension = -1

[Variable_Read_Test.Substeps[0].TS.SData.Call.Params[3]]
Name = "reportText"
ArgVal = "Step.Result.ReportText"
Type = 1
Direction = 10
Dimension = -1

[DEF, Variable_Read_Test.Substeps[0]]
%NAME = "Post"

[DEF, Variable_Read_Test.Substeps[1].TS]
SData = "TYPE, FlexCStepAdditions"
%FLG: SData = 2097152

[Variable_Read_Test.Substeps[1].TS.SData.Call]
LibPath = "CommonSubsteps.dll"
Func = "EditNumericMeasurementStepUsingExpr"
%HI: Parms = [3]

[DEF, Variable_Read_Test.Substeps[1].TS.SData.Call.Parms]
%[0] = "TYPE, FCParameter"
%[1] = "TYPE, FCParameter"
%[2] = "TYPE, FCParameter"
%[3] = "TYPE, FCParameter"

[Variable_Read_Test.Substeps[1].TS.SData.Call.Parms[0]]
Name = "Return Value"
Type = 3
ObjType = 2

[Variable_Read_Test.Substeps[1].TS.SData.Call.Parms[1]]
Name = "sequenceContext"
Type = 4
ArgVal = "ThisContext"

[Variable_Read_Test.Substeps[1].TS.SData.Call.Parms[2]]
Name = "reserved"
ArgVal = "0"

[Variable_Read_Test.Substeps[1].TS.SData.Call.Parms[3]]
Name = "changed"
ObjType = 2
NumPass = 1
ArgVal = "RunState.InitialSelection.SelectedFile.ChangeCount"
ArgDisplayVal = "RunState.InitialSelection.SelectedFile.ChangeCount"

[Variable_Read_Test.Substeps[1].TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 2097152
ConnectionLifetime = 0

[Variable_Read_Test.Substeps[1]]
%INSTFLG: TS = 262168
%INSTOVRD: TS = 262168
MenuName = "ResStr(\"NI_STEPTYPES\", \"NUM_EDIT_STEP_MENU_NAME\")"
SupportsReadOnly = True

[Variable_Read_Test.Substeps[1].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0
%INSTOVRD: Common = 4194304

[DEF, Variable_Read_Test.Substeps[1]]
%NAME = "Edit"

[Variable_Read_Test.TS]
Icon = "TrainTracer\\\\Read.ico"
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7274521
UnloadOpt = "UnloadAfterSequenceExecution"
StatusExpr Line0001 = "CheckLimits(Step.DataSource == \"Step.Result.Numeric\" ? Step.Result.Numeric : (Step.Result.Numeric = Evaluate(Step.Data"
StatusExpr Line0002 = "Source)), Step.Limits.High, Step.Limits.Low,Step.Comp)"

[DEF, Variable_Read_Test.TS]
SData = "TYPE, DotNetStepAdditions"
%FLG: SData = 2097152

[Variable_Read_Test.TS.SData.Call]
AssemblyPath = "C:\\\\Program Files\\\\Alstom\\\\UTM4TS\\\\bin\\\\TestStand_TrainTracer.dll"
ClassVariable = "FileGlobals.WebMaintenance"
ClassName = "Alstom.eTrain.TestStand.Maintenance"
DisposeCreatedInstance = False
Constructor = "Maintenance()"
MemberType = 1
MemberName = "ReadVariable"
HasReturnValue = True
%HI: Params = [2]

[DEF, Variable_Read_Test.TS.SData.Call.Params]
%[0] = "TYPE, DotNetParameter"
%[1] = "TYPE, DotNetParameter"
%[2] = "TYPE, DotNetParameter"

[Variable_Read_Test.TS.SData.Call.Params[0]]
Name = "Return Value"
ArgVal = "Step.Result.APICall"
Type = 2
Direction = 6
Dimension = -1

[Variable_Read_Test.TS.SData.Call.Params[1]]
Name = "variableSymbolicPath"
ArgVal = "Step.VariableName"
Type = 1
Direction = 1
Dimension = -1

[Variable_Read_Test.TS.SData.Call.Params[2]]
Name = "value"
ArgVal = "Step.Result.Numeric"
Type = 12
Direction = 10
Dimension = -1

[DEF, Variable_Read_Test.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"
APICall = Bool
Numeric = Num
Units = Str

[Variable_Read_Test.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304
%FLG: Numeric = 9216
%INSTFLG: Numeric = 9216
%FLG: Units = 10240
%INSTFLG: Units = 8192

[DEF, Variable_Read_Test.Limits]
Low = Num
High = Num
LowExpr = "TYPE, Expression"
HighExpr = "TYPE, Expression"
UseLowExpr = Bool
UseHighExpr = Bool

[Variable_Read_Test.Limits]
Low = 9
High = 11
%FLG: LowExpr = 2048
%FLG: HighExpr = 2048
%FLG: UseLowExpr = 2048
%FLG: UseHighExpr = 2048

[DEF, Statement]
%LOCATION = "StepTypes"
Substeps = "TYPE, StepTypeSubstepsArray"
DescriptionFormat = "TYPE, Expression"
%FLG: DescriptionFormat = 524288
DefaultNameFormat = "TYPE, Expression"
%FLG: DefaultNameFormat = 524288
Menu = "TYPE, StepTypeMenu"
TS = "TYPE, TEInf"
Result = Obj
CodeTemplates = Str
%FLG: CodeTemplates = 524288
%ROOT_TYPE = True

[Statement]
%HI: Substeps = [0]
%INSTFLG: Substeps = 524312
%INSTOVRD: Substeps = 655384
DescriptionFormat = "LocalizeExpression(Step.TS.PostExpr)"
%FLG: DescriptionFormat = 4194328
%INSTFLG: DescriptionFormat = 4718616
%INSTOVRD: DescriptionFormat = 4718616
DefaultNameFormat = "ResStr(\"NI_SUBSTEPS\", \"STATEMENT_DEF_STEP_NAME\")"
%FLG: DefaultNameFormat = 4194328
%INSTFLG: DefaultNameFormat = 4718616
%INSTOVRD: DefaultNameFormat = 4718616
%INSTFLG: Menu = 524312
%INSTOVRD: Menu = 524312
%INSTFLG: TS = 4456472
%INSTOVRD: TS = 4456472
%FLG: Result = 4194304
%INSTFLG: Result = 4194304
%FLG: CodeTemplates = 4194328
%TIMESTAMP = 1035380509
%VERSION = "3.0.0.12373"
%TYPELASTMOD = "3.0.0.12373"
%MINPRODVER = "3.0.0.0"
%TYPE_FLG = 33554446

[DEF, Statement.Substeps]
%[0] = Step
%TYPE: %[0] = "EditSubstep"

[DEF, Statement.Substeps[0].TS]
SData = "TYPE, FlexCStepAdditions"
%FLG: SData = 2097152

[Statement.Substeps[0].TS.SData.Call]
LibPath = "CommonSubsteps.dll"
Func = "EditStatementStep"
%HI: Parms = [3]

[DEF, Statement.Substeps[0].TS.SData.Call.Parms]
%[0] = "TYPE, FCParameter"
%[1] = "TYPE, FCParameter"
%[2] = "TYPE, FCParameter"
%[3] = "TYPE, FCParameter"

[Statement.Substeps[0].TS.SData.Call.Parms[0]]
Name = "Return Value"
Type = 3
ObjType = 2

[Statement.Substeps[0].TS.SData.Call.Parms[1]]
Name = "context"
Type = 4
ArgVal = "ThisContext"

[Statement.Substeps[0].TS.SData.Call.Parms[2]]
Name = "reserved"
ArgVal = "0"

[Statement.Substeps[0].TS.SData.Call.Parms[3]]
Name = "changed"
ObjType = 2
NumPass = 1
ArgVal = "RunState.InitialSelection.SelectedFile.ChangeCount"

[Statement.Substeps[0].TS]
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 2097152
ConnectionLifetime = 0

[Statement.Substeps[0]]
%INSTFLG: TS = 262168
%INSTOVRD: TS = 262168
MenuName = "ResStr(\"NI_SUBSTEPS\", \"STATEMENT_EDIT_STEP_MENU_NAME\")"
SupportsReadOnly = True

[Statement.Substeps[0].Result]
%INSTOVRD: Error = 4194304
%INSTFLG: ReportText = 0
%INSTOVRD: Common = 4194304

[DEF, Statement.Substeps[0]]
%NAME = "Edit"

[Statement.Menu]
ItemName = "ResStr(\"NI_SUBSTEPS\", \"STATEMENT_MENU_ITEM_NAME\")"
Group = "NI_Miscellaneous"

[Statement.TS]
Icon = "Statement.ico"
%INSTFLG: SData = 2097152
%INSTOVRD: SData = 7143448
ConnectionLifetime = 0
CanSpecifyModule = False

[DEF, Statement.TS]
SData = "TYPE, NoneStepAdditions"
%FLG: SData = 2097152

[DEF, Statement.Result]
Error = "TYPE, Error"
Status = Str
ReportText = Str
Common = "TYPE, CommonResults"

[Statement.Result]
%FLG: Error = 4194304
%INSTFLG: Error = 4194304
%INSTOVRD: Error = 4194304
%FLG: Status = 4194304
%INSTFLG: Status = 4194304
%FLG: ReportText = 4194304
%INSTFLG: ReportText = 4194304
%INSTOVRD: Common = 4194304
>>