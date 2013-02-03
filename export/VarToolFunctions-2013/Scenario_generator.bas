Attribute VB_Name = "Scenario_generator"
Option Explicit

Private Enum tableType
    TABLE_ACTIONS
    TABLE_CHECKS
End Enum

Public Sub Generate_scenario(ByVal testNumber As String)
'Public Sub Generate_scenario()
'    Dim testNumber As String
'    testNumber = "1.2"

    Dim wsCurrentTestSheet As Worksheet, _
        wsResultSheet As Worksheet, _
        loActionsTable As ListObject, _
        loChecksTable As ListObject

        
    Dim scenario_shName As String
 
    'optimisation excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.CutCopyMode = False
    Application.DisplayAlerts = False

    Set wsCurrentTestSheet = ActiveSheet 'Worksheets("Scenario")
    'wsCurrentTestSheet.Activate

    Set loActionsTable = wsCurrentTestSheet.ListObjects(PR_TEST_TABLE_ACTION_PREFIX & testNumber)
    Set loChecksTable = wsCurrentTestSheet.ListObjects(PR_TEST_TABLE_CHECK_PREFIX & testNumber)
    
      ' Pour l'instant on ne vérifie pas le formalisme
    'If Not checkingTestFormat(lcActionsTableColumns) Then GoTo fin

    scenario_shName = PR_TEST_SCENARIO_PREFIX & testNumber

    Dim o_test As CTest
    Set o_test = parseSingleTest(scenario_shName, loActionsTable, loChecksTable)
    
    Dim o_testContainer As CTestContainer
    Set o_testContainer = New CTestContainer
    
    o_testContainer.AddTest o_test
    
    Dim genTs As GeneratorTs401
    Set genTs = New GeneratorTs401
    
    Call genTs.writeScenario("C:\\macros_alstom\\test\\testGen.seq", o_testContainer)
    
    
End_GenScenario:
    'optimisation excel
    Debug.Print "End of scenario"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Function parseSingleTest(title As String, loActionsTable As ListObject, loChecksTable As ListObject) As CTest
    Dim lcActionsTableColumns As ListColumns, _
        lcChecksTableColumns As ListColumns
    
    Set parseSingleTest = New CTest
    parseSingleTest.title = title
    Set lcActionsTableColumns = loActionsTable.ListColumns
    Set lcChecksTableColumns = loChecksTable.ListColumns
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Writing inputs
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim CurrentLine As Integer, CurrentColumn As Integer
    CurrentLine = 1
    Dim OffsetSection As Integer
    OffsetSection = 3
    
    For CurrentColumn = 3 To lcActionsTableColumns.Count
        ' Writing header
        Debug.Print "Processing Step : " & lcActionsTableColumns.Item(CurrentColumn)
        Dim o_step As CStep
        Set o_step = New CStep
        o_step.title = lcActionsTableColumns.Item(CurrentColumn).name
        o_step.DescAction = "TBD" ' getComment(wsCurrentTestSheet, loActionsTable, CurrentColumn, "TBD")
        o_step.DescCheck = "TBD" ' getComment(wsCurrentTestSheet, loChecksTable, CurrentColumn, "Verifications to perform")
    
        
        Call fillWithActions(o_step, TABLE_ACTIONS, loActionsTable, CurrentColumn)
        
        Call addTempoIfExists(o_step, loActionsTable, CurrentColumn)
        
        Call fillWithActions(o_step, TABLE_CHECKS, loChecksTable, CurrentColumn)
        
        parseSingleTest.AddStep o_step
    Next CurrentColumn
End Function

Private Sub fillWithActions(o_step As CStep, typeOfTable As tableType, loSourceFiles As ListObject, ColumnIndex As Integer)
    Dim line As Integer
    
    For line = 1 To loSourceFiles.ListRows.Count
        Dim lrCurrent As ListRow, _
            Target As Variant, _
            Location As Variant, _
            CellValue As Variant
            
        Set lrCurrent = loSourceFiles.ListRows(line)
        Target = lrCurrent.Range(1, 1).value
        Location = lrCurrent.Range(1, 2).value
        CellValue = lrCurrent.Range(1, ColumnIndex)
        

        If Not IsEmpty(CellValue) Then
            Dim o_instruction As CInstruction
            Set o_instruction = detectAndBuildInstruction(Target, Location, CellValue, typeOfTable)
            o_step.AddInstruction o_instruction
        End If
    Next line
End Sub

Private Function detectAndBuildInstruction(Target As Variant, Location As Variant, CellValue As Variant, typeOfTable As tableType) As CInstruction
    Set detectAndBuildInstruction = New CInstruction
    
    detectAndBuildInstruction.category = UNIMPLEMENTED
    detectAndBuildInstruction.Data = Null
    
    Dim o_variable As CVariable
    Set o_variable = buildVariable(Target, Location, CellValue)
    
    If (typeOfTable = TABLE_ACTIONS) Then
        Debug.Print "Detection of an ACTION type"
        
        If (CellValue Like "U") Then
            detectAndBuildInstruction.category = A_UNFORCE
        Else
            detectAndBuildInstruction.category = A_FORCE
        End If
        Set detectAndBuildInstruction.Data = o_variable
    ElseIf (typeOfTable = TABLE_CHECKS) Then
        Debug.Print "Detection of an CHECK type"
        
        detectAndBuildInstruction.category = A_TEST
        Set detectAndBuildInstruction.Data = o_variable
    End If
End Function

Private Function buildVariable(Target As Variant, Location As Variant, CellValue As Variant) As CVariable
    
    Set buildVariable = New CVariable
    buildVariable.name = Target
    buildVariable.path = Location
    buildVariable.value = CellValue
    Dim offset
    If (InStr(1, Target, "I:", 1) = 1) Then
        buildVariable.typeOfVar = T_INTEGER
        buildVariable.name = Mid(buildVariable.name, 3)
    ElseIf (InStr(1, Target, "DT:", 1) = 1) Then
        buildVariable.typeOfVar = T_DATE_AND_TIME
         buildVariable.name = Mid(buildVariable.name, 4)
    Else
        buildVariable.typeOfVar = T_BOOLEAN
    End If

End Function

Sub addTempoIfExists(o_step As CStep, loSourceFiles As ListObject, ColumnIndex As Integer)
    'Delay retrieval. We know that data is contained inside Total line property
    Dim delay As String
    delay = loSourceFiles.TotalsRowRange.Cells(1, ColumnIndex)
    If delay <> "" Then
        Dim o_tempo As CInstruction
        Set o_tempo = New CInstruction
        o_tempo.category = A_WAIT
        o_tempo.Data = delay
        o_step.AddInstruction o_tempo
    End If
End Sub

Function getComment(wsCurrentTestSheet As Worksheet, lcTable As ListObject, CurrentColumn As Integer, OldComment As String) As String
    Dim ColumnsHeaderPosition As Integer
    
    getComment = OldComment
    
    xPosition = lcTable.HeaderRowRange.Row - 1
    yPosition = lcTable.ListColumns.Item(CurrentColumn).Range.Column
    
    If xPosition > 0 And Not IsEmpty(wsCurrentTestSheet.Cells(xPosition, yPosition)) Then
        getComment = wsCurrentTestSheet.Cells(xPosition, yPosition).value
    End If
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checking that columns are the same between two tables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function checkingTestFormat(lcActionsTableColumns As ListColumns) As Boolean

    If lcActionsTableColumns.Count <> lcChecksTableColumns.Count Then
        Debug.Print "Wrong column count"
        Exit Function
    End If
    
    If lcActionsTableColumns.Item(1) <> "Variable" Then
        Debug.Print "First column must be called Variable"
        Exit Function
    End If
    
    If lcActionsTableColumns.Item(2) <> "Type" Then
        Debug.Print "Second column must be called Type"
        Exit Function
    End If
    
    If lcActionsTableColumns.Item(3) <> "Localisation" Then
        Debug.Print "Second column must be called Localisation"
        Exit Function
    End If
    
    If lcActionsTableColumns.Item(4) <> "Section" Then
        Debug.Print "Second column must be called Section"
        Exit Function
    End If
    
    For i = 1 To lcActionsTableColumns.Count
        If lcActionsTableColumns.Item(i) <> lcChecksTableColumns.Item(i) Then
            Debug.Print "Columns has not same names : " & lcActionsTableColumns.Item(i) & " / " & lcChecksTableColumns.Item(i)
            Exit Function
        End If
    Next i
End Function




