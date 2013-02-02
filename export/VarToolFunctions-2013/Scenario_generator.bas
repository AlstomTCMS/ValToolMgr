Attribute VB_Name = "Scenario_generator"
'Public Sub Generate_scenario(ByVal testNumber As String)
Public Sub Generate_scenario()
    Dim testNumber As String
    testNumber = "1.2"

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
    o_test = parseSingleTest(loActionsTable, loCheckTable)
    o_test.title = scenario_shName
    
End_GenScenario:
    'optimisation excel
    Debug.Print "End of scenario"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Function parseSingleTest(loActionsTable As ListObject, loChecksTable As ListObject) As CTest
    Dim lcActionsTableColumns As ListColumns, _
        lcChecksTableColumns As ListColumns
    
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
        
        With wsResultSheet
            .Cells(CurrentLine, OffsetSection + 1).value = lcActionsTableColumns.Item(CurrentColumn).name
            .Range(.Cells(CurrentLine, OffsetSection + 3), .Cells(CurrentLine, OffsetSection + 7)).Merge
            .Range(.Cells(CurrentLine, OffsetSection + 8), .Cells(CurrentLine, OffsetSection + 14)).Merge
            '.Cells(CurrentLine, OffsetSection + 3).Value = getComment(wsCurrentTestSheet, loActionsTable, CurrentColumn, "TBD")
            '.Cells(CurrentLine, OffsetSection + 8).Value = getComment(wsCurrentTestSheet, loChecksTable, CurrentColumn, "Verifications to perform")
            .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Interior.ColorIndex = 37
            .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Characters.Font.ColorIndex = 2
        End With
        
        CurrentLine = CurrentLine + 1
        ScenarioOffsetActions = fillInputs(OffsetSection, "Force", CurrentLine, wsResultSheet, loActionsTable, CurrentColumn)
        
        'ScenarioOffsetDelays = fillInputs(OffsetSection, "Wait", CurrentLine, wsResultSheet, loActionsTable, CurrentColumn)
        
        ScenarioOffsetChecks = fillInputs(OffsetSection + 5, "Test", CurrentLine, wsResultSheet, loChecksTable, CurrentColumn)
        
        If ScenarioOffsetActions > ScenarioOffsetChecks Then
            CurrentLine = CurrentLine + ScenarioOffsetActions
        Else
            CurrentLine = CurrentLine + ScenarioOffsetChecks
        End If

    Next CurrentColumn
    
    'Finalisation du test avec la ligne END
    With wsResultSheet
        .Cells(CurrentLine, OffsetSection + 1).value = "END"
        .Range(.Cells(CurrentLine, OffsetSection + 3), .Cells(CurrentLine, OffsetSection + 7)).Merge
        .Range(.Cells(CurrentLine, OffsetSection + 8), .Cells(CurrentLine, OffsetSection + 14)).Merge
        .Cells(CurrentLine, OffsetSection + 3).value = ""
        .Cells(CurrentLine, OffsetSection + 8).value = ""
        .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Interior.ColorIndex = 37
        .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Characters.Font.ColorIndex = 2
    End With
End Function

Function fillInputs(OffsetSection As Integer, Instruction As String, CurrentLine As Integer, wsResultSheet As Worksheet, loSourceFiles As ListObject, ColumnIndex As Integer) As Integer
    Dim lrCurrent As ListRow, _
    strFileNameDate As String, _
    ValueOfCell As Variant
    
    fillInputs = 0
    
    For i = 1 To loSourceFiles.ListRows.Count
        Set lrCurrent = loSourceFiles.ListRows(i)

        ValueOfCell = lrCurrent.Range(1, ColumnIndex)
        Debug.Print lrCurrent.Range(1, ColumnIndex).Address & " : " & IsEmpty(ValueOfCell) & " ( " & varType(ValueOfCell) & ")"
        
        If Not IsEmpty(ValueOfCell) Then
        
            If varType(ValueOfCell) >= vbInteger And varType(ValueOfCell) <= vbDouble And lrCurrent.Range(1, 2) = "BOOL" Then
                If ValueOfCell = 0 Then
                    ValueOfCell = "'False"
                Else
                    ValueOfCell = "'True"
                End If
            End If
            
            With wsResultSheet
                .Cells(CurrentLine + fillInputs, OffsetSection + 3).value = Instruction
                .Cells(CurrentLine + fillInputs, OffsetSection + 4).value = lrCurrent.Range(1, 1).value
                .Cells(CurrentLine + fillInputs, OffsetSection + 5).value = lrCurrent.Range(1, 3).value
                .Cells(CurrentLine + fillInputs, OffsetSection + 6).value = ValueOfCell
                .Cells(CurrentLine + fillInputs, OffsetSection + 7).value = lrCurrent.Range(1, 4).value
            End With
        
            fillInputs = fillInputs + 1
        End If
    Next i
    
    'Traitement des temporisations
    If loSourceFiles.name Like PR_TEST_TABLE_ACTION_PREFIX & "*" Then
        delay = loSourceFiles.TotalsRowRange.Cells(1, ColumnIndex)
        If delay <> "" Then
            With wsResultSheet
                .Cells(CurrentLine + fillInputs, OffsetSection + 3).value = "Wait"
                .Cells(CurrentLine + fillInputs, OffsetSection + 6).value = delay
            End With
            fillInputs = fillInputs + 1
        End If
    End If
    
End Function

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




