Attribute VB_Name = "Scenario_generator"
Public Sub Generate_scenario()
    Dim wsCurrentSheet As Worksheet, _
        wsResultSheet As Worksheet, _
        loInputsTable As ListObject, _
        loOutputsTable As ListObject, _
        lcInputsTableColumns As ListColumns, _
        lcOutputsTableColumns As ListColumns





    Set wsCurrentSheet = Worksheets("Scenario")
    wsCurrentSheet.Activate
    
    
    Set wsResultSheet = Worksheets("Result")
    
    Set loInputsTable = wsCurrentSheet.ListObjects("Inputs_table")
    Set lcInputsTableColumns = loInputsTable.ListColumns
    
    Set loOutputsTable = wsCurrentSheet.ListObjects("Outputs_table")
    Set lcOutputsTableColumns = loOutputsTable.ListColumns

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Checking that columns are the same between two tables
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If lcInputsTableColumns.Count <> lcOutputsTableColumns.Count Then
        Debug.Print "Wrong column count"
        Exit Sub
    End If
    
    If lcInputsTableColumns.Item(1) <> "Variable" Then
        Debug.Print "First column must be called Variable"
        Exit Sub
    End If
    
    If lcInputsTableColumns.Item(2) <> "Type" Then
        Debug.Print "Second column must be called Type"
        Exit Sub
    End If
    
    If lcInputsTableColumns.Item(3) <> "Localisation" Then
        Debug.Print "Second column must be called Localisation"
        Exit Sub
    End If
    
    If lcInputsTableColumns.Item(4) <> "Section" Then
        Debug.Print "Second column must be called Section"
        Exit Sub
    End If
    
    For i = 1 To lcInputsTableColumns.Count
        If lcInputsTableColumns.Item(i) <> lcOutputsTableColumns.Item(i) Then
            Debug.Print "Columns has not same names : " & lcInputsTableColumns.Item(i) & " / " & lcOutputsTableColumns.Item(i)
            Exit Sub
        End If
    Next i
    
    'optimisation excel
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.DisplayAlerts = False
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Removing everything
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    wsResultSheet.Cells.Clear
    
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Writing inputs
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim CurrentLine As Integer, CurrentColumn As Integer
    CurrentLine = 1
    Dim OffsetSection As Integer
    OffsetSection = 3
    
    For CurrentColumn = 5 To lcInputsTableColumns.Count
        ' Writing header
        Debug.Print "Processing Step : " & lcInputsTableColumns.Item(CurrentColumn)
        
       
        
        With wsResultSheet
            .Cells(CurrentLine, OffsetSection + 1).Value = lcInputsTableColumns.Item(CurrentColumn).Name
            .Range(.Cells(CurrentLine, OffsetSection + 3), .Cells(CurrentLine, OffsetSection + 7)).Merge
            .Range(.Cells(CurrentLine, OffsetSection + 8), .Cells(CurrentLine, OffsetSection + 14)).Merge
            .Cells(CurrentLine, OffsetSection + 3).Value = getComment(wsCurrentSheet, loInputsTable, CurrentColumn, "TBD")
            .Cells(CurrentLine, OffsetSection + 8).Value = getComment(wsCurrentSheet, loOutputsTable, CurrentColumn, "Verifications to perform")
            .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Interior.ColorIndex = 37
            .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Characters.Font.ColorIndex = 2
        End With
        
        CurrentLine = CurrentLine + 1
        ScenarioOffsetInputs = fillInputs(OffsetSection, "Force", CurrentLine, wsResultSheet, loInputsTable, CurrentColumn)
        
        ScenarioOffsetOutputs = fillInputs(OffsetSection + 5, "Test", CurrentLine, wsResultSheet, loOutputsTable, CurrentColumn)
        
        If ScenarioOffsetInputs > ScenarioOffsetOutputs Then
            CurrentLine = CurrentLine + ScenarioOffsetInputs
        Else
            CurrentLine = CurrentLine + ScenarioOffsetOutputs
        End If

    Next CurrentColumn
    
        With wsResultSheet
            .Cells(CurrentLine, OffsetSection + 1).Value = "END"
            .Range(.Cells(CurrentLine, OffsetSection + 3), .Cells(CurrentLine, OffsetSection + 7)).Merge
            .Range(.Cells(CurrentLine, OffsetSection + 8), .Cells(CurrentLine, OffsetSection + 14)).Merge
            .Cells(CurrentLine, OffsetSection + 3).Value = ""
            .Cells(CurrentLine, OffsetSection + 8).Value = ""
            .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Interior.ColorIndex = 37
            .Range(.Cells(CurrentLine, OffsetSection + 1), .Cells(CurrentLine, OffsetSection + 14)).Characters.Font.ColorIndex = 2
        End With
    
    Debug.Print "End of scenario"

End Sub

Function fillInputs(OffsetSection As Integer, Instruction As String, CurrentLine As Integer, wsResultSheet As Worksheet, loSourceFiles As ListObject, ColumnIndex As Integer) As Integer
    Dim lrCurrent As ListRow, _
        strFileNameDate As String, _
        ValueOfCell As Variant
        
        fillInputs = 0
    
    For i = 1 To loSourceFiles.ListRows.Count
        Set lrCurrent = loSourceFiles.ListRows(i)

        ValueOfCell = lrCurrent.Range(1, ColumnIndex)
        Debug.Print lrCurrent.Range(1, ColumnIndex).Address & " : " & IsEmpty(ValueOfCell) & " ( " & VarType(ValueOfCell) & ")"
        
        If Not IsEmpty(ValueOfCell) Then
        
            If VarType(ValueOfCell) >= vbInteger And VarType(ValueOfCell) <= vbDouble And lrCurrent.Range(1, 2) = "BOOL" Then
                If ValueOfCell = 0 Then
                    ValueOfCell = "'False"
                Else
                    ValueOfCell = "'True"
                End If
            End If
            
            With wsResultSheet
                .Cells(CurrentLine + fillInputs, OffsetSection + 3).Value = Instruction
                .Cells(CurrentLine + fillInputs, OffsetSection + 4).Value = lrCurrent.Range(1, 1).Value
                .Cells(CurrentLine + fillInputs, OffsetSection + 5).Value = lrCurrent.Range(1, 3).Value
                .Cells(CurrentLine + fillInputs, OffsetSection + 6).Value = ValueOfCell
                .Cells(CurrentLine + fillInputs, OffsetSection + 7).Value = lrCurrent.Range(1, 4).Value
            End With
        
            fillInputs = fillInputs + 1
        End If
    Next i
    
    'optimisation excel
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Function

Function getComment(wsCurrentSheet As Worksheet, lcTable As ListObject, CurrentColumn As Integer, OldComment As String) As String
    Dim ColumnsHeaderPosition As Integer
    
    getComment = OldComment
    
    xPosition = lcTable.HeaderRowRange.Row - 1
    yPosition = lcTable.ListColumns.Item(CurrentColumn).Range.Column
    
    If xPosition > 0 And Not IsEmpty(wsCurrentSheet.Cells(xPosition, yPosition)) Then
        getComment = wsCurrentSheet.Cells(xPosition, yPosition).Value
    End If
End Function
