Attribute VB_Name = "Template_events_handler"
Sub testAddRef()
    Call Add_AddInRef_to_WorkBook(ActiveWorkbook)
End Sub

Sub Add_AddInRef_to_WorkBook(wb As Workbook)
    On Error Resume Next
    Application.DisplayAlerts = False
    With wb.VBProject.References
        .AddFromFile MacroPath & "\Functions_PrimaELII_2-A0.xlam"
    End With
    Application.DisplayAlerts = True
    On Error GoTo 0
 End Sub

Public Sub Template_Workbook_Open(wb As Workbook)
    Dim oList As ListObject
    Dim sh As Worksheet
    Dim nm_r As Name, nm_c As Name
    Dim strName_r As String, strName_c As String
    
    'Refresh linked data sources
    wb.RefreshAll
    
    Application.EnableEvents = True
    
End Sub

Public Sub Endpaper_PR_Sheet_change(wb As Workbook, ByVal Target As range)
    Dim functionName As String
    Dim targetValue As String
    Dim nm As Name
    
    'update FunctionName from user choice
    On Error Resume Next
    Set nm = wb.Names("FunctionIndex")
    If nm Is Nothing Then
        Set nm = wb.Sheets(ENDPAPER_PR_NAME).Names("FunctionIndex")
    End If
    On Error GoTo 0
    If Not nm Is Nothing Then
        If Not Application.Intersect(Target, nm.RefersToRange) Is Nothing Then
            targetValue = Target.Cells(1, 1).value
            If targetValue = "" Then 'in order to handle range deletion by user
                functionName = "" '=Endpaper_PR!$D$9"
            Else
                If InStr(targetValue, ":") <> 0 Then
                    functionName = Left(targetValue, InStr(targetValue, ":") - 2)
                Else
                    'TODO: find the whole function description
                    functionName = targetValue
                End If
            End If
            Call UpdateName(wb, "FunctionName", functionName) 'ActiveWorkbook.Names("Functions_2ES5").RefersToRange.Range("A" & 1)
            
        End If
    End If
End Sub


Public Sub activate_event_handling()
    Application.EnableEvents = True
End Sub

Private Sub UpdateName(wb As Workbook, strName As String, ByVal value As String)
    Dim nm As Name
    
    Set nm = Nothing
    On Error Resume Next
    Set nm = wb.Names(strName)
    On Error GoTo 0
    If nm Is Nothing Then
        'If Not value = -1 Then
            Set nm = wb.Names.Add(strName, value)
            'nm.Visible = False
        'End If
    Else
        If Not value = Empty Then
            nm.RefersTo = value
        Else
            nm.RefersTo = " "
        End If
    End If
    
End Sub

