Attribute VB_Name = "NewPR_fromScratch"
Sub NewPR()
    'MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
    
    'Demander à l'utilisateur le nom qu'il veut mettre
    'fileSaveFullName = Application.GetSaveAsFilename(InitialFileName:="B2_XXX_Y_A0", _
    'fileFilter:="xls Files (*.xls), *.xls")
    
    'Créer l'ensemble des éléments du format
    Call createWholeTestFormat("1.3")
    
    'Sauvegarder
    
End Sub

' Créé l'ensemble des éléments du format de test 2013
Sub createWholeTestFormat(ByVal testName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(PR_TEST_PREFIX & testName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    InitSheet (PR_TEST_PREFIX & testName)
    With Sheets(PR_TEST_PREFIX & testName).Tab
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    
    Call TableAction(testName)
    Call TableCheck(testName)
    Call AddTestTitle(testName)
    'Call AddDescTableFormat
    Call AddActionLabel(testName)
    Call AddCheckLabel(testName)
End Sub

Sub AddCheckLabel(ByVal testName As String)
    Call DefineVerticalLabel(testName, PR_TEST_CHECK)
End Sub

Sub AddActionLabel(ByVal testName As String)
    Call DefineVerticalLabel(testName, PR_TEST_ACTION)
End Sub

Sub DefineVerticalLabel(ByVal testName As String, ByVal label As String)
    
    With Sheets(PR_TEST_PREFIX & testName)
        .Columns("A:A").ColumnWidth = 5.5
        
        tableAddress = .ListObjects("Table" & label & testName).Range.Address
        tableAddressArray = Split(tableAddress, "$")
        tableAddress = "A" & tableAddressArray(2) & "A" & tableAddressArray(4)
        Set LabelRange = .Range(tableAddress)
        With LabelRange
            .MergeCells = True
            .Value = label
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            
            With .Font
                .Name = "Calibri"
                .Size = 14
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
                .Bold = True
            End With
        End With
        
    End With
End Sub

Sub AddTestTitle(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName).Range("B3")
        .Value = PR_TEST_PREFIX & testName
        'TODO: Donner un nom
        With .Font
            .Name = "Calibri"
            .Size = 14
            .Bold = True
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    With Sheets(PR_TEST_PREFIX & testName)
        .Columns("B:B").ColumnWidth = 25
        .Rows("3:3").RowHeight = 30
    End With

End Sub

Sub TableCheck(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName)
        tableName = "Table" & PR_TEST_CHECK & testName
        .ListObjects.Add(xlSrcRange, .Range("B8"), , xlYes).Name = tableName
        .ListObjects(tableName).TableStyle = "TableStyleMedium12"
        .Range("B8:D8") = Array("Target", "Location", PR_TEST_STEP_PATERN)
        
        With .Range(tableName & "[[#Headers],[" & PR_TEST_STEP_PATERN & "]]")
            .AddIndent = True
            .IndentLevel = 1
        End With
        .ListObjects(tableName).ShowHeaders = False
        
        'Coloration de la colonne des variables
        With .Range("B9")
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Bold = True
            End With
        End With
    End With
End Sub

Sub TableAction(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName)
        
        tableName = "Table" & PR_TEST_ACTION & testName
        .ListObjects.Add(xlSrcRange, .Range("$B$5"), , xlYes).Name = tableName
        .ListObjects(tableName).TableStyle = "TableStyleMedium9"
        .Range("B5:D5") = Array("Target", "Location", PR_TEST_STEP_PATERN)
        .ListObjects(tableName).ShowTotals = True
        
        With .Range(tableName & "[[#Totals],[Target]]")
            .FormulaR1C1 = "TEMPO"
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        .Range("D7") = ""
        With .Range(tableName & "[[#Headers],[" & PR_TEST_STEP_PATERN & "]]")
            .AddIndent = True
            .IndentLevel = 1
        End With
        
        'Coloration de la colonne des variables
        With .Range("B6")
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Bold = True
            End With
        End With
    End With
End Sub

' Ajoute au workbook le style de tableau pour la partie descriptive s'il n'existe pas déjà
Sub AddDescTableFormat()

    With ActiveWorkbook
        On Error GoTo Add:
        BuiltIn = .TableStyles("desc table").BuiltIn
        GoTo NoAdd
        
Add:
        .TableStyles.Add ("desc table")
        With .TableStyles("desc table")
            .ShowAsAvailablePivotTableStyle = False
            .ShowAsAvailableTableStyle = True
            .ShowAsAvailableSlicerStyle = False
            
            With .TableStyleElements(xlTotalRow).Borders(xlEdgeTop)
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314
                .Weight = xlThin
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlFirstColumn).Font
                .FontStyle = "Gras"
                .TintAndShade = 0
                .ThemeColor = xlThemeColorAccent1
            End With
            With .TableStyleElements(xlFirstColumn).Borders(xlEdgeTop)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Weight = xlThick
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlFirstColumn).Borders(xlEdgeBottom)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Weight = xlThick
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlFirstColumn).Borders(xlEdgeLeft)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Weight = xlThick
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlFirstColumn).Borders(xlInsideHorizontal)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .Weight = xlThick
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlRowStripe1).Borders(xlEdgeTop)
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.599963377788629
                .Weight = xlThin
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlRowStripe2).Borders(xlEdgeTop)
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.599963377788629
                .Weight = xlThin
                .LineStyle = xlNone
            End With
            With .TableStyleElements(xlColumnStripe2).Interior
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314
            End With
        End With
NoAdd:
    End With
End Sub

