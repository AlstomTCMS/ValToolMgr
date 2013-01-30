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
End Sub


Sub AddTestTitle()
    With Sheets(PR_TEST_PREFIX & testName)
        
    End With


    Sheets("Test_1.2").Select
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Test 1.2"
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Sheets("Test_1.1").Select
    Range("B3").Select
    Sheets("Test_1.2").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 14
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
    Selection.Font.Bold = True
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").ColumnWidth = 25.29
    Range("B3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
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
    Rows("3:3").RowHeight = 30.75
End Sub

Sub TableCheck(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName)
        tableName = "TableCheck" & testName
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
        
        tableName = "TableAction" & testName
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
