Attribute VB_Name = "NewPR_fromScratch"
Sub NewPR()
    Application.ScreenUpdating = False
    'MsgBox ERROR_NOT_IMPLEMENTED_FUNCTION
    
    'Demander à l'utilisateur le nom qu'il veut mettre
    'fileSaveFullName = Application.GetSaveAsFilename(InitialFileName:="B2_XXX_Y_A0", _
    'fileFilter:="xls Files (*.xls), *.xls")
    DefaultValue = "1."
    testName = InputBox(Prompt:="Please, give a name to your test.", _
          title:="Test Name", Default:=DefaultValue)
    
    'Créer l'ensemble des éléments du format
    If testName <> "" And testName <> DefaultValue Then
        'TODO: Tester si le test existe deja...
        Call createWholeTestFormat(testName)
    End If
    
    'Sauvegarder
    
    Application.ScreenUpdating = True
    
End Sub

Sub defige()
    Application.ScreenUpdating = True
End Sub

' Créé l'ensemble des éléments du format de test 2013
Sub createWholeTestFormat(ByVal testName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(PR_TEST_PREFIX & testName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'Ajout TEMPORAIRE d'un workbook s'il n'en n'existe pas
    If Not HasActiveBook(False) Then
        Workbooks.Add
    End If
    
    InitSheet (PR_TEST_PREFIX & testName)
    With Sheets(PR_TEST_PREFIX & testName).Tab
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    
    Call AddTableAction(testName)
    Call AddTableCheck(testName)
    Call AddTestTitle(testName)
    Call AddActionLabel(testName)
    Call AddCheckLabel(testName)
    Call AddTableDescription(testName)
End Sub

'Ajoute la table de description en haut
Sub AddTableDescription(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName)
    
        'on insert une ligne supplémentaire pour les titres (qu'il n'y a pas)
        .Rows("1:1").Insert Shift:=xlDown
        tableName = PR_TEST_TABLE_DESCRIPTION_PREFIX & testName
        .ListObjects.Add(xlSrcRange, .Range("C1:D5"), , xlYes).name = tableName
        Call AddDescTableFormat
        .ListObjects(tableName).TableStyle = PR_TEST_DESCRIPTION_TABLE_STYLE & " " & PR_TEST_DESCRIPTION_TABLE_STYLE_VERSION
        .ListObjects(tableName).ShowHeaders = False
        .ListObjects(tableName).ShowTableStyleFirstColumn = True
        .ListObjects(tableName).ShowTableStyleColumnStripes = True

        'On réefface cette ligne qui ne sert plus
        .Rows("1:1").Delete Shift:=xlUp
        
        'Ajoute les labels des titres verticaux
        .Range("C1:C3") = Application.Transpose(Array(PR_TEST_ACTION, PR_TEST_CHECK, "Name"))
        
        ' Efface la mise en forme de la première case de la ligne des totaux
        With .Range("C4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        With .Range(tableName & "[[#All],[Colonne1]]")
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With
    End With
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
        
        tableAddress = .ListObjects(TABLE_PREFIX & label & "_" & testName).Range.Address
        tableAddressArray = Split(tableAddress, "$")
        tableAddress = "A" & tableAddressArray(2) & "A" & tableAddressArray(4)
        Set LabelRange = .Range(tableAddress)
        With LabelRange
            .MergeCells = True
            .value = label
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            
            With .Font
                .name = "Calibri"
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
        .value = Replace(PR_TEST_PREFIX, "_", " ") & testName
        'TODO: Donner un nom
        With .Font
            .name = "Calibri"
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

Sub AddTableCheck(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName)
        tableName = PR_TEST_TABLE_CHECK_PREFIX & testName
        .ListObjects.Add(xlSrcRange, .Range("B8"), , xlYes).name = tableName
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

Sub AddTableAction(ByVal testName As String)
    With Sheets(PR_TEST_PREFIX & testName)
        
        tableName = PR_TEST_TABLE_ACTION_PREFIX & testName
        .ListObjects.Add(xlSrcRange, .Range("$B$5"), , xlYes).name = tableName
        .ListObjects(tableName).TableStyle = "TableStyleMedium9"
        .Range("B5:D5") = Array("Target", "Location", PR_TEST_STEP_PATERN)
        .ListObjects(tableName).ShowTotals = True
        
        With .Range(tableName & "[[#Totals],[Target]]")
            .FormulaR1C1 = "DELAY"
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
' est egale à vrai s'il faut mettre à jour le style du tableau
Function AddDescTableFormat() As Boolean
Dim FirstColumnEdges As Variant
Dim Style As TableStyle
Dim updateTable As Boolean
Dim addTable As Boolean

    addTable = True
    With ActiveWorkbook
    
        ' Vérifie que le style existe déjà et à quelle version
        For Each Style In .TableStyles
            If Style.name Like PR_TEST_DESCRIPTION_TABLE_STYLE & "*" Then
                addTable = False
                tableName = Split(Style.name, " ")
                If UBound(tableName) = 2 Then ' Si on a la version
                    'If tableName(2) < PR_TEST_DESCRIPTION_TABLE_STYLE_VERSION Then
                        updateTable = True
                    'End If
                Else
                    updateTable = True
                End If
            End If
        Next
        
        
        If addTable Then
            .TableStyles.Add (PR_TEST_DESCRIPTION_TABLE_STYLE & " " & PR_TEST_DESCRIPTION_TABLE_STYLE_VERSION)
        End If
        
        If addTable Or updateTable Then
        With .TableStyles(PR_TEST_DESCRIPTION_TABLE_STYLE & " " & PR_TEST_DESCRIPTION_TABLE_STYLE_VERSION)
            .ShowAsAvailablePivotTableStyle = False
            .ShowAsAvailableTableStyle = True
            .ShowAsAvailableSlicerStyle = False
            
            
            If updateTable Then
                AddDescTableFormat = True
                'effacer les styles avant de les définir
                For Each stylesElement In .TableStyleElements
                    stylesElement.Clear
                Next stylesElement
            End If
            
            ' -------------------------------------------------------------
            ' LA Première colonne
            ' -------------------------------------------------------------
            With .TableStyleElements(xlFirstColumn)
                With .Font
                    .Bold = True: .TintAndShade = 0:  .ThemeColor = xlThemeColorDark1
                End With
                With .Interior
                    .Color = 12419407: .TintAndShade = 0
                End With
                FirstColumnEdges = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlInsideHorizontal)
                For Each edge In FirstColumnEdges
                    With .Borders(edge)
                          .LineStyle = xlNone: .Weight = xlThick: .ThemeColor = xlThemeColorDark1: .TintAndShade = 0:
                    End With
                Next edge
            End With
            
            
            
            ' -------------------------------------------------------------
            ' Colonnes impaires
            ' -------------------------------------------------------------
            
            With .TableStyleElements(xlColumnStripe1)
                With .Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = 0
                    .Color = 15853276
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
            

            
            ' -------------------------------------------------------------
            ' Lignes impaires
            ' -------------------------------------------------------------
            
            With .TableStyleElements(xlRowStripe1).Borders(xlEdgeTop)
                .Weight = xlThin
                .LineStyle = xlNone
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314 '0.599963377788629
            End With
            
            ' -------------------------------------------------------------
            ' Lignes paires
            ' -------------------------------------------------------------
            
            
            With .TableStyleElements(xlRowStripe2).Borders(xlEdgeTop)
                .Weight = xlThin
                .LineStyle = xlNone
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314 '0.599963377788629
            End With
            
            ' -------------------------------------------------------------
            ' Ligne des Totaux
            ' -------------------------------------------------------------
            With .TableStyleElements(xlTotalRow)
                With .Borders(xlEdgeTop)
                    .ThemeColor = xlThemeColorLight2
                    .TintAndShade = 0.799951170384838 '0.799981688894314
                    .Weight = xlThin
                    .LineStyle = 0 'xlNone
                End With
                With .Font
                    .TintAndShade = 0
                    .ThemeColor = xlThemeColorDark1
                End With
            End With
            
        End With
        End If
    End With
End Function

