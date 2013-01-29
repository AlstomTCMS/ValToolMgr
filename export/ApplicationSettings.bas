Attribute VB_Name = "ApplicationSettings"


Public SortedSettingsArray As Variant  '''Array that holds the seeting data in memory
            '''var that is used to hold ref to this workbook path,
                                       '''i.e the setting file will be in the same dir ass the workboo this form runs in
                                       '''Change as needed - set in for Activate code


''' ###############################
''' NB: ALL THINGS ARE STRINGS!
''' V: 0.0.1
''' 05 June 2012
''' This is attempt 1 at a genaral purpose settings control for Excel VBA,
''' writen by MIE Ross Mclea, rossmclean@gmail.com... while ill in hospital - hence its a bit crappy!
''' ###############################


'''The 3 expposed functions,
'    .AddValue - addes a new key-value pair to and existing setting file, or creates the file if its not there first time #
'         - file path and name are this workbook.path, and a CONST ans set above
'         - can not add 2 keys the same
'         - returns T/F
'    .UpdateValue, updates X value for Y key, if key is there
'           - returns T/F
'    .GetVallue, get a vlaue from Key Y, if key Y exists
'           - returns value, or "False" as string


'''Adds new property to settings list
Public Function AddValue(ByRef PropertyName As String, ByRef PropertyValue As String) As Boolean
    If FileExists(MacroPath & "\" & SETTING_FILE_NAME) = False Then
        ''iff the seeting file has not been made, make one!
        MakeFile MacroPath & "\" & SETTING_FILE_NAME
        ''' if its the first one we do it directly! - save fannying around with other code
        Open MacroPath & "\" & SETTING_FILE_NAME For Append As #1
        Write #1, PropertyName & "|" & PropertyValue
        Close #1    ' Close file.
    End If
    If ValueExists(PropertyName) = False Then
        'Add Values
        Open MacroPath & "\" & SETTING_FILE_NAME For Append As #1
        Write #1, PropertyName & "|" & PropertyValue
        Close #1    ' Close file.
        AddValue = True
    Else
        AddValue = False
    End If
End Function

Public Function UpdateValue(ByRef PropertyName As String, ByRef PropertyValue As String) As Boolean
'''Not an eleegent soltion - a clooection would have been better... but we are where we are - i was in hospital when i wrote this!
    If FileExists(MacroPath & "\" & SETTING_FILE_NAME) = True Then
        SortToArray (MacroPath & "\" & SETTING_FILE_NAME)
        'lookin array for same key value
        Dim i As Long
        For i = 0 To UBound(SortedSettingsArray)
            If SortedSettingsArray(i, 0) = PropertyName Then

                ''Update Value into array
                SortedSettingsArray(i, 1) = PropertyValue

                'DropFile
                Open MacroPath & "\" & SETTING_FILE_NAME For Output As #1
                Close #1

                'Repoulate with array - hope for no crashes!
                Dim k As Long
                Open MacroPath & "\" & SETTING_FILE_NAME For Append As #1
                For k = 0 To UBound(SortedSettingsArray)
                    Write #1, SortedSettingsArray(k, 0) & "|" & SortedSettingsArray(k, 1)
                Next
                Close #1    ' Close file

                UpdateValue = True
                Exit Function
            End If
        Next
    Else
        UpdateValue = "False"
    End If
    '''Loop array to file
End Function


'''Gets value of seting by name
Public Function GetValue(ByRef PropertyName As String) As String
    If FileExists(MacroPath & "\" & SETTING_FILE_NAME) = True Then
        SortToArray (MacroPath & "\" & SETTING_FILE_NAME)

        'lookin array for same key value
        Dim i As Long
        For i = 0 To UBound(SortedSettingsArray)
            If SortedSettingsArray(i, 0) = PropertyName Then
                GetValue = SortedSettingsArray(i, 1)
                Exit Function
            End If
        Next
    Else
        GetValue = "Error"
    End If
End Function

Private Function ValueExists(ByRef PropertyName As String) As Boolean
'Update Array
    If FileExists(MacroPath & "\" & SETTING_FILE_NAME) = True Then
        SortToArray (MacroPath & "\" & SETTING_FILE_NAME)
        'lookin array for same key value
        Dim i As Long
        For i = 0 To UBound(SortedSettingsArray)
            If SortedSettingsArray(i, 0) = PropertyName Then
                ValueExists = True
                Exit Function
            End If
        Next
    Else
        ValueExists = False
    End If
End Function



'''Dose text file already exist?
Private Function FileExists(textFile) As Boolean
' If Dir() returns something, the file exists.
' On Error Resume Next
    FileExists = (Dir(textFile) <> "")
    If FileExists = True Then
        Exit Function
    Else
        FileExist = False
    End If
End Function

'''Make text File
Private Function MakeFile(ByRef textFile As String) As Boolean
    If FileExists(textFile) = False Then
        Open textFile For Append As #1
        Close #1    ' Close file.
    End If
End Function

'''Sorts Text File in to Array
Private Function SortToArray(ByRef textFile As String) As Boolean
    Dim intFileNum As Integer
    Dim intCount As Integer
    Dim strRecordData As String
    Dim TempArray As Variant
    Dim FileLines As Long

    Open textFile For Input As #1
    fileinfo = Input(LOF(1), #1)
    Close #1
    TempArray = Split(fileinfo, vbCrLf)
    FileLines = UBound(TempArray)

    '''a setting file is there but has nothing in it...
    If FileLines <= -1 Then
        Exit Function
    End If
    'Required public defined array
    ReDim SortedSettingsArray(FileLines - 1, 1)

    intFileNum = FreeFile
    intCount = 0
    Open textFile For Input As #intFileNum
    Do Until EOF(intFileNum)
        Input #intFileNum, strRecordData
        TempArray = Split(strRecordData, "|")
        SortedSettingsArray(intCount, 0) = TempArray(0)
        SortedSettingsArray(intCount, 1) = TempArray(1)
        intCount = intCount + 1
    Loop
    Close #intFileNum
    
    ''' Sort Array
    BubbleSort SortedSettingsArray, 0
End Function

'''Bubble Sort sub, sort data in array
'''Thanks Andy Pope!
Private Sub BubbleSort(TempArray As Variant, SortIndex As Long)
    Dim blnNoSwaps As Boolean
    Dim lngItem As Long
    Dim vntTemp(0 To 1) As Variant
    Dim lngCol As Long
    Do
        blnNoSwaps = True
        For lngItem = LBound(TempArray) To UBound(TempArray) - 1
            If TempArray(lngItem, SortIndex) > TempArray(lngItem + 1, SortIndex) Then
                blnNoSwaps = False
                For lngCol = 0 To 1
                    vntTemp(lngCol) = TempArray(lngItem, lngCol)
                    TempArray(lngItem, lngCol) = TempArray(lngItem + 1, lngCol)
                    TempArray(lngItem + 1, lngCol) = vntTemp(lngCol)
                Next
            End If
        Next
    Loop While Not blnNoSwaps
End Sub
