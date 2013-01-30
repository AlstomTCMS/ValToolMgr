VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_zApplication_Settings 
   Caption         =   "UserForm1"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   OleObjectBlob   =   "uf_zApplication_Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_zApplication_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'''CONSTANTS
Const FORM_CAPTION As String = "Application Settings"





''''UI CODE
''''
''''

'''SET UP CODE ETC
Private Sub UserForm_Activate()
    ''Set Caption
    Me.Caption = FORM_CAPTION
    ''Set up list box up
    Me.lbHeadings.ColumnCount = 2
    Me.lbData.ColumnCount = 2
    Me.lbHeadings.ColumnWidths = "108;170"
    Me.lbData.ColumnWidths = "108; 170"
    Me.lbHeadings.AddItem "Headers"
    Me.lbHeadings.List(0, 0) = "Property"
    Me.lbHeadings.List(0, 1) = "Value"
    Me.lbHeadings.Enabled = False
    ''Load current values
    LoadCurrentValues
End Sub

Private Sub UserForm_Initialize()
    '''Set file path location
 FilePath = ThisWorkbook.Path
End Sub

'''Loads the current values into dispalybox - lbData
Private Sub LoadCurrentValues()
    Me.lbData.Clear
    If FileExists(FilePath & "\" & SETTING_FILE_NAME) = True Then
        SortToArray (FilePath & "\" & SETTING_FILE_NAME)
        Me.lbData.List = SortedSettingsArray
    End If
End Sub

'''Updates Values in from edit boxes
Private Sub cmdSave_Click()
    Dim x As Boolean
    x = UpdateValue(Me.txtName.Text, Me.txtValue.Text)
    LoadCurrentValues
End Sub

'''Load Data in to edit boxes when dbl clicked
Private Sub lbData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long
    For i = 0 To Me.lbData.ListCount - 1
        If Me.lbData.Selected(i) = True Then
            '''update value user slected
            Me.txtName = Me.lbData.List(i, 0)
            Me.txtValue = Me.lbData.List(i, 1)
        End If
    Next
End Sub



''''TESTING/DEBUG CODES
''''
''''
Private Sub CommandButton1_Click()
    CommandButton1.Caption = GetValue(Me.TextBox1.Text)
End Sub

Private Sub CommandButton2_Click()
    Dim x As Boolean
    x = UpdateValue(Me.TextBox1.Text, Me.TextBox2.Text)
    LoadCurrentValues
End Sub

Private Sub CommandButton3_Click()
    Dim x As Boolean
    x = AddValue(Me.TextBox3.Value, Me.TextBox2.Value)
    LoadCurrentValues
End Sub


