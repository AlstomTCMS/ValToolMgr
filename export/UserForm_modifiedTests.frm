VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_modifiedTests 
   Caption         =   "Test(s) modifi�(s)"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   OleObjectBlob   =   "UserForm_modifiedTests.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_modifiedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox_SelectAll_Click()
    If CheckBox_SelectAll Then
        For Each checkBox In Frame_TestsList.Controls
            checkBox.Value = True
        Next checkBox
        CheckBox_UnselectAll.Value = False
        CheckBox_SelectAll.Value = False
    End If
End Sub

Private Sub CheckBox_UnselectAll_Click()
'Dim CheckBox As Controls
    If CheckBox_UnselectAll Then
        For Each checkBox In Frame_TestsList.Controls
            checkBox.Value = False
        Next checkBox
        CheckBox_SelectAll.Value = False
        CheckBox_UnselectAll.Value = False
    End If
End Sub

Private Sub CommandButton_Annuler_Click()
    cancel_Synth2Tests = True
    Unload Me
End Sub

Private Sub CommandButton_Valider_Click()
    'Dim CheckBox as
    
    'R�cuperer la liste des cases coch�es
    For Each checkBox In Frame_TestsList.Controls
        If checkBox Then
            modifiedTests = modifiedTests + checkBox.Caption + ";"
        End If
    Next
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim rowtop As Integer
    Dim testsList As Variant
    
    ' faire la liste des tests
    With Sheets(SYNTHESE_NAME)
        With .range("A2:I" & .range("F1").End(xlDown).row)
            .AutoFilter field:=1, Criteria1:="<>"
            Application.CutCopyMode = False
            .Columns(1).Copy Destination:=.range("K1")
            .AutoFilter field:=1
            
            testsList = .range("K1:" & .range("K1").End(xlDown).Offset(-1, 0).Address)
            Application.DisplayAlerts = False
            .Columns("K:K").Delete
            Application.DisplayAlerts = True
        End With
    End With
    
    'Ajout des CheckBox
    Frame_TestsList.Controls.Clear
    NumeroTextBox = 1: rowtop = 0
        
    For i = 1 To UBound(testsList)
        Set checkBox = Frame_TestsList.Controls.Add("Forms.CheckBox.1")
        
        With checkBox
            .Value = False
            .Caption = testsList(i, 1)
            .visible = True
            .Top = rowtop
            .Left = 0
            .Width = 108
            .Height = 18
            .Font.Size = 8
        End With
        
        rowtop = rowtop + 18
    Next
    
    
    'Formatage
    Frame_TestsList.ScrollHeight = rowtop
    'Frame_TestsList.Height = rowtop
    'CommandButton_Valider.Top = Frame_TestsList.Top + Frame_TestsList.Height + Label_Instructions.Top
    'CommandButton_Annuler.Top = CommandButton_Valider.Top
    
    'Height = CommandButton_Valider.Top + 2 * CommandButton_Valider.Height
    
    cancel_Synth2Tests = False
End Sub
