VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "MIE App Settings Demo"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim AppSett As New uf_zApplication_Settings
AppSett.Show
Set AppSett = Nothing
End Sub

Private Sub CommandButton2_Click()
Dim AppSett As New uf_zApplication_Settings
Dim x As Boolean
x = AppSett.AddValue(Me.TextBox1, Me.TextBox2)
Set AppSett = Nothing
End Sub

Private Sub CommandButton3_Click()
Dim AppSett As New uf_zApplication_Settings
Dim x As Boolean
x = AppSett.UpdateValue(Me.TextBox3, Me.TextBox4)
Set AppSett = Nothing
End Sub

Private Sub CommandButton4_Click()
Dim AppSett As New uf_zApplication_Settings
Dim x As String ''' Note The return type is a string!
x = AppSett.GetValue(Me.TextBox5)
Me.TextBox6.Text = x
Set AppSett = Nothing
End Sub
