VERSION 5.00
Begin VB.Form frmprincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Darker"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdchequear 
      Caption         =   "&Check"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdverpassword 
      Caption         =   "&View Password"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblpassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IMPORTANT NOTE: The program is not fully tested there may
'be errors.
Dim Password As String

Private Sub cmdchequear_Click()
'This is a very simple method you could read the
'password from windows registry, a file, database
'even the password you read may be encripted.
'The following code is just for testing
If Password = "matias" Then
    MsgBox "Correct Password.", vbInformation, "Password"
Else
    MsgBox "Passwords do not match.", vbExclamation, "Password"
End If
End Sub

Private Sub cmdverpassword_Click()
'shows the value of password
MsgBox "The typed password is: " & Password & " And the correct password is: matias", vbInformation, "Password"
End Sub

Private Sub Form_Load()
'eeemmm, just credits!
'Created by Matías Ariel Villagarcía!
frmprincipal.Caption = "Password Darker Version: " & App.Major & "." & App.Minor & "." & App.Revision & " By Matías A. Villagarcía"
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
'This is the code you should use in your programs
If KeyAscii = 8 Then
    If Len(txtpassword.Text) = 1 Then
        txtpassword.Text = ""
        Password = ""
    End If
    If Len(txtpassword.Text) = 0 Then Exit Sub
    txtpassword.Text = Mid(txtpassword.Text, 1, Len(txtpassword.Text) - 1)
    Password = Mid(Password, 1, Len(Password) - 1)
    Exit Sub
End If
If Password = "" Then
    Password = Chr(KeyAscii)
    txtpassword.Text = "*"
Else
    Password = Password & Chr(KeyAscii)
    txtpassword.Text = txtpassword.Text & "*"
End If
End Sub
