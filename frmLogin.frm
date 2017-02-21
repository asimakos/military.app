VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Change(x As Boolean)
Form2.Adodc1.Visible = x
Form3.Adodc1.Visible = x
Form1.insert_group.Enabled = x
Form1.insert_com_group.Enabled = x
Form1.append_data.Enabled = x
Form18.DataGrid1.AllowUpdate = x
Form18.Command2.Enabled = x
Form18.Command3.Enabled = x
Form19.DataGrid1.AllowUpdate = x
Form19.Command2.Enabled = x
Form19.Command3.Enabled = x
Form20.DataGrid1.AllowUpdate = x
Form20.Command2.Enabled = x
Form20.Command3.Enabled = x
Form21.DataGrid1.AllowUpdate = x
Form21.Command2.Enabled = x
Form21.Command3.Enabled = x
Form22.DataGrid1.AllowUpdate = x
Form22.Command2.Enabled = x
Form22.Command3.Enabled = x
End Sub

Private Sub cmdCancel_Click()
    
    txtUserName.Text = " "
    txtPassword.Text = " "
    End
End Sub

Private Sub cmdOK_Click()
   If (Trim(txtUserName.Text) = "Admin") And (Trim(txtPassword.Text) = "123") Then
   frmLogin.Hide
   Form1.Show
   Form1.Caption = "Επιστράτευση  724 ΤΜΧ" & "   ... Εχετε εισέλθει ως Admin !"
   Call Change(True)
   ElseIf (Trim(txtUserName.Text) = "User") And (Trim(txtPassword.Text) = "123") Then
   frmLogin.Hide
   Form1.Show
   Form1.Caption = "Επιστράτευση  724 ΤΜΧ" & "   ... Εχετε εισέλθει ως User !"
   Call Change(False)
   Else
   MsgBox "Λάθος κωδικός ή όνομα! Προσπαθήστε ξανά"
   txtUserName.Text = " "
   txtPassword.Text = " "
   End If
End Sub
