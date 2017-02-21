VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form23 
   Caption         =   "Δύναμη   προσωπικού   για   κάθε   λόχο"
   ClientHeight    =   5775
   ClientLeft      =   2280
   ClientTop       =   1500
   ClientWidth     =   7500
   LinkTopic       =   "Form23"
   ScaleHeight     =   5775
   ScaleWidth      =   7500
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3855
      Left            =   1800
      OleObjectBlob   =   "Form23.frx":0000
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Απεικόνιση   των   δεδομένων"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label4 
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label3 
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim data(1 To 4, 1 To 2)
Dim i As Integer

data(1, 1) = "ΠΡΩΤΟΣ"
data(2, 1) = "ΔΕΥΤΕΡΟΣ"
data(3, 1) = "Γ/Φ"
data(4, 1) = "Δ/Σ"

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"

Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM a_in"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
data(1, 2) = RS.Fields(0)
Set RS = Nothing
Set CMD = Nothing
Set Conn = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"

Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM b_in"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
data(2, 2) = RS.Fields(0)
Set RS = Nothing
Set CMD = Nothing
Set Conn = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"

Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM c_in"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
data(3, 2) = RS.Fields(0)
Set RS = Nothing
Set CMD = Nothing
Set Conn = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"

Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM d_in"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
data(4, 2) = RS.Fields(0)
Set RS = Nothing
Set CMD = Nothing
Set Conn = Nothing

Label1.Caption = " ΠΡΩΤΟΣ : " & CStr(data(1, 2))
Label2.Caption = " ΔΕΥΤΕΡΟΣ : " & CStr(data(2, 2))
Label3.Caption = " Γ/Φ : " & CStr(data(3, 2))
Label4.Caption = " Δ/Σ : " & CStr(data(4, 2))

MSChart1.Visible = True
MSChart1.ChartType = VtChChartType3dBar
MSChart1.ChartData = data
End Sub


Private Sub MSChart1_Click()

Dim height As Integer

height = MSChart1.height + 500
MSChart1.Move 1800, 960, 5175, height

End Sub

