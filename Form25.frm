VERSION 5.00
Begin VB.Form Form25 
   Caption         =   "Λόχοι   (Καταχώρηση  εγγραφών  ανα  ειδικότητες)"
   ClientHeight    =   4950
   ClientLeft      =   2160
   ClientTop       =   1395
   ClientWidth     =   7830
   LinkTopic       =   "Form25"
   ScaleHeight     =   4950
   ScaleWidth      =   7830
   Begin VB.OptionButton Option4 
      Caption         =   "Δ/Σ"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Γ/Φ"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form25.frx":0000
      Left            =   2400
      List            =   "Form25.frx":005E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Πρώτος"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Δεύτερος"
      Height          =   435
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form25.frx":026A
      Left            =   1800
      List            =   "Form25.frx":028C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Αναλυτικά   στοιχεία"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Καταχώρηση   για   "
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Ειδικότητα"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Τόπος   διομονής"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Επιλέξτε  έναν  από  τους   παρακάτω  λόχους:"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   4455
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

Dim res(1 To 40) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset, RS1 As ADODB.Connection
Dim i As Integer
Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=warx"
Conn.Open
CMD.ActiveConnection = Conn

If (Option1.Value = True) Then      'First choice
Combo1.Enabled = True
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  a_party WHERE spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=warx"
Conn.Open
CMD.ActiveConnection = Conn
Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 40
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  staff WHERE p_spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing

Form26.MSHFlexGrid2.Cols = 2
Form26.MSHFlexGrid2.Rows = 40
For i = 1 To Combo2.ListCount
Form26.MSHFlexGrid2.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid2.TextMatrix(i, 1) = res(i)
Next

ElseIf (Option2.Value = True) Then       'Second  choice
Combo1.Enabled = True
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  b_party WHERE spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=warx"
Conn.Open
CMD.ActiveConnection = Conn
Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 40
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  staff WHERE p_spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing

Form26.MSHFlexGrid2.Cols = 2
Form26.MSHFlexGrid2.Rows = 40
For i = 1 To Combo2.ListCount
Form26.MSHFlexGrid2.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid2.TextMatrix(i, 1) = res(i)
Next

ElseIf (Option3.Value = True) Then       'Third choice
Combo1.Enabled = True
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  c_party WHERE spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=warx"
Conn.Open
CMD.ActiveConnection = Conn
Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 40
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  staff WHERE p_spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing

Form26.MSHFlexGrid2.Cols = 2
Form26.MSHFlexGrid2.Rows = 40
For i = 1 To Combo2.ListCount
Form26.MSHFlexGrid2.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid2.TextMatrix(i, 1) = res(i)
Next

ElseIf (Option4.Value = True) Then        'Fourth choice
Combo1.Enabled = True
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  d_party WHERE spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=warx"
Conn.Open
CMD.ActiveConnection = Conn
Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 40
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  staff WHERE p_spec LIKE '" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing

Form26.MSHFlexGrid2.Cols = 2
Form26.MSHFlexGrid2.Rows = 40
For i = 1 To Combo2.ListCount
Form26.MSHFlexGrid2.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid2.TextMatrix(i, 1) = res(i)
Next
End If
Form26.Show

End Sub

