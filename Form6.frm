VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Διοίκηση - Επιτελείο   (Καταχώρηση  εγγραφών  ανα  ειδικότητες)"
   ClientHeight    =   5775
   ClientLeft      =   2610
   ClientTop       =   1665
   ClientWidth     =   7830
   LinkTopic       =   "Form6"
   ScaleHeight     =   5775
   ScaleWidth      =   7830
   Begin VB.CommandButton Command7 
      Caption         =   "Επιλέξτε  ξανά "
      Height          =   615
      Left            =   5160
      TabIndex        =   12
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Επιλέξτε  τα  κριτήρια  και ...  "
      Height          =   975
      Left            =   5520
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Καταχώρηση   για    "
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Καταχώρηση   για     "
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form6.frx":0000
      Left            =   2280
      List            =   "Form6.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Καταχώρηση    για   "
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Στατιστικά  στοιχεία  για  θέσεις && ενδιαφερόμενους"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   4560
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form6.frx":0018
      Left            =   1920
      List            =   "Form6.frx":003D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form6.frx":00A9
      Left            =   1920
      List            =   "Form6.frx":00AB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "με   βάση   τα   παρακάτω   κριτήρια"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Πυρήνας"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Νομός   διομονής"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ειδικότητα"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stat(1 To 11, 1 To 2) As Integer

Sub Check_status()
Dim i As Integer, j As Integer

For i = 1 To Combo1.ListCount
If (stat(i, 1) > stat(i, 2)) Then
j = stat(i, 1) - stat(i, 2)
MsgBox "Ο αριθμός των " & Form26.MSHFlexGrid3.TextMatrix(i, 0) & " ξεπερνά το όριο κατά " & CStr(j) & " !"
End If
Next
End Sub

Private Sub Command2_Click()

Dim res(1 To 11) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  com_party WHERE spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'" & _
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
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
Form12.MSHFlexGrid1.Cols = 2
Form12.MSHFlexGrid1.Rows = 14
For i = 1 To Combo1.ListCount
Form12.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 2) = res(i)
Form12.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  staff WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing

Form12.MSHFlexGrid2.Cols = 2
Form12.MSHFlexGrid2.Rows = 14
For i = 1 To Combo1.ListCount
Form12.MSHFlexGrid2.TextMatrix(i, 0) = Combo1.List(i - 1)
Form12.MSHFlexGrid2.TextMatrix(i, 1) = res(i)
Next

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  com_in WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next

Form12.MSHFlexGrid3.Cols = 2
Form12.MSHFlexGrid3.Rows = 14
For i = 1 To Combo1.ListCount
Form12.MSHFlexGrid3.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 1) = res(i)
Form12.MSHFlexGrid3.TextMatrix(i, 1) = res(i)
Next
Set CMD = Nothing
Set Conn = Nothing
Form12.Show
Call Check_status
End Sub

Private Sub Command3_Click()

Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset, RS1 As ADODB.Recordset
Dim RSN As ADODB.Recordset, RSM As ADODB.Recordset
Dim RSK As ADODB.Recordset
Dim x As String, y As String, w As String
Dim i As Integer, k As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT * FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

If IsNull(RS) Then
MsgBox "Δεν υπάρχουν εγγραφές σχετικά με " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM com_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM com_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "Δεν υπάρχουν θέσεις στην διοίκηση - επιτελείο για προσωπικό !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "Δεν υπάρχουν θέσεις στην διοίκηση - επιτελείο ή διαθέσιμο προσωπικό! Προσπαθήστε ξανά."
Exit Sub
End If

i = 1
Do
Form24.Adodc1.Recordset.AddNew
Form24.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form24.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form24.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form24.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form24.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form24.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form24.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form24.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form24.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form24.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form24.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form24.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form24.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form24.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form24.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form24.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form24.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form24.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form24.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form24.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form24.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "Δεν υπάρχουν αρκετοί " & Combo1.Text & " για να καλυφτεί ο πρώτος λόχος!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "Οι εγγραφές που καταχωρήθηκαν για  " & w & "  είναι  " & CStr(i - 1)

RS.MoveFirst
Set CMD = Nothing
CMD.ActiveConnection = Conn
While (Not RS.EOF)
CMD.CommandText = "UPDATE staff SET status='Y' WHERE surname= '" & RS.Fields(2) & "'"
CMD.CommandType = adCmdText
Set RS1 = CMD.Execute
RS.MoveNext
Wend

Set RS = Nothing
Set RS1 = Nothing
Set CMD = Nothing
Set RSN = Nothing
Set RSM = Nothing
Set RSK = Nothing

CMD.ActiveConnection = Conn
CMD.CommandText = "DELETE FROM staff WHERE status='Y'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
Set CMD = Nothing
Set RS = Nothing
Form24.Show
Form31.Show

End Sub

Private Sub Command4_Click()

Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset, RS1 As ADODB.Recordset
Dim RSN As ADODB.Recordset, RSM As ADODB.Recordset
Dim RSK As ADODB.Recordset
Dim x As String, y As String, w As String
Dim i As Integer, k As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT * FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'" & _
                  "and place_resid LIKE '%" & Trim(Combo2.Text) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

If IsNull(RS) Then
MsgBox "Δεν υπάρχουν εγγραφές σχετικά με " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'" & _
                  "and place_resid LIKE '%" & Trim(Combo2.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM com_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM com_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "Δεν υπάρχουν θέσεις στην διοίκηση - επιτελείο για προσωπικό !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "Δεν υπάρχουν θέσεις στην διοίκηση - επιτελείο ή διαθέσιμο προσωπικό! Προσπαθήστε ξανά."
Exit Sub
End If

i = 1
Do
Form24.Adodc1.Recordset.AddNew
Form24.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form24.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form24.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form24.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form24.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form24.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form24.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form24.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form24.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form24.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form24.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form24.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form24.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form24.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form24.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form24.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form24.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form24.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form24.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form24.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form24.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "Δεν υπάρχουν αρκετοί " & Combo1.Text & " για να καλυφτεί ο πρώτος λόχος!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "Οι εγγραφές που καταχωρήθηκαν για  " & w & "  είναι  " & CStr(i - 1)

RS.MoveFirst
Set CMD = Nothing
CMD.ActiveConnection = Conn
While (Not RS.EOF)
CMD.CommandText = "UPDATE staff SET status='Y' WHERE surname= '" & RS.Fields(2) & "'"
CMD.CommandType = adCmdText
Set RS1 = CMD.Execute
RS.MoveNext
Wend

Set RS = Nothing
Set RS1 = Nothing
Set CMD = Nothing
Set RSN = Nothing
Set RSM = Nothing
Set RSK = Nothing

CMD.ActiveConnection = Conn
CMD.CommandText = "DELETE FROM staff WHERE status='Y'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
Set CMD = Nothing
Set RS = Nothing
Form24.Show
Form31.Show
End Sub

Private Sub Command5_Click()

Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset, RS1 As ADODB.Recordset
Dim RSN As ADODB.Recordset, RSM As ADODB.Recordset
Dim RSK As ADODB.Recordset
Dim x As String, y As String, w As String
Dim i As Integer, k As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT * FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'" & _
                  "and place_resid LIKE '%" & Trim(Combo2.Text) & "%'" & _
                  "and PYR LIKE '%" & Trim(Combo3.Text) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

If IsNull(RS) Then
MsgBox "Δεν υπάρχουν εγγραφές σχετικά με " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'" & _
                  "and place_resid LIKE '%" & Trim(Combo2.Text) & "%'" & _
                  "and PYR LIKE '%" & Trim(Combo3.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM com_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM com_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "Δεν υπάρχουν θέσεις στην διοίκηση - επιτελείο για προσωπικό !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "Δεν υπάρχουν θέσεις στην διοίκηση - επιτελείο ή διαθέσιμο προσωπικό! Προσπαθήστε ξανά."
Exit Sub
End If

i = 1
Do
Form24.Adodc1.Recordset.AddNew
Form24.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form24.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form24.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form24.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form24.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form24.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form24.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form24.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form24.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form24.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form24.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form24.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form24.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form24.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form24.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form24.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form24.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form24.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form24.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form24.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form24.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "Δεν υπάρχουν αρκετοί " & Combo1.Text & " για να καλυφτεί ο πρώτος λόχος!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "Οι εγγραφές που καταχωρήθηκαν για  " & w & "  είναι  " & CStr(i - 1)

RS.MoveFirst
Set CMD = Nothing
CMD.ActiveConnection = Conn
While (Not RS.EOF)
CMD.CommandText = "UPDATE staff SET status='Y' WHERE surname= '" & RS.Fields(2) & "'"
CMD.CommandType = adCmdText
Set RS1 = CMD.Execute
RS.MoveNext
Wend

Set RS = Nothing
Set RS1 = Nothing
Set CMD = Nothing
Set RSN = Nothing
Set RSM = Nothing
Set RSK = Nothing

CMD.ActiveConnection = Conn
CMD.CommandText = "DELETE FROM staff WHERE status='Y'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
Set CMD = Nothing
Set RS = Nothing
Form24.Show
Form31.Show

End Sub

Private Sub Command6_Click()
If ((Combo2.ListIndex = -1) And (Combo3.ListIndex = -1)) Then
Command3.Visible = True
ElseIf (Combo3.ListIndex = -1) Then
Command4.Visible = True
Else
Command5.Visible = True
End If
End Sub

Private Sub Command7_Click()
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Combo3.ListIndex = -1
End Sub

Private Sub Form_Load()

Dim data(1 To 12) As String
Dim i As Integer

Combo1.Clear
data(1) = "ΔΙΟΙΚΗΤΗΣ"
data(2) = "ΥΠΟΔΙΟΙΚΗΤΗΣ"
data(3) = "ΥΠΑΣΠΙΣΤΗΣ"
data(4) = "ΓΡΑΦΕΑΣ"
data(5) = "ΔΙΕΥΘΥΝΤΗΣ"
data(6) = "ΣΚΑΠΑΝΕΑΣ"
data(7) = "ΥΠΑΞΙΩΜΑΤΙΚΟΣ  ΑΠΟΛΥΜΑΝΣΕΩΣ"
data(8) = "ΙΑΤΡΟΣ"
data(9) = "ΑΞΙΩΜΑΤΙΚΟΣ  ΑΝΑΓΝΩΡΙΣΗΣ"

Command3.Caption = "Καταχώρηση  για" & " Διοίκηση - Επιτελείο"
Command4.Caption = "Καταχώρηση  για" & " Διοίκηση - Επιτελείο"
Command5.Caption = "Καταχώρηση  για" & " Διοίκηση - Επιτελείο"

For i = 1 To 3
Combo1.AddItem data(i), i - 1
Next
Combo1.AddItem "*---------------------*", 3
For i = 4 To 9
Combo1.AddItem data(i), i
Next

End Sub
