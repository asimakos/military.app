VERSION 5.00
Begin VB.Form Form25 
   Caption         =   "�����    (����������   ��������  ���   �����������)"
   ClientHeight    =   6330
   ClientLeft      =   2100
   ClientTop       =   1575
   ClientWidth     =   7830
   LinkTopic       =   "Form28"
   ScaleHeight     =   6330
   ScaleWidth      =   7830
   Begin VB.OptionButton Option4 
      Caption         =   "����������"
      Height          =   615
      Left            =   5760
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "������������"
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   3960
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "��������"
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "������ "
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form28.frx":0000
      Left            =   1920
      List            =   "Form28.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form28.frx":0004
      Left            =   1920
      List            =   "Form28.frx":0026
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����������  ��������  ���  ������ && ���������������"
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����������    ���   "
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form28.frx":0087
      Left            =   2280
      List            =   "Form28.frx":0091
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����������   ���     "
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����������   ���    "
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��������  ��  ��������  ��� ...  "
      Height          =   975
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��������  ���� "
      Height          =   615
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "              �����:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   3480
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "����������"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "�����   ��������"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "�������"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "��   ����   ��   ��������   ��������"
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
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stat(1 To 33, 1 To 2) As Integer

Sub d_pyr()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM d_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM d_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ���������� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ���������� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form30.Adodc1.Recordset.AddNew
Form30.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form30.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form30.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form30.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form30.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form30.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form30.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form30.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form30.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form30.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form30.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form30.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form30.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form30.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form30.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form30.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form30.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form30.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form30.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form30.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form30.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form30.Show
Form31.Show
End Sub

Sub c_pyr()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM c_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM c_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ������������ ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ������������ ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form29.Adodc1.Recordset.AddNew
Form29.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form29.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form29.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form29.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form29.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form29.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form29.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form29.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form29.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form29.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form29.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form29.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form29.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form29.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form29.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form29.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form29.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form29.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form29.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form29.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form29.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form29.Show
Form31.Show
End Sub
Sub b_pyr()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM b_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM b_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ������� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ������� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form27.Adodc1.Recordset.AddNew
Form27.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form27.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form27.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form27.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form27.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form27.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form27.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form27.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form27.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form27.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form27.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form27.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form27.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form27.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form27.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form27.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form27.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form27.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form27.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form27.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form27.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form27.Show
Form31.Show
End Sub
Sub a_pyr()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM a_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM a_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ����� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ����� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form28.Adodc1.Recordset.AddNew
Form28.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form28.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form28.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form28.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form28.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form28.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form28.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form28.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form28.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form28.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form28.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form28.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form28.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form28.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form28.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form28.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form28.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form28.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form28.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form28.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form28.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form28.Show
Form31.Show
End Sub

Sub d_nom()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM d_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM d_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ���������� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ���������� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form30.Adodc1.Recordset.AddNew
Form30.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form30.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form30.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form30.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form30.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form30.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form30.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form30.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form30.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form30.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form30.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form30.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form30.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form30.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form30.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form30.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form30.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form30.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form30.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form30.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form30.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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
Form30.Show
Form31.Show

End Sub

Sub c_nom()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM c_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM c_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ������������ ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ������������ ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form29.Adodc1.Recordset.AddNew
Form29.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form29.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form29.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form29.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form29.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form29.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form29.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form29.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form29.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form29.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form29.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form29.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form29.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form29.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form29.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form29.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form29.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form29.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form29.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form29.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form29.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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
Form29.Show
Form31.Show

End Sub
Sub b_nom()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM b_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM b_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ������� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ������� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form27.Adodc1.Recordset.AddNew
Form27.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form27.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form27.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form27.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form27.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form27.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form27.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form27.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form27.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form27.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form27.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form27.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form27.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form27.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form27.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form27.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form27.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form27.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form27.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form27.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form27.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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
Form27.Show
Form31.Show

End Sub

Sub a_nom()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
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
CMD.CommandText = "SELECT spec,rank,corps,team FROM a_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM a_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ����� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ����� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form28.Adodc1.Recordset.AddNew
Form28.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form28.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form28.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form28.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form28.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form28.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form28.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form28.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form28.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form28.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form28.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form28.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form28.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form28.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form28.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form28.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form28.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form28.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form28.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form28.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form28.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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
Form28.Show
Form31.Show

End Sub

Sub group_d_spec()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM d_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM d_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ���������� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ���������� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form30.Adodc1.Recordset.AddNew
Form30.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form30.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form30.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form30.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form30.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form30.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form30.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form30.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form30.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form30.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form30.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form30.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form30.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form30.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form30.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form30.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form30.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form30.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form30.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form30.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form30.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form30.Show
Form31.Show
End Sub

Sub group_c_spec()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM c_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM c_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ������������ ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ������������ ����  � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form29.Adodc1.Recordset.AddNew
Form29.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form29.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form29.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form29.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form29.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form29.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form29.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form29.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form29.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form29.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form29.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form29.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form29.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form29.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form29.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form29.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form29.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form29.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form29.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form29.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form29.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form29.Show
Form31.Show
End Sub

Sub group_b_spec()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM b_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM b_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ���� ������� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ������� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form27.Adodc1.Recordset.AddNew
Form27.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form27.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form27.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form27.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form27.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form27.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form27.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form27.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form27.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form27.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form27.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form27.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form27.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form27.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form27.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form27.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form27.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form27.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form27.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form27.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form27.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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

Form27.Show
Form31.Show
End Sub


Sub group_a_spec()

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
MsgBox "��� �������� �������� ������� �� " & Combo1.Text & "!"
Exit Sub
End If

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM staff WHERE p_spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSK = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,team FROM a_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSM = CMD.Execute

Set CMD = Nothing
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT COUNT(*) FROM a_party WHERE spec LIKE '%" & Trim(Combo1.Text) & "%'"
CMD.CommandType = adCmdText
Set RSN = CMD.Execute

If IsNull(RSN) Then
MsgBox "��� �������� ������ ��o� ����� ���� ��� ��������� !"
Exit Sub
End If

If (RSK.Fields(0) <= RSN.Fields(0)) Then
k = RSK.Fields(0)
Else
k = RSN.Fields(0)
End If

If (k <= 0) Then
MsgBox "��� �������� ������ ���� ����� ���� � ��������� ���������! ����������� ����."
Exit Sub
End If

i = 1
Do
Form28.Adodc1.Recordset.AddNew
Form28.Adodc1.Recordset.Fields(0) = RS.Fields(0)
Form28.Adodc1.Recordset.Fields(1) = RS.Fields(1)
Form28.Adodc1.Recordset.Fields(2) = RS.Fields(2)
Form28.Adodc1.Recordset.Fields(3) = RS.Fields(3)
Form28.Adodc1.Recordset.Fields(4) = RS.Fields(4)
Form28.Adodc1.Recordset.Fields(5) = RS.Fields(5)
Form28.Adodc1.Recordset.Fields(6) = RS.Fields(6)
Form28.Adodc1.Recordset.Fields(7) = RS.Fields(7)
Form28.Adodc1.Recordset.Fields(8) = RS.Fields(8)
Form28.Adodc1.Recordset.Fields(9) = RS.Fields(9)
Form28.Adodc1.Recordset.Fields(10) = RS.Fields(10)
Form28.Adodc1.Recordset.Fields(11) = RS.Fields(11)
Form28.Adodc1.Recordset.Fields(12) = RS.Fields(12)
Form28.Adodc1.Recordset.Fields(13) = RS.Fields(13)
Form28.Adodc1.Recordset.Fields(14) = RS.Fields(14)
Form28.Adodc1.Recordset.Fields(15) = RS.Fields(15)
Form28.Adodc1.Recordset.Fields(16) = RSM.Fields(0)
Form28.Adodc1.Recordset.Fields(17) = RSM.Fields(1)
Form28.Adodc1.Recordset.Fields(18) = RSM.Fields(2)
Form28.Adodc1.Recordset.Fields(19) = RSM.Fields(3)
Form28.Adodc1.Recordset.Update
i = i + 1
RS.MoveNext
If IsNull(RS) Then
MsgBox "��� �������� ������� " & Combo1.Text & " ��� �� �������� � ������ �����!"
Exit Do
End If
Loop Until i >= k
w = Combo1.Text
MsgBox "�� �������� ��� ������������� ���  " & w & "  �����  " & CStr(i - 1)

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
Form28.Show
Form31.Show
 
End Sub
Sub Check_status()
Dim i As Integer, j As Integer

For i = 1 To Combo1.ListCount
If (stat(i, 1) > stat(i, 2)) Then
j = stat(i, 1) - stat(i, 2)
MsgBox "� ������� ��� " & Form26.MSHFlexGrid3.TextMatrix(i, 0) & " ������� �� ���� ���� " & CStr(j) & " !"
End If
Next
End Sub

Sub d_party()

Dim res(1 To 33) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  d_party WHERE spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 34
For i = 1 To Combo1.ListCount - 1
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 2) = res(i)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  d_in WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next

Form26.MSHFlexGrid3.Cols = 2
Form26.MSHFlexGrid3.Rows = 34
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid3.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 1) = res(i)
Form26.MSHFlexGrid3.TextMatrix(i, 1) = res(i)
Next
Set CMD = Nothing
Set Conn = Nothing

End Sub


Sub c_party()

Dim res(1 To 33) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  c_party WHERE spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 34
For i = 1 To Combo1.ListCount - 1
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 2) = res(i)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  c_in WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next

Form26.MSHFlexGrid3.Cols = 2
Form26.MSHFlexGrid3.Rows = 34
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid3.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 1) = res(i)
Form26.MSHFlexGrid3.TextMatrix(i, 1) = res(i)
Next
Set CMD = Nothing
Set Conn = Nothing
End Sub

Sub b_party()

Dim res(1 To 33) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  b_party WHERE spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 34
For i = 1 To Combo1.ListCount - 1
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 2) = res(i)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  b_in WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next

Form26.MSHFlexGrid3.Cols = 2
Form26.MSHFlexGrid3.Rows = 34
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid3.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 1) = res(i)
Form26.MSHFlexGrid3.TextMatrix(i, 1) = res(i)
Next
Set CMD = Nothing
Set Conn = Nothing

End Sub

Sub a_party()

Dim res(1 To 33) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(*) FROM  a_party WHERE spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'" & _
                  "AND isNull(surname)"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing
Set RS = Nothing

Form26.MSHFlexGrid1.Cols = 2
Form26.MSHFlexGrid1.Rows = 34
For i = 1 To Combo1.ListCount - 1
Form26.MSHFlexGrid1.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 2) = res(i)
Form26.MSHFlexGrid1.TextMatrix(i, 1) = res(i)
Next

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn

For i = 1 To Combo1.ListCount
CMD.CommandText = "SELECT COUNT(p_spec) FROM  a_in WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next

Form26.MSHFlexGrid3.Cols = 2
Form26.MSHFlexGrid3.Rows = 34
For i = 1 To Combo1.ListCount
Form26.MSHFlexGrid3.TextMatrix(i, 0) = Combo1.List(i - 1)
stat(i, 1) = res(i)
Form26.MSHFlexGrid3.TextMatrix(i, 1) = res(i)
Next
Set CMD = Nothing
Set Conn = Nothing

End Sub

Private Sub Command2_Click()

Dim res(1 To 33) As Integer
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer

If (Option1.Value = True) Then
Call a_party
Form26.Label1.Caption = Form26.Label1.Caption & " ��� ����� ����"
ElseIf (Option2.Value = True) Then
Call b_party
Form26.Label1.Caption = Form26.Label1.Caption & " ��� ������� ����"
ElseIf (Option3.Value = True) Then
Call c_party
Form26.Label1.Caption = Form26.Label1.Caption & " ��� ������������ ����"
ElseIf (Option4.Value = True) Then
Call d_party
Form26.Label1.Caption = Form26.Label1.Caption & " ��� ���������� ����"
End If

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn

For i = 1 To Combo1.ListCount - 1
CMD.CommandText = "SELECT COUNT(p_spec) FROM  staff WHERE p_spec LIKE '%" & Trim(Combo1.List(i - 1)) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute
res(i) = RS.Fields(0)
Set RS = Nothing
Next
Set CMD = Nothing
Set Conn = Nothing

Form26.MSHFlexGrid2.Cols = 2
Form26.MSHFlexGrid2.Rows = 34
For i = 1 To Combo1.ListCount - 1
Form26.MSHFlexGrid2.TextMatrix(i, 0) = Combo1.List(i - 1)
Form26.MSHFlexGrid2.TextMatrix(i, 1) = res(i)
Next
Form26.Show
Call Check_status
End Sub

Private Sub Command3_Click()

If (Option1.Value = True) Then
Call group_a_spec
ElseIf (Option2.Value = True) Then
Call group_b_spec
ElseIf (Option3.Value = True) Then
Call group_c_spec
ElseIf (Option4.Value = True) Then
Call group_d_spec
End If

End Sub

Private Sub Command4_Click()

If (Option1.Value = True) Then
Call a_nom
ElseIf (Option2.Value = True) Then
Call b_nom
ElseIf (Option3.Value = True) Then
Call c_nom
ElseIf (Option4.Value = True) Then
Call d_nom
End If
End Sub

Private Sub Command5_Click()

If (Option1.Value = True) Then
Call a_pyr
ElseIf (Option2.Value = True) Then
Call b_pyr
ElseIf (Option3.Value = True) Then
Call c_pyr
ElseIf (Option4.Value = True) Then
Call d_pyr
End If

End Sub

Private Sub Command6_Click()

If (Option1.Value = True) Then
Command3.Caption = "���������� ��� ��� ����� ����"
Command4.Caption = "���������� ��� ��� ����� ����"
Command5.Caption = "���������� ��� ��� ����� ����"
ElseIf (Option2.Value = True) Then
Command3.Caption = "���������� ��� ��� ������� ����"
Command4.Caption = "���������� ��� ��� ������� ����"
Command5.Caption = "���������� ��� ��� ������� ����"
ElseIf (Option3.Value = True) Then
Command3.Caption = "���������� ��� ��� ������������ ����"
Command4.Caption = "���������� ��� ��� ������������ ����"
Command5.Caption = "���������� ��� ��� ������������ ����"
ElseIf (Option4.Value = True) Then
Command3.Caption = "���������� ��� ��� ���������� ����"
Command4.Caption = "���������� ��� ��� ���������� ����"
Command5.Caption = "���������� ��� ��� ���������� ����"
End If
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

Dim data(1 To 33) As String
Dim i As Integer

Combo1.Clear
data(1) = "����. ����. ����"
data(2) = "����������� ���/���"
data(3) = "������ ������"
data(4) = "�������"
data(5) = "����������� - ������ �����������"
data(6) = "������������"
data(7) = "������"
data(8) = "������ ������"
data(9) = "���������������"
data(10) = "�������"
data(11) = "��������"
data(12) = "������������"
data(13) = "���������"
data(14) = "������ �����������"
data(15) = "������ �������� ������������"
data(16) = "������ �/�"
data(17) = "���������"
data(18) = "����������"
data(19) = "��������"
data(20) = "���������"
data(21) = "�������� �����������"
data(22) = "��������"
data(23) = "������� ��������"
data(24) = "������� ����"
data(25) = "��������� �/�"
data(26) = "��������� �/�"
data(27) = "��������� �/�"
data(28) = "��������� �/�"
data(29) = "��������� ���. ������"
data(30) = "��������� ���/��� ��"
data(31) = "��������� �/�"
data(32) = "��������� �/�-�/�"

For i = 1 To 32
Combo1.AddItem data(i), i - 1
Next

End Sub
