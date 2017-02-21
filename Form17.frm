VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form17 
   Caption         =   "Διοικήσεως  λόχος  -  Επιλογή  προσωπικού  ανάλογα  με  τα   καθήκοντα"
   ClientHeight    =   6195
   ClientLeft      =   2220
   ClientTop       =   1290
   ClientWidth     =   7950
   LinkTopic       =   "Form17"
   ScaleHeight     =   6195
   ScaleWidth      =   7950
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form17.frx":0000
      Left            =   3480
      List            =   "Form17.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εμφάνιση  εγγραφών"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διαγραφή   δεδομένων"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3255
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5741
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Επιλέξτε  την  διμοιρία  του  προσωπικού"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   6735
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim x As String, i As Integer
Dim data(6) As String


Select Case Trim(Combo1.Text)
Case "Μισθοτροφοδοσίας"
x = "ΜΙΣΘ"
Case "Διαχύσεως  Υλικών"
x = "ΔΧΣΕΩΣ"
Case "Μεταφορών"
x = "ΜΕΤΑΦ"
Case "Μηχανημάτων"
x = "ΜΗΧΑΝ"
Case "Συντηρήσεως"
x = "ΣΥΝΤ"
Case "Διαβιβάσεων"
x = "ΔΙΑΒ"
Case "Δ/Σ"
x = "Δ/Σ"
End Select

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"

Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec,rank,corps,p_spec,p_rank," & _
                "surname,f_name FROM d_in WHERE team LIKE '%" & Trim(x) & "%'"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

If (RS.EOF) Then
MsgBox "Δεν υπάρχουν εγγραφές διαθέσιμες για " & Combo1.Text & ". Προσπαθήστε ξανά !"
Exit Sub
End If

Grid1.Cols = RS.Fields.Count + 1
Grid1.Rows = 10

data(0) = "ΕΙΔΙΚΟΤΗΤΑ"
data(1) = "ΒΑΘΜΟΣ (O?-A?)"
data(2) = "ΟΠΛΟ"
data(3) = "ΕΙΔΙΚΟΤΗΤΑ -(ΤΟΠΟΘΕΤΗΜΕΝΟΙ)"
data(4) = "ΒΑΘΜΟΣ"
data(5) = "ΟΝΟΜΑΤΕΜΩΝΥΜΟ"
data(6) = "ΠΑΤΡΩΝΥΜΟ"

For j = 1 To RS.Fields.Count
Grid1.TextMatrix(0, j) = data(j - 1)
Next

i = 1
While Not RS.EOF
If (i = Grid1.Rows) Then
Grid1.Rows = Grid1.Rows + 10
End If
For j = 0 To RS.Fields.Count - 1
If Not IsNull(RS.Fields(j)) Then
 Grid1.TextMatrix(i, j + 1) = RS.Fields(j)
 End If
Next j
RS.MoveNext
i = i + 1
 Wend

Label2.Caption = " Αριθμός εγγραφών για " & Combo1.Text & " ΔΡΙΑ :  " & CStr(i - 1)
Set RS = Nothing
Set CMD = Nothing
Set Conn = Nothing

End Sub

Private Sub Command2_Click()
Grid1.Clear
Label2.Caption = " "
End Sub
