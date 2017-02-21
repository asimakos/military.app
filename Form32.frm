VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form29 
   Caption         =   "Καταχώρηση  δεδομένων  για   λόχο   γεφυροσκευής"
   ClientHeight    =   6390
   ClientLeft      =   285
   ClientTop       =   1230
   ClientWidth     =   8190
   LinkTopic       =   "Form29"
   ScaleHeight     =   6390
   ScaleWidth      =   8190
   Begin VB.CommandButton Command4 
      Caption         =   "Προεπισκόπηση   εκτύπωσης"
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Εκτύπωση"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Προσθήκη   εγγραφής"
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εύρεση"
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διαγραφή   εγγραφής"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "p_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      DataField       =   "p_rank"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   3360
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      DataField       =   "p_corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text7 
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text8 
      DataField       =   "f_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text9 
      DataField       =   "mil_num"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text10 
      DataField       =   "occup"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text11 
      DataField       =   "degree"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text12 
      DataField       =   "PYR"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text13 
      DataField       =   "MEE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text14 
      DataField       =   "ammunition"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text16 
      DataField       =   "comments"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text17 
      DataField       =   "resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text18 
      DataField       =   "place_resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text19 
      DataField       =   "sec_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6000
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=war"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=war"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "c_in"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form32.frx":0000
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2760
      TabIndex        =   17
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "surname"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form32.frx":002D
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "                       Γεφυροσκευής   λόχος"
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "spec"
         Caption         =   "Καθήκοντα"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "corps"
         Caption         =   "Σώμα"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "rank"
         Caption         =   "Βαθμός"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "p_spec"
         Caption         =   "Πρωτ.  ειδικότητα"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "sec_spec"
         Caption         =   "Δευτερ.  ειδικότητα"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "p_rank"
         Caption         =   "Βαθμός"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "p_corps"
         Caption         =   "Σώμα"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "surname"
         Caption         =   "Ονοματεπώνυμο"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "f_name"
         Caption         =   "Πατρώνυμο"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "resid"
         Caption         =   "Τόπος  διαμονής"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "place_resid"
         Caption         =   "Νομός  διαμονής"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "mil_num"
         Caption         =   "Στρατ.  αριθμός"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "occup"
         Caption         =   "Επάγγελμα"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "degree"
         Caption         =   "Γραμ.  γνώσεις"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "PYR"
         Caption         =   "Πυρήνας"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "MEE"
         Caption         =   "Α/Α  του  ΜΕΕ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "ammunition"
         Caption         =   "Οπλισμός"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "date"
         Caption         =   "Ημερομηνία"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "comments"
         Caption         =   "Παρατηρήσεις"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "team"
         Caption         =   "Ομάδα - Δρία"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Εισάγετε  το   ονοματεπώνυμο"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Adodc1.Recordset.Bookmark = DataCombo1.SelectedItem
End Sub

Private Sub Command2_Click()
Dim reply

reply = MsgBox("Επιθυμείτε την αλλαγή των δεδομένων?", vbYesNo)
If (reply = vbYes) Then
Adodc1.Recordset.Delete
ElseIf (reply = vbNo) Then
Adodc1.Recordset.CancelUpdate
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command4_Click()

Dim x As String, y As String
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset


y = "ΚΑΘΗΚΟΝΤΑ   ΒΑΘΜΟΣ   ΣΩΜΑ   ΒΑΘΜΟΣ   ΣΩΜΑ  ΟΝΟΜΑΤΕΠΩΝΥΜΟ   ΠΑΤΡΩΝΥΜΟ   ΣΤΡ.ΑΡΙΘΜΟΣ   ΠΡΩΤ.ΕΙΔΙΚΟΤΗΤΑ  " & _
"ΔΕΥΤ.ΕΙΔΙΚΟΤΗΤΑ   ΤΟΠΟΣ ΔΙΑΜΟΝΗΣ   ΝΟΜΟΣ  ΔΙΑΜΟΝΗΣ   ΕΠΑΓΓΕΛΜΑ   Α/Α ΤΟΥ ΜΕΕ   ΠΥΡΗΝΑΣ   ΟΠΛΙΣΜΟΣ   ΗΜΕΡΟΜΗΝΙΑ   ΠΑΡΑΤΗΡΗΣΕΙΣ  "

y = y & "    " & "                                      "
                                                        
Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT * FROM c_in"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

While Not (RS.EOF)
x = x & RS.Fields(16) & "  " & RS.Fields(17) & "  " & RS.Fields(18) & "  " & RS.Fields(0) & "  " & RS.Fields(1) & _
RS.Fields(2) & "  " & RS.Fields(3) & "  " & RS.Fields(4) & "  " & RS.Fields(5) & "  " & RS.Fields(6) & _
RS.Fields(7) & "  " & RS.Fields(8) & "  " & RS.Fields(9) & "  " & RS.Fields(10) & "  " & RS.Fields(11) & _
RS.Fields(12) & "  " & RS.Fields(13) & "  " & RS.Fields(14) & "  " & RS.Fields(15)
RS.MoveNext
Wend
Form35.RichTextBox1.Text = y & x
Form35.Show
End Sub

Private Sub Command5_Click()

Form35.CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
If Form35.RichTextBox1.SelLength = 0 Then
Form35.CommonDialog1.Flags = Form35.CommonDialog1.Flags + cdlPDAllPages
Else
Form35.CommonDialog1.Flags = Form35.CommonDialog1.Flags + cdlPDSelection
End If
Form35.CommonDialog1.ShowPrinter
Printer.Print ""
Form35.RichTextBox1.SelPrint Form35.CommonDialog1.hDC
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim reply

reply = MsgBox("Επιθυμείτε την αλλαγή των δεδομένων?", vbYesNo)
If (reply = vbYes) Then
Cancel = False
DataGrid1.Columns(ColIndex).Value = OldValue
ElseIf (reply = vbNo) Then
Cancel = True
Adodc1.Recordset.CancelUpdate
End If
End Sub


