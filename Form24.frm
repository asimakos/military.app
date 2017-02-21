VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form24 
   Caption         =   "Καταχώρηση    δεδομένων  για  διοίκηση - επιτελείο"
   ClientHeight    =   6390
   ClientLeft      =   1680
   ClientTop       =   960
   ClientWidth     =   8190
   LinkTopic       =   "Form24"
   ScaleHeight     =   6390
   ScaleWidth      =   8190
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "Form24.frx":0000
      Left            =   6840
      List            =   "Form24.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   28
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Δημιουργία   λίστας"
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Κλείσιμο   λίστας"
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "’νοιγμα   λίστας"
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Προσθήκη   εγγραφής"
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διαγραφή   εγγραφής"
      Height          =   375
      Left            =   4440
      TabIndex        =   23
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εύρεση"
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   360
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form24.frx":0004
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2880
      TabIndex        =   21
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "surname"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form24.frx":0031
      Height          =   3975
      Left            =   120
      TabIndex        =   19
      Top             =   960
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
      Caption         =   "                       Διοίκηση - Επιτελείο"
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
   Begin VB.TextBox Text19 
      DataField       =   "sec_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text18 
      DataField       =   "place_resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text17 
      DataField       =   "resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text16 
      DataField       =   "comments"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text15 
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text14 
      DataField       =   "ammunition"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text13 
      DataField       =   "MEE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text12 
      DataField       =   "PYR"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text11 
      DataField       =   "degree"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text10 
      DataField       =   "occup"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text9 
      DataField       =   "mil_num"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text8 
      DataField       =   "f_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text7 
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      DataField       =   "p_corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      DataField       =   "p_rank"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   3360
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      DataField       =   "p_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      DataField       =   "corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      DataField       =   "rank"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5880
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
      RecordSource    =   "com_in"
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
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   615
      Left            =   1800
      OleObjectBlob   =   "Form24.frx":0046
      SourceDoc       =   "D:\downloads\Army\ΔΙΟΙΚΗΣΗ.xls"
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Εισάγετε  το   ονοματεπώνυμο"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form24"
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

OLE1.DoVerb (vbOLEPrimary)
OLE1.DoVerb (vbOLEShow)
OLE1.DoVerb (vbOLEOpen)
OLE1.DoVerb (vbOLEUIActivate)
End Sub

Private Sub Command5_Click()

Set xlBook = GetObject("D:\downloads\Army\ΔΙΟΙΚΗΣΗ.xls")
xlBook.Application.Visible = False
xlBook.Close
Set xlBook = Nothing
End Sub

Private Sub Command6_Click()

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim CMD As New ADODB.Command
Dim Conn As New ADODB.Connection
Dim RS As ADODB.Recordset
Dim i As Integer, j As Integer
Dim x As String, y As String

Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

xlSheet.Cells(1, 7).Value = "ΠΙΝΑΚΑΣ (1)"
xlSheet.Cells(1, 14).Value = "ΓΕΣ/ΔΕΣ/Υποδ. 05/79"
xlSheet.Cells(2, 1).Value = "ΑΠΟΡΡΗΤΟ   ΕΠΙΣΤΡΑΤΕΥΣΕΩΣ"
xlSheet.Cells(2, 4).Value = "ΟΡΓΑΝΩΣΗ  Τ(2)"
xlSheet.Cells(2, 5).Value = "ΔΚΣΗΣ-ΕΠΙΤΕΛΕΙΟΥ"
xlSheet.Cells(2, 6).Value = "ΤΗΣ (3)"
xlSheet.Cells(2, 7).Value = "4252 ΑΣ    ΑΣ"
xlSheet.Cells(2, 13).Value = "ΣΧΕΔΙΟ  ΕΠΙΣΤΡΑΤΕΥΣΕΩΣ"
xlSheet.Cells(3, 1).Value = "ΕΑ  40 ΜΕ"
xlSheet.Cells(3, 14).Value = "ΑΠΡΙΛΙΟΣ 2005"
xlSheet.Cells(4, 1).Value = "ΑΣ  4252ΑΣ (724 ΤΜΧ)"
xlSheet.Cells(7, 1).Value = "ΠΡΟΒΛΕΠΟΝΤΑΙ ΑΠΟ ΠΟΥ (5)"
xlSheet.Cells(7, 4).Value = "ΤΟΠΟΘΕΤΗΜΕΝΟΙ"
xlSheet.Cells(7, 12).Value = "Π"
xlSheet.Cells(7, 13).Value = "Α/Α"
xlSheet.Cells(7, 14).Value = "ΠΡΟΒΛΕΠΟ-"
xlSheet.Cells(8, 8).Value = "ΟΝΟΜΑ"
xlSheet.Cells(8, 11).Value = "ΓΡΑΜΜΑΤΙΚΕΣ"
xlSheet.Cells(8, 12).Value = "Υ"
xlSheet.Cells(8, 13).Value = "ΤΟΥ"
xlSheet.Cells(8, 14).Value = "ΜΕΝΟΣ"
xlSheet.Cells(9, 1).Value = "ΚΑΘΗΚΟΝΤΑ-ΕΙΔΙΚΟΤΗΤΑ"
xlSheet.Cells(9, 2).Value = "ΒΑΘΜΟΣ"
xlSheet.Cells(9, 3).Value = "ΟΠΛΟ"
xlSheet.Cells(9, 4).Value = "ΕΙΔΙΚΟΤΗΤΑ"
xlSheet.Cells(9, 5).Value = "ΒΑΘΜΟΣ"
xlSheet.Cells(9, 6).Value = "ΟΠΛΟ"
xlSheet.Cells(9, 7).Value = "ΟΝΟΜΑΤΕΠΩΝΥΜΟ"
xlSheet.Cells(9, 8).Value = "ΠΑΤΕΡΑ"
xlSheet.Cells(9, 9).Value = "ΣΑ"
xlSheet.Cells(9, 10).Value = "ΕΠΑΓΓΕΛΜΑ"
xlSheet.Cells(9, 11).Value = "ΓΝΩΣΕΙΣ"
xlSheet.Cells(9, 12).Value = "Ρ"
xlSheet.Cells(9, 13).Value = "ΜΕΕ"
xlSheet.Cells(9, 14).Value = "ΟΠΛΙΣΜΟΣ"
xlSheet.Cells(9, 15).Value = "ΗΜΕΡΟΜΗΝΙΑ"
xlSheet.Cells(9, 16).Value = "ΠΑΡΑΤΗΡΗΣΕΙΣ"
For i = 1 To 16
xlSheet.Cells(11, i).Value = CStr(i)
Next
xlSheet.Cells(12, 1).Value = "1. ΔΙΟΙΚΗΣΗ"
xlSheet.Cells(17, 1).Value = "2. ΕΠΙΤΕΛΕΙΟ"
xlSheet.Cells(18, 1).Value = "α. 1ο ΓΡΑΦΕΙΟ"
xlSheet.Cells(22, 1).Value = "β. 2ο ΓΡΑΦΕΙΟ"
xlSheet.Cells(23, 1).Value = "(1) Διεύθυνση"
xlSheet.Cells(27, 1).Value = "(2) Ομάδα Ελέγχου ΠΒΧΠ"
xlSheet.Cells(31, 1).Value = "γ. 3ο ΓΡΑΦΕΙΟ"
xlSheet.Cells(37, 1).Value = "δ. 4ο ΓΡΑΦΕΙΟ"

Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT spec FROM com_party"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

While Not RS.EOF
If Not IsNull(RS.Fields(0)) Then
List1.AddItem RS.Fields(0)
End If
RS.MoveNext
Wend
For i = 0 To 2
xlSheet.Cells(13 + i, 1).Value = List1.List(i)
Next
For i = 0 To 1
xlSheet.Cells(19 + i, 1).Value = List1.List(3 + i)
Next
For i = 0 To 2
xlSheet.Cells(24 + i, 1).Value = List1.List(5 + i)
Next
For i = 0 To 1
xlSheet.Cells(28 + i, 1).Value = List1.List(8 + i)
Next
For i = 0 To 3
xlSheet.Cells(32 + i, 1).Value = List1.List(10 + i)
Next
For i = 0 To 1
xlSheet.Cells(38 + i, 1).Value = List1.List(14 + i)
Next

j = 12
While Not Adodc1.Recordset.EOF
If Not IsNull(Adodc1.Recordset.Fields(16).Value) Then
x = Adodc1.Recordset.Fields(16).Value
End If
For i = j + 1 To 39
If (Trim(xlSheet.Cells(i, 1).Value) = x) Then
j = i
If Not IsNull(Adodc1.Recordset.Fields(17).Value) Then
y = Adodc1.Recordset.Fields(17).Value
xlSheet.Cells(i, 2).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(18).Value) Then
y = Adodc1.Recordset.Fields(18).Value
xlSheet.Cells(i, 3).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(16).Value) Then
y = Adodc1.Recordset.Fields(16).Value
xlSheet.Cells(i, 4).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(0).Value) Then
y = Adodc1.Recordset.Fields(0).Value
xlSheet.Cells(i, 5).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(1).Value) Then
y = Adodc1.Recordset.Fields(1).Value
xlSheet.Cells(i, 6).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(2).Value) Then
y = Adodc1.Recordset.Fields(2).Value
xlSheet.Cells(i, 7).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(3).Value) Then
y = Adodc1.Recordset.Fields(3).Value
xlSheet.Cells(i, 8).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(4).Value) Then
y = Adodc1.Recordset.Fields(4).Value
xlSheet.Cells(i, 9).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(9).Value) Then
y = Adodc1.Recordset.Fields(9).Value
xlSheet.Cells(i, 10).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(12).Value) Then
y = Adodc1.Recordset.Fields(12).Value
xlSheet.Cells(i, 11).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(11).Value) Then
y = Adodc1.Recordset.Fields(11).Value
xlSheet.Cells(i, 12).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(10).Value) Then
y = Adodc1.Recordset.Fields(10).Value
xlSheet.Cells(i, 13).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(13).Value) Then
y = Adodc1.Recordset.Fields(13).Value
xlSheet.Cells(i, 14).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(14).Value) Then
y = Adodc1.Recordset.Fields(14).Value
xlSheet.Cells(i, 15).Value = y
End If
If Not IsNull(Adodc1.Recordset.Fields(15).Value) Then
y = Adodc1.Recordset.Fields(15).Value
xlSheet.Cells(i, 16).Value = y
End If
Exit For
End If
Next
Adodc1.Recordset.MoveNext
Wend
xlSheet.SaveAs "D:\downloads\Army\ΔΙΟΙΚΗΣΗ.xls"
xlBook.Close
xlApp.Quit
MsgBox "Το αρχείο xls με τους επιστρατεύσιμους δημιουργήθηκε επιτυχώς!"
Set RS = Nothing
Set RSK = Nothing
Set CMD = Nothing
Set Conn = Nothing
Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing
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



