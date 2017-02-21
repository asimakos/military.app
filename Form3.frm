VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Λόχοι  724  ΤΜΧ   (Καταχώρηση  δεδομένων  κατά   άτομο)"
   ClientHeight    =   8595
   ClientLeft      =   1845
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form3"
   ScaleHeight     =   8595
   ScaleWidth      =   8775
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   4440
      TabIndex        =   41
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      DataField       =   "status"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      DataField       =   "place_resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   39
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      DataField       =   "resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6720
      TabIndex        =   37
      Top             =   6120
      Width           =   1935
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "Form3.frx":0000
      Left            =   1800
      List            =   "Form3.frx":0064
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text14 
      DataField       =   "status"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   1800
      TabIndex        =   34
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1032
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6720
      TabIndex        =   32
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form3.frx":0283
      Left            =   2040
      List            =   "Form3.frx":02B7
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "f_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "mil_num"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form3.frx":036D
      Left            =   6720
      List            =   "Form3.frx":0377
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      DataField       =   "degree"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form3.frx":0385
      Left            =   6720
      List            =   "Form3.frx":0392
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      DataField       =   "comments"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "occup"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form3.frx":03A8
      Left            =   2040
      List            =   "Form3.frx":03C1
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "Form3.frx":03E1
      Left            =   1800
      List            =   "Form3.frx":0445
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Προσθήκη   εγγραφής"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Αποθήκευση   εγγραφής"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      DataField       =   "p_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      DataField       =   "p_rank"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      DataField       =   "p_corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      DataField       =   "PYR"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      DataField       =   "ammunition"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   8040
      Width           =   8415
      _ExtentX        =   14843
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
      RecordSource    =   "staff"
      Caption         =   "Εγγραφές   για   Λόχους   724  ΤΜΧ"
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
   Begin VB.Label Label3 
      Caption         =   "Νομός   διαμονής"
      Height          =   255
      Left            =   2640
      TabIndex        =   38
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Τόπος   διαμονής"
      Height          =   375
      Left            =   5040
      TabIndex        =   36
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Δευτερεύουσα         ειδικότητα"
      Height          =   735
      Left            =   240
      TabIndex        =   33
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Ειδικότητα"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Βαθμός"
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Όπλο"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Ονοματεπώνυμο"
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "¨Ονομα  πατέρα"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Στρατιωτικός  αριθμός"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "     Επάγγελμα"
      Height          =   375
      Left            =   240
      MousePointer    =   11  'Hourglass
      TabIndex        =   25
      ToolTipText     =   "Εισαγωγή  του  επαγγέλματός του"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Γραμματικές  γνώσεις"
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "ΠΥΡ"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Α/Α  του  ΜΕΕ  "
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Προβλεπόμενος   οπλισμός"
      Height          =   495
      Left            =   4440
      TabIndex        =   21
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label15 
      Caption         =   "       Ημερομηνία"
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Παρατηρήσεις"
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Adodc1.Recordset.AddNew
Command2.Enabled = True

End Sub

Private Sub Command2_Click()

On Error GoTo CancelUpdate
Text9.Text = Combo5.Text
Text14.Text = Combo6.Text
Text10.Text = Combo1.Text
Text11.Text = Combo4.Text
Text12.Text = Combo2.Text
Text13.Text = Combo3.Text
Text18.Text = Text6.Text
Text17.Text = "N"

Adodc1.Recordset.Update
MsgBox "Η εγγραφή αποθηκεύτηκε επιτυχώς!"
Exit Sub

CancelUpdate:
MsgBox Err.Description
Adodc1.Recordset.CancelUpdate

End Sub
