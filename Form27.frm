VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form27 
   Caption         =   "Καταχώρηση   δεδομένων"
   ClientHeight    =   4590
   ClientLeft      =   4065
   ClientTop       =   1395
   ClientWidth     =   4680
   LinkTopic       =   "Form27"
   ScaleHeight     =   4590
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      DataField       =   "spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      DataField       =   "rank"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "p_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "p_rank"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "p_corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      DataField       =   "f_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      DataField       =   "mil_num"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      DataField       =   "occup"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      DataField       =   "degree"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      DataField       =   "PYR"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      DataField       =   "MEE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text14 
      DataField       =   "ammunition"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text15 
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text16 
      DataField       =   "comments"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   3840
      Width           =   4335
      _ExtentX        =   7646
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=warx"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "warx"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "com_party"
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
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
