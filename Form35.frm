VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form32 
   Caption         =   "Προεπισκόπηση    εκτύπωσης    για    διοίκηση - επιτελείο "
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   990
   ClientWidth     =   11580
   LinkTopic       =   "Form32"
   ScaleHeight     =   5775
   ScaleWidth      =   11580
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Κλείσιμο"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   5160
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
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
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8281
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Form35.frx":0000
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form32.Hide
End Sub
