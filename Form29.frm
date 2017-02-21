VERSION 5.00
Begin VB.Form Form28 
   Caption         =   "Λόχοι   (Καταχώρηση  εγγραφών  ανα   ειδικότητες)"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form28"
   ScaleHeight     =   5775
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form29.frx":0000
      Left            =   1680
      List            =   "Form29.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form29.frx":0004
      Left            =   1800
      List            =   "Form29.frx":002C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Στατιστικά  στοιχεία  για  θέσεις && ενδιαφερόμενους"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Καταχώρηση    για   "
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form29.frx":00A1
      Left            =   2040
      List            =   "Form29.frx":00AB
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Καταχώρηση   για     "
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Καταχώρηση   για    "
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Επιλέξτε  τα  κριτήρια  και ...  "
      Height          =   975
      Left            =   5280
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Επιλέξτε  ξανά "
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ειδικότητα"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Νομός   διομονής"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Πυρήνας"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   975
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
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
