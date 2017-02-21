VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form26 
   Caption         =   "Εμφάνιση   αναλυτικών   δεδομένων"
   ClientHeight    =   6405
   ClientLeft      =   2055
   ClientTop       =   1545
   ClientWidth     =   8985
   LinkTopic       =   "Form26"
   ScaleHeight     =   6405
   ScaleWidth      =   8985
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   2175
      Left            =   4800
      TabIndex        =   4
      Top             =   3720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3836
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3836
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2175
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "Καλυμένες   θέσεις"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Αριθμός   θέσεων   εργασίας"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Αριθμός   ενδιαφερόμενων:"
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
