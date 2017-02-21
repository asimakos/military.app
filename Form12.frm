VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form12 
   Caption         =   "Εμφάνιση  αναλυτικών   δεδόμενων"
   ClientHeight    =   6405
   ClientLeft      =   1860
   ClientTop       =   1350
   ClientWidth     =   8925
   LinkTopic       =   "Form12"
   ScaleHeight     =   6405
   ScaleWidth      =   8925
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
      Height          =   2175
      Left            =   4680
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2055
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3625
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "Καλυμένες    θέσεις"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Αριθμός   ενδιαφερόμενων:"
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Αριθμός   θέσεων   εργασίας:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
