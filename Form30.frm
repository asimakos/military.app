VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form28 
   Caption         =   "����������   ���������  ���   �����   ����"
   ClientHeight    =   6390
   ClientLeft      =   495
   ClientTop       =   1230
   ClientWidth     =   8190
   LinkTopic       =   "Form28"
   ScaleHeight     =   6390
   ScaleWidth      =   8190
   Begin VB.CommandButton Command4 
      Caption         =   "�������   ������"
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��������   ������"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "����������   ������"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��������   ��������"
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������   ��������"
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "p_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      DataField       =   "p_rank"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   3240
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
      DataField       =   "p_corps"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text7 
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text8 
      DataField       =   "f_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text9 
      DataField       =   "mil_num"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text10 
      DataField       =   "occup"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text11 
      DataField       =   "degree"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text12 
      DataField       =   "PYR"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text13 
      DataField       =   "MEE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text14 
      DataField       =   "ammunition"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text15 
      DataField       =   "date"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text16 
      DataField       =   "comments"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text17 
      DataField       =   "resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text18 
      DataField       =   "place_resid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text19 
      DataField       =   "sec_spec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
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
      RecordSource    =   "a_in"
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
      Bindings        =   "Form30.frx":0000
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2760
      TabIndex        =   18
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "surname"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form30.frx":002D
      Height          =   3975
      Left            =   120
      TabIndex        =   19
      Top             =   1080
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
      Caption         =   "                       ������   �����"
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "spec"
         Caption         =   "���������"
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
         Caption         =   "����"
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
         Caption         =   "������"
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
         Caption         =   "����.  ����������"
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
         Caption         =   "������.  ����������"
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
         Caption         =   "������"
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
         Caption         =   "����"
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
         Caption         =   "�������������"
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
         Caption         =   "���������"
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
         Caption         =   "�����  ��������"
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
         Caption         =   "�����  ��������"
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
         Caption         =   "�����.  �������"
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
         Caption         =   "���������"
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
         Caption         =   "����.  �������"
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
         Caption         =   "�������"
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
         Caption         =   "�/�  ���  ���"
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
         Caption         =   "��������"
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
         Caption         =   "����������"
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
         Caption         =   "������������"
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
         Caption         =   "����� - ����"
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
      Caption         =   "��������  ��   �������������"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Adodc1.Recordset.Bookmark = DataCombo1.SelectedItem
End Sub

Private Sub Command2_Click()
Dim reply

reply = MsgBox("���������� ��� ������ ��� ���������?", vbYesNo)
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


y = "���������   ������   ����   ������   ����  �������������   ���������   ���.�������   ����.����������  " & _
"����.����������   ����� ��������   �����  ��������   ���������   �/� ��� ���   �������   ��������   ����������   ������������  "

y = y & "    " & "                                      "
                                                        
Conn.ConnectionString = "Provider=MSDASQL.1;" & _
"Persist Security Info=False;Data Source=war"
Conn.Open
CMD.ActiveConnection = Conn
CMD.CommandText = "SELECT * FROM a_in"
CMD.CommandType = adCmdText
Set RS = CMD.Execute

While Not (RS.EOF)
x = x & RS.Fields(16) & "  " & RS.Fields(17) & "  " & RS.Fields(18) & "  " & RS.Fields(0) & "  " & RS.Fields(1) & _
RS.Fields(2) & "  " & RS.Fields(3) & "  " & RS.Fields(4) & "  " & RS.Fields(5) & "  " & RS.Fields(6) & _
RS.Fields(7) & "  " & RS.Fields(8) & "  " & RS.Fields(9) & "  " & RS.Fields(10) & "  " & RS.Fields(11) & _
RS.Fields(12) & "  " & RS.Fields(13) & "  " & RS.Fields(14) & "  " & RS.Fields(15)
RS.MoveNext
Wend
Form34.RichTextBox1.Text = y & x
Form34.Show
End Sub

Private Sub Command5_Click()

Form34.CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
If Form34.RichTextBox1.SelLength = 0 Then
Form34.CommonDialog1.Flags = Form34.CommonDialog1.Flags + cdlPDAllPages
Else
Form34.CommonDialog1.Flags = Form34.CommonDialog1.Flags + cdlPDSelection
End If
Form34.CommonDialog1.ShowPrinter
Printer.Print ""
Form34.RichTextBox1.SelPrint Form34.CommonDialog1.hDC
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
Set xlSheet = xlBook.Windows(1).Application

xlSheet.Cells(1, 7).Value = "������� (1)"
xlSheet.Cells(1, 14).Value = "���/���/����. 05/79"
xlSheet.Cells(2, 1).Value = "��������   ��������������"
xlSheet.Cells(2, 4).Value = "��������  �(2)"
xlSheet.Cells(2, 5).Value = "1�� ����� ���������"
xlSheet.Cells(2, 6).Value = "��� (3)"
xlSheet.Cells(2, 7).Value = "4252 ��    ��"
xlSheet.Cells(2, 13).Value = "������  ��������������"
xlSheet.Cells(3, 1).Value = "��  40 ��"
xlSheet.Cells(3, 14).Value = "�������� 2005"
xlSheet.Cells(4, 1).Value = "��  4252�� (724 ���)"
xlSheet.Cells(7, 1).Value = "������������ ��� ��� (5)"
xlSheet.Cells(7, 4).Value = "�������������"
xlSheet.Cells(7, 12).Value = "�"
xlSheet.Cells(7, 13).Value = "�/�"
xlSheet.Cells(7, 14).Value = "��������-"
xlSheet.Cells(8, 8).Value = "�����"
xlSheet.Cells(8, 11).Value = "�����������"
xlSheet.Cells(8, 12).Value = "�"
xlSheet.Cells(8, 13).Value = "���"
xlSheet.Cells(8, 14).Value = "�����"
xlSheet.Cells(9, 1).Value = "���������-����������"
xlSheet.Cells(9, 2).Value = "������"
xlSheet.Cells(9, 3).Value = "����"
xlSheet.Cells(9, 4).Value = "����������"
xlSheet.Cells(9, 5).Value = "������"
xlSheet.Cells(9, 6).Value = "����"
xlSheet.Cells(9, 7).Value = "�������������"
xlSheet.Cells(9, 8).Value = "������"
xlSheet.Cells(9, 9).Value = "��"
xlSheet.Cells(9, 10).Value = "���������"
xlSheet.Cells(9, 11).Value = "�������"
xlSheet.Cells(9, 12).Value = "�"
xlSheet.Cells(9, 13).Value = "���"
xlSheet.Cells(9, 14).Value = "��������"
xlSheet.Cells(9, 15).Value = "����������"
xlSheet.Cells(9, 16).Value = "������������"
xlSheet.SaveAs "c:\Army\������������.xls"
xlBook.Close
xlApp.Quit
MsgBox "�� ������ xls �� ���� ���������������� ������������� ��������!"
Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing
End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim reply

reply = MsgBox("���������� ��� ������ ��� ���������?", vbYesNo)
If (reply = vbYes) Then
Cancel = False
DataGrid1.Columns(ColIndex).Value = OldValue
ElseIf (reply = vbNo) Then
Cancel = True
Adodc1.Recordset.CancelUpdate
End If
End Sub


