VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "������������  724  ���"
   ClientHeight    =   6720
   ClientLeft      =   1980
   ClientTop       =   1605
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9285
   Visible         =   0   'False
   Begin VB.Label Label1 
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Menu insert_mem 
      Caption         =   "��������   ���������"
      Begin VB.Menu insert_party 
         Caption         =   "�����"
         Begin VB.Menu insert_group 
            Caption         =   "�������   ����������"
         End
         Begin VB.Menu insert_staff 
            Caption         =   "�������   ����������"
         End
      End
      Begin VB.Menu insert_com 
         Caption         =   "��������� - ��������"
         Begin VB.Menu insert_com_group 
            Caption         =   "�������   ����������"
         End
         Begin VB.Menu insert_com_staff 
            Caption         =   "�������   ����������"
         End
      End
      Begin VB.Menu exit 
         Caption         =   "Exit  (Log out)"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu append_data 
      Caption         =   "����������   ��������"
      Begin VB.Menu append_com 
         Caption         =   "��  �������� - ���������"
      End
      Begin VB.Menu append_party 
         Caption         =   "��  ������"
      End
      Begin VB.Menu staff_append 
         Caption         =   "���������  ���������  ���   ����������"
      End
      Begin VB.Menu in_rec 
         Caption         =   "�������������  ��������"
         Begin VB.Menu inse_a 
            Caption         =   "������  �����"
         End
         Begin VB.Menu inse_b 
            Caption         =   "��������  �����"
         End
         Begin VB.Menu inse_c 
            Caption         =   "������������  �����"
         End
         Begin VB.Menu inse_d 
            Caption         =   "����������  �����"
         End
         Begin VB.Menu inse_com 
            Caption         =   "�������� - ���������"
         End
      End
   End
   Begin VB.Menu select_mem 
      Caption         =   "������  �����  ���  ������"
      Begin VB.Menu select_a 
         Caption         =   "������  �����"
         Begin VB.Menu rank_a 
            Caption         =   "������   ������"
         End
         Begin VB.Menu duty_a 
            Caption         =   "���������  �����"
         End
      End
      Begin VB.Menu select_b 
         Caption         =   "��������  �����"
         Begin VB.Menu rank_b 
            Caption         =   "������  ������"
         End
         Begin VB.Menu duty_b 
            Caption         =   "���������  �����"
         End
      End
      Begin VB.Menu select_c 
         Caption         =   "������������  �����"
         Begin VB.Menu rank_c 
            Caption         =   "������  ������"
         End
         Begin VB.Menu duty_c 
            Caption         =   "���������  �����"
         End
      End
      Begin VB.Menu select_d 
         Caption         =   "����������  �����"
         Begin VB.Menu rank_d 
            Caption         =   "������  ������"
         End
         Begin VB.Menu duty_d 
            Caption         =   "���������  �����"
         End
      End
      Begin VB.Menu select_com 
         Caption         =   "��������"
         Begin VB.Menu rank_com 
            Caption         =   "������  ������"
         End
         Begin VB.Menu duty_com 
            Caption         =   "���������  �����"
         End
      End
      Begin VB.Menu arith_data 
         Caption         =   "�����������  �����������  ���������"
      End
   End
   Begin VB.Menu chang_elem 
      Caption         =   "������  ���������  ���  ����"
      Begin VB.Menu chan_a 
         Caption         =   "������  �����"
      End
      Begin VB.Menu chan_b 
         Caption         =   "��������  �����"
      End
      Begin VB.Menu chan_c 
         Caption         =   "������������  �����"
      End
      Begin VB.Menu chan_d 
         Caption         =   "����������  �����"
      End
      Begin VB.Menu chan_com 
         Caption         =   "��������"
      End
      Begin VB.Menu diagr 
         Caption         =   "�����������  �������  �����"
      End
   End
   Begin VB.Menu help 
      Caption         =   "�������"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub append_com_Click()
Form6.Show
End Sub

Private Sub append_party_Click()
Form25.Show
End Sub

Private Sub arith_data_Click()
Dim hello
hello = Shell("D:\downloads\Army\Calc.exe", 1)
End Sub

Private Sub chan_a_Click()
Form19.Show
End Sub

Private Sub chan_b_Click()
Form20.Show
End Sub

Private Sub chan_c_Click()
Form21.Show
End Sub

Private Sub chan_com_Click()
Form18.Show
End Sub

Private Sub chan_d_Click()
Form22.Show
End Sub

Private Sub diagr_Click()
Form23.Show
End Sub

Private Sub duty_a_Click()
Form14.Show
End Sub

Private Sub duty_b_Click()
Form15.Show
End Sub

Private Sub duty_c_Click()
Form16.Show
End Sub

Private Sub duty_com_Click()
Form13.Show
End Sub

Private Sub duty_d_Click()
Form17.Show
End Sub

Private Sub exit_Click()
Form1.Hide
frmLogin.txtPassword = " "
frmLogin.txtUserName = " "
frmLogin.Show
End Sub

Private Sub Form_Load()
frmLogin.Show
Label1.Caption = "  ���������� " & "   " & Date
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim reply
reply = MsgBox("����� �������� ��� ������ �� ����������� ��� �������� ?", vbYesNo)
If (reply = vbYes) Then
End
ElseIf (reply = vbNo) Then
Cancel = -1
End If
End Sub

Private Sub help_Click()
frmAbout.Show
End Sub

Private Sub inse_a_Click()
Form28.Show
End Sub

Private Sub inse_b_Click()
Form27.Show
End Sub

Private Sub inse_c_Click()
Form29.Show
End Sub

Private Sub inse_com_Click()
Form24.Show
End Sub

Private Sub inse_d_Click()
Form30.Show
End Sub

Private Sub insert_com_group_Click()
Form4.Show
End Sub

Private Sub insert_com_staff_Click()
Form2.Show
Form2.Combo5.SetFocus
End Sub

Private Sub insert_group_Click()
Form5.Show
End Sub

Private Sub insert_staff_Click()
Form3.Show
End Sub

Private Sub rank_a_Click()
Form8.Show
End Sub

Private Sub rank_b_Click()
Form9.Show
End Sub

Private Sub rank_c_Click()
Form10.Show
End Sub

Private Sub rank_com_Click()
Form7.Show
End Sub

Private Sub rank_d_Click()
Form11.Show
End Sub

Private Sub staff_append_Click()
Form31.Show
End Sub
