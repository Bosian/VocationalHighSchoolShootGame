VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "����ܱ�g���C��"
   ClientHeight    =   4890
   ClientLeft      =   3735
   ClientTop       =   3390
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "����ܱ�.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   7920
   Begin VB.CommandButton Command2 
      Caption         =   "�W�@��"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�U�@��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��N��/�@N��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4005
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7635
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer '���ƼȦs
Dim X As Integer '���ĴX���t��k�����G�ܼ�
Dim total As Integer '�̲׭����Ȧs
Dim eee As String '��L��
Dim fff As String
Dim ggg As String
Dim bbb As String '��L��
Private Sub Form_Load() '�����J��
n = 13 '=---------------------------������̫ܳ�@��
total = n
Call multi
Call thirteen
End Sub
Private Sub Form_Activate() '���Ұʡ�
Form2.Width = 0
For f = 0 To 8010 Step 230
    Form2.Width = f
    DoEvents
Next
End Sub
Private Sub Command1_Click() '�U�@�ӡ�
Command2.Enabled = True
n = n + 1
Call multi '=-------------------------�ĴX��
Call test '=--------------------------��������
Select Case X
    Case 2
        Call two
    Case 3
        Call three
    Case 4
        Call four
    Case 5
        Call five
    Case 6
        Call six
    Case 7
        Call seven
    Case 8
        Call eight
    Case 9
        Call nine
    Case 10
        Call ten
    Case 11
        Call eleven
    Case 12
        Call twelve
    Case 13
        Call thirteen
End Select
If n = total Then Command1.Enabled = False
End Sub
Private Sub Command2_Click() '�W�@�ӡ�
Command1.Enabled = True
n = n - 1
Call multi '=-------------------------�ĴX��
Call test '=--------------------------��������
Select Case X '=----------------------��ܭ���
    Case 1
        Call one
    Case 2
        Call two
    Case 3
        Call three
    Case 4
        Call four
    Case 5
        Call five
    Case 6
        Call six
    Case 7
        Call seven
    Case 8
        Call eight
    Case 9
        Call nine
    Case 10
        Call ten
    Case 11
        Call eleven
    Case 12
        Call twelve
End Select
If n = 1 Then Command2.Enabled = False
End Sub
Private Sub one() '�Ĥ@�������e��
Label1.Caption = "�������� v 4.6.1" & vbCrLf & "�ץ��G�N���^�������d��"
End Sub
Private Sub two() '�ĤG�������e��
Label1.Caption = "�������� v 4.7" & vbCrLf & "1.�ץ��ܱ�ϥΩG�N���^�ɡA�j����    �L�ܱ𤣷|�l�媺���D" & vbCrLf & "2.�N�쥻3�Ӱ}�C�X�֦��@��2���}�C" & vbCrLf & "3.�W�[�F�������\��" & vbCrLf & "4.�W�j�����j���P�_��= ="
End Sub
Private Sub three() '�ĤT�������e��
Label1.Caption = "�������� v 5.0" & vbCrLf & "�� 1.�ϥΤF�j�q���ʵe�ĪG�έץ��x      ��������" & vbCrLf & "�� 2.�s�W�F����ۭq�\��" & vbCrLf & "    3.�ץ��F1UP�@�����X�{�b������       �~�������D" & vbCrLf & "    4.�W�[�D��檺���仡���P����]      �w�P�B�\��" & vbCrLf & "�� 5.�W�[�F�֫���U�ֶ]���\��"
End Sub
Private Sub four() '�ĥ|�������e��
Call changes
Label1.Caption = "�������� v 5.0" & vbCrLf & "�� 6.�W�[�j����" & vbCrLf & "      �Ϊk�G" & vbCrLf & "      ��|��ɯ�Q5�㰣�ɡA" & vbCrLf & "      ��" & eee & fff & ggg & "�����ǫ��C"
End Sub
Private Sub five() '�Ĥ��������e��
Call changes
Label1.Caption = "�������� v 5.1" & vbCrLf & "1.�ץ����U" & bbb & "���Ϳ��~������" & vbCrLf & "2.�ץ��j���۪����k�G" & eee & fff & ggg & """�P��""���U   �~���@��" & vbCrLf & "3.�ץ����|��ɬ�3�Ψ䭿�ƥi�o��   ������"
End Sub
Private Sub six() '�Ĥ��������e��
Label1.Caption = "�������� v 6.0" & vbCrLf & "�� 1.�W�[�ĤG���M�ĤT��" & vbCrLf & "�� 2.�W�[�������ƭp��" & vbCrLf & "�� 3.�ץ�40000���i�o�@���P" & vbCrLf & "�� 4.�ץ�50000���i�o�@���R"
End Sub
Private Sub seven() '�ĤC�������e��
Label1.Caption = "�������� v 6.1" & vbCrLf & "�� 1.�W�[�ĥ|��(�p�]����)"
End Sub
Private Sub eight() '�ĤK�������e��
Label1.Caption = "�������� v 6.2" & vbCrLf & "1.�ץ�6.1���@�ǿ��~" & vbCrLf & "2.�̷Ӥ��P���רӼW�[�ĥ|������q    �y���ͼƶq"
End Sub
Private Sub nine() '�ĤE�������e��
Label1.Caption = "�������� v 6.3" & vbCrLf & "1.�ץ�6.2���@�ǿ��~" & vbCrLf & "2.�s�W²�����ƦW"
End Sub
Private Sub ten() '�ĤQ�������e��
Label1.Caption = "�������� v 6.4" & vbCrLf & "�� 1.�W�[�Ĥ���"
End Sub
Private Sub eleven() '��11�������e��
Label1.Caption = "�������� v 6.5" & vbCrLf & "�� 1.�W�[���s�����d(���դ�)" & vbCrLf & "    2.�[�֦^��^�]" & vbCrLf & "    3.�W�[�Y�|��ɥi�o���ƪ�����"
End Sub
Private Sub twelve() '��12�������e��
Label1.Caption = "�������� v 6.6" & vbCrLf & "�� 1.�W�[����j�����S��(���դ�)"
End Sub
Private Sub thirteen() '��13�������e��
Label1.Caption = "�������� v 7.0(�ĥ|�N)" & vbCrLf & "�� 1.�[�J�F�۰ʱ��ʦ��I��" & vbCrLf & "�� 2.�W�[�����S��" & vbCrLf & "�� 3.�֤߬[�c����½�s.."
End Sub
Private Sub multi() '�@�h�֭���
Label2.Caption = "��" & n & "��/�@" & total & "��"
End Sub
Private Sub test() '�������ơ�
X = n Mod 100 '=-----------------��������
End Sub
Private Sub Form_Unload(Cancel As Integer) '������桴
Form1.Show
End Sub
Private Sub changes() '�ഫ�r��
Call asc(a)
    eee = ss
Call asc(s)
    fff = ss
Call asc(d)
    ggg = ss
Call asc(enter)
    bbb = ss
End Sub
