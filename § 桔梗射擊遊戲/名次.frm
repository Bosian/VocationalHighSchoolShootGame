VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�W���Ʀ�]"
   ClientHeight    =   3900
   ClientLeft      =   3735
   ClientTop       =   3390
   ClientWidth     =   7920
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "�W��.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   7920
   Begin VB.CommandButton Command1 
      Caption         =   "�M���Ҧ��Ʀ�"
      Height          =   495
      Index           =   2
      Left            =   6000
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Index           =   1
      Left            =   4440
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   28
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '�z��
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '�z��
      Caption         =   "�W�r"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '�z��
      Caption         =   "�W��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   7440
      TabIndex        =   24
      Top             =   2760
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   7440
      TabIndex        =   23
      Top             =   2280
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   7440
      TabIndex        =   22
      Top             =   1800
      Width           =   210
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   7440
      TabIndex        =   21
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   2880
      TabIndex        =   20
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   2880
      TabIndex        =   19
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   2880
      TabIndex        =   18
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2880
      TabIndex        =   17
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   7440
      TabIndex        =   16
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   2880
      TabIndex        =   15
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "5."
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "4."
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   645
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lab(4, 4)
Dim unid As Integer
Public Sub swap(a As Object, b As Object, c As Object, d As Object, e As Object, f As Object) '�洫��
t = a.Caption
a.Caption = b.Caption
b.Caption = t

t = c.Caption
c.Caption = d.Caption
d.Caption = t

t = e.Caption
e.Caption = f.Caption
f.Caption = t
End Sub
Private Sub Command1_Click(index As Integer) '(0)�T�w(1)����(2)�M�űƦ�
Select Case index
    Case 0
        Unload Me
    Case 1
        If win = 1 Then
            If MsgBox("�A�T�w�n���ثe���Ʀ�O����?", 52, "�T��") = 6 Then
                win = 0
                Call Form_Load
                Unload Me
            End If
        Else
            Unload Me
        End If
    Case 2
        If MsgBox("�z�T�w�n�M�ũҦ����Ʀ��", 52, "�R��") = 6 Then
            Unload Me
            Kill "SAVE\rage.sav"
            ppi = 0
            For i = 2 To 4
                For j = 0 To 4
                    lab(i, j) = ""
                Next
            Next
        End If
End Select
End Sub
Private Sub Form_Activate() '���Ұʡ�
Form4.Width = 0
For f = 0 To 8010 Step 230
    Form4.Width = f
    DoEvents
Next
End Sub
Private Sub Form_Load() '�W�����J��
Open "SAVE\rage.sav" For Append As #1
Close #1
Open "SAVE\rage.sav" For Input As #1
    If Not EOF(1) Then
        For i = 2 To 4
            For j = 0 To 4
                Input #1, ppi, lab(i, j)
            Next
        Next
    End If
Close #1
Call acbd
If win = 1 Then 'Ĺ�ο�h
    Command1(0).Enabled = False
    For j = 0 To 4
        If unid = 0 Then '�M�w����
            If lab(2, j) = "" Then ppi = j: unid = 1
        End If
    Next
    If unid = 0 Then
        If mark >= Val(Label4(4).Caption) Then
            ppi = 4
        Else
            Command1(0).Enabled = True
            MsgBox "���Ƭ��G " & mark & " �L�i�J�Ʀ�"
            Exit Sub
        End If
    End If
    Label4(ppi).Caption = mark
    Text1(ppi).Visible = True
    Select Case Y(1)
        Case 75 '²��
            Label3(ppi).Caption = "²��"
        Case 50 '���q
            Label3(ppi).Caption = "���q"
        Case 25 '�x��
            Label3(ppi).Caption = "�x��"
        Case 1 '�a��
            Label3(ppi).Caption = "�a��"
    End Select
End If
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer) '���UEnter��
If KeyAscii = 13 Then
    If Text1(ppi).Text = "" Then
        MsgBox "�z�|����J�W�r�A�Э��s��J��", 16, "�T��"
        Exit Sub
    Else
        Text1(ppi).Visible = False
        Label2(ppi).Caption = Text1(ppi).Text
        Label2(ppi).Visible = True
        Call decide(Label4)
        Command1(0).Enabled = True
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer) '������
unid = 0
win = 0
For j = 0 To 4
    lab(2, j) = Label2(j).Caption
    lab(3, j) = Label3(j).Caption
    lab(4, j) = Val(Label4(j).Caption)
Next
Call saves
End Sub
Private Sub saves() '�����W���æs�ɡ�
Open "SAVE\rage.sav" For Output As #1
    For i = 2 To 4
        For j = 0 To 4
            Write #1, ppi, lab(i, j)
        Next
    Next
Close #1
End Sub
Private Sub acbd() '�٭�{�ǡ�(�NŪ�����W����ƼȦs�٭즨�i�θ��)
For j = 0 To 4
    Label2(j).Caption = lab(2, j)
    Label3(j).Caption = lab(3, j)
    Label4(j).Caption = lab(4, j)
Next
End Sub
Private Sub decide(a As Object) '�ƧǺt��k��
For i = 0 To 3
    For j = i + 1 To 4
        If Val(a(i).Caption) < Val(a(j).Caption) Then
            Call swap(Label4(i), Label4(j), Label3(i), Label3(j), Label2(i), Label2(j))
        End If
    Next
Next
End Sub
