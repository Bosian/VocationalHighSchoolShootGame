VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  '單線固定
   Caption         =   "鍵盤設定"
   ClientHeight    =   3405
   ClientLeft      =   3120
   ClientTop       =   3075
   ClientWidth     =   9210
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9210
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   12
      Top             =   1560
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   7
      Left            =   7920
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   6
      Left            =   6960
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   5
      Left            =   6000
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   4
      Left            =   5040
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   3
      Left            =   4080
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   2
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   1
      Left            =   2160
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "預設值"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消 (ESC)"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定 (Enter)"
      Default         =   -1  'True
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "變更按鍵"
      Height          =   615
      Index           =   0
      Left            =   1200
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w As Integer '表單原始寬度暫存
Private Sub Command1_Click() '確定●
Open "SAVE\keycode.sav" For Output As #1
Write #1, Key_up, key_down, key_left, key_right, a, s, d, enter
Close #1
Call Command2_Click
End Sub
Private Sub Command2_Click() '取消●
Open "SAVE\keycode.sav" For Input As #1
Input #1, Key_up, key_down, key_left, key_right, a, s, d, enter
Close #1
Unload Me
Form1.Show
End Sub
Private Sub Command3_Click() '預設值●
Key_up = 38
key_down = 40
key_left = 37
key_right = 39
a = 65
s = 83
d = 68
enter = 13
n = Array("↑", "↓", "←", "→", "A", "S", "D", "Enter")
MSFlexGrid1.Row = 1
For i = 0 To 7
    MSFlexGrid1.Col = i + 1
    MSFlexGrid1.Text = n(i)
Next
End Sub
Private Sub Form_Activate() '啟動●
Form3.Width = 0
For f = 0 To w Step 230
    Form3.Width = f
    DoEvents
Next
End Sub
Private Sub Form_Load() '表單載入●
w = Form3.Width
MSFlexGrid1.Rows = 2 '=----------------------------------------------------8列
MSFlexGrid1.Cols = 9 '=----------------------------------------------------8欄
MSFlexGrid1.Col = 0 '=-----------------------------------------------------從第0欄開始放
MSFlexGrid1.Row = 1 '=-----------------------------------------------------從第一列開始放
MSFlexGrid1.Text = "對應按鍵"
MSFlexGrid1.Row = 0 '=-----------------------------------------------------從第0列開始放
n = Array("上", "下", "左", "右", "破魔之箭", "射箭", "咒術飭回", "暫停") 'n=索引值XXX,YYY.......
For i = 0 To 7
    MSFlexGrid1.Col = i + 1 '=---------------------------------------------第i+1欄開始放
    MSFlexGrid1.Text = n(i) '=---------------------------------------------放入n(i)的索引值
Next
Open "SAVE\Keycode.sav" For Append As #1
Close #1
Open "SAVE\keycode.sav" For Input As #1
If Not EOF(1) Then
    Input #1, Key_up, key_down, key_left, key_right, a, s, d, enter
    Call asc(Key_up): aaa = ss
    Call asc(key_down): bbb = ss
    Call asc(key_left): ccc = ss
    Call asc(key_right): ddd = ss
    Call asc(a): eee = ss
    Call asc(s): fff = ss
    Call asc(d): ggg = ss
    Call asc(enter): hhh = ss
    m = Array(aaa, bbb, ccc, ddd, eee, fff, ggg, hhh)
    MSFlexGrid1.Row = 1
    For i = 0 To 7
        MSFlexGrid1.Col = i + 1
        MSFlexGrid1.Text = m(i)
    Next
End If
Close #1
Text1.Text = "更變按鍵的方法：" & vbCrLf & "    點一下""變更按鍵""，" & vbCrLf & "    再輸入鍵盤上的按鍵即可。" & "如要變更↑↓←→請按預設值"
End Sub
Private Sub Form_Unload(Cancel As Integer) '表單移除●
Call Command2_Click
End Sub
Private Sub Option1_Click(index As Integer)
For i = 0 To 7
    Select Case index
        Case i
            MSFlexGrid1.Col = i + 1
            MSFlexGrid1.Text = ""
            Option1(i).Caption = "請按下"
    End Select
Next
End Sub
Private Sub Option1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
For i = 0 To 7
    Select Case index
        Case i
            MSFlexGrid1.Col = i + 1
            MSFlexGrid1.Text = ""
            Select Case i
                Case 0
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    Key_up = KeyCode
                Case 1
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    key_down = KeyCode
                Case 2
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    key_left = KeyCode
                Case 3
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    key_right = KeyCode
                Case 4
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    a = KeyCode
                Case 5
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    s = KeyCode
                Case 6
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    d = KeyCode
                Case 7
                    Call asc(KeyCode)
                    MSFlexGrid1.Text = ss
                    enter = KeyCode
            End Select
            Option1(i).Value = False
            Option1(i).Caption = "變更按鍵"
    End Select
Next
End Sub
