VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '單線固定
   Caption         =   "關於桔梗射擊遊戲"
   ClientHeight    =   4890
   ClientLeft      =   3735
   ClientTop       =   3390
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "關於桔梗.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   7920
   Begin VB.CommandButton Command2 
      Caption         =   "上一個"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一個"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "第N頁/共N頁"
      BeginProperty Font 
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "標楷體"
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
Dim n As Integer '頁數暫存
Dim X As Integer '為第幾頁演算法之結果變數
Dim total As Integer '最終頁的暫存
Dim eee As String '鍵盤↓
Dim fff As String
Dim ggg As String
Dim bbb As String '鍵盤↑
Private Sub Form_Load() '表單載入●
n = 13 '=---------------------------直接顯示最後一頁
total = n
Call multi
Call thirteen
End Sub
Private Sub Form_Activate() '表單啟動●
Form2.Width = 0
For f = 0 To 8010 Step 230
    Form2.Width = f
    DoEvents
Next
End Sub
Private Sub Command1_Click() '下一個●
Command2.Enabled = True
n = n + 1
Call multi '=-------------------------第幾頁
Call test '=--------------------------偵測頁數
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
Private Sub Command2_Click() '上一個●
Command1.Enabled = True
n = n - 1
Call multi '=-------------------------第幾頁
Call test '=--------------------------偵測頁數
Select Case X '=----------------------選擇頁數
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
Private Sub one() '第一頁的內容●
Label1.Caption = "此版本為 v 4.6.1" & vbCrLf & "修正咒術飭回的攻擊範圍"
End Sub
Private Sub two() '第二頁的內容●
Label1.Caption = "此版本為 v 4.7" & vbCrLf & "1.修正桔梗使用咒術飭回時，磚塊穿    過桔梗不會損血的問題" & vbCrLf & "2.將原本3個陣列合併成一個2維陣列" & vbCrLf & "3.增加了說明的功能" & vbCrLf & "4.增強型的磚塊判斷式= ="
End Sub
Private Sub three() '第三頁的內容●
Label1.Caption = "此版本為 v 5.0" & vbCrLf & "★ 1.使用了大量的動畫效果及修正困      難的難度" & vbCrLf & "★ 2.新增了按鍵自訂功能" & vbCrLf & "    3.修正了1UP一部份出現在視窗的       外面的問題" & vbCrLf & "    4.增加主選單的按鍵說明與按鍵設      定同步功能" & vbCrLf & "★ 5.增加了快按兩下快跑的功能"
End Sub
Private Sub four() '第四頁的內容●
Call changes
Label1.Caption = "此版本為 v 5.0" & vbCrLf & "★ 6.增加大絕招" & vbCrLf & "      用法：" & vbCrLf & "      當四魂之玉能被5整除時，" & vbCrLf & "      照" & eee & fff & ggg & "的順序按。"
End Sub
Private Sub five() '第五頁的內容●
Call changes
Label1.Caption = "此版本為 v 5.1" & vbCrLf & "1.修正按下" & bbb & "產生錯誤的情形" & vbCrLf & "2.修正大絕招的按法：" & eee & fff & ggg & """同時""按下   才有作用" & vbCrLf & "3.修正為四魂之玉為3或其倍數可發動   必殺技"
End Sub
Private Sub six() '第六頁的內容●
Label1.Caption = "此版本為 v 6.0" & vbCrLf & "★ 1.增加第二關和第三關" & vbCrLf & "★ 2.增加打擊分數計算" & vbCrLf & "★ 3.修正40000分可得一顆星" & vbCrLf & "★ 4.修正50000分可得一條命"
End Sub
Private Sub seven() '第七頁的內容●
Label1.Caption = "此版本為 v 6.1" & vbCrLf & "★ 1.增加第四關(小魔王關)"
End Sub
Private Sub eight() '第八頁的內容●
Label1.Caption = "此版本為 v 6.2" & vbCrLf & "1.修正6.1的一些錯誤" & vbCrLf & "2.依照不同難度來增加第四關的能量    球產生數量"
End Sub
Private Sub nine() '第九頁的內容●
Label1.Caption = "此版本為 v 6.3" & vbCrLf & "1.修正6.2的一些錯誤" & vbCrLf & "2.新增簡易的排名"
End Sub
Private Sub ten() '第十頁的內容●
Label1.Caption = "此版本為 v 6.4" & vbCrLf & "★ 1.增加第五關"
End Sub
Private Sub eleven() '第11頁的內容●
Label1.Caption = "此版本為 v 6.5" & vbCrLf & "★ 1.增加全新的關卡(測試中)" & vbCrLf & "    2.加快回血回魔" & vbCrLf & "    3.增加吃四魂之玉可得分數的機制"
End Sub
Private Sub twelve() '第12頁的內容●
Label1.Caption = "此版本為 v 6.6" & vbCrLf & "★ 1.增加打到磚塊的特效(測試中)"
End Sub
Private Sub thirteen() '第13頁的內容●
Label1.Caption = "此版本為 v 7.0(第四代)" & vbCrLf & "★ 1.加入了自動捲動式背景" & vbCrLf & "★ 2.增加打擊特效" & vbCrLf & "★ 3.核心架構全面翻新.."
End Sub
Private Sub multi() '共多少頁●
Label2.Caption = "第" & n & "頁/共" & total & "頁"
End Sub
Private Sub test() '偵測頁數●
X = n Mod 100 '=-----------------偵測頁數
End Sub
Private Sub Form_Unload(Cancel As Integer) '關閉表單●
Form1.Show
End Sub
Private Sub changes() '轉換字元
Call asc(a)
    eee = ss
Call asc(s)
    fff = ss
Call asc(d)
    ggg = ss
Call asc(enter)
    bbb = ss
End Sub
