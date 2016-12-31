Attribute VB_Name = "Module1"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Key_up As Integer '宣告跨表單的鍵盤按鍵變數---↓
Public key_down As Integer
Public key_left As Integer
Public key_right As Integer
Public a As Integer
Public s As Integer
Public d As Integer
Public enter As Integer '宣告跨表單的鍵盤按鍵變數---↑
Public ss As String '還原成字元的暫存器
Public win As Integer '為1時,勝利或掛掉的名次之介面
Public mark '分數
Public ppi As Integer '決定順位(空白之分數)
Public Y(2) As Integer '=----------(0)磚塊的血量(1)磚塊損血亂數之最小值(2)磚塊損血亂數之最大值
Public ssheight '讓form5與form1一樣
Public sswidth '讓form5與form1一樣
Public sstop '同上
Public ssleft '同上
Public clock 'cpu的頻率
Public Function coll(a As Object, b As Object) As Boolean '碰撞涵數●
If b.Left + b.Width > a.Left And a.Left + a.Width > b.Left And _
   b.Top + b.Height > a.Top And a.Top + a.Height > b.Top Then coll = True

'If (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Or _
   (a.Left >= b.Left And a.Left <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top + a.Height >= b.Top And a.Top + a.Height <= b.Top + b.Height) Or _
   (a.Left + a.Width >= b.Left And a.Left + a.Width <= b.Left + b.Width And _
   a.Top >= b.Top And a.Top <= b.Top + b.Height) Then coll = True
End Function
Public Function fure1(a As Object) As Boolean '上邊界●
If a.Top < 0 Then fure1 = True
End Function
Public Function fure2(a As Object) As Boolean '下邊界●
If a.Top > Form1.ScaleHeight - a.Height Then fure2 = True
End Function
Public Function fure3(a As Object) As Boolean '右邊界●
If a.Left + a.Width > Form1.ScaleWidth Then fure3 = True
End Function
Public Function fure4(a As Object) As Boolean '左邊界●
If a.Left < 0 Then fure4 = True
End Function
Public Sub asc(a) 'ASCII還原成字元●
Select Case a
    Case 9
        ss = "Tab"
    Case 13
        ss = "Enter"
    Case 16
        ss = "Shift"
    Case 17
        ss = "Ctrl"
    Case 18
        ss = "Alt"
    Case 32
        ss = "空白鍵"
    Case 37
        ss = "←"
    Case 38
        ss = "↑"
    Case 39
        ss = "→"
    Case 40
        ss = "↓"
    Case 65
        ss = "A"
    Case 66
        ss = "B"
    Case 67
        ss = "C"
    Case 68
        ss = "D"
    Case 69
        ss = "E"
    Case 70
        ss = "F"
    Case 71
        ss = "G"
    Case 72
        ss = "H"
    Case 73
        ss = "I"
    Case 74
        ss = "J"
    Case 75
        ss = "K"
    Case 76
        ss = "L"
    Case 77
        ss = "M"
    Case 78
        ss = "N"
    Case 79
        ss = "O"
    Case 80
        ss = "P"
    Case 81
        ss = "Q"
    Case 82
        ss = "R"
    Case 83
        ss = "S"
    Case 84
        ss = "T"
    Case 85
        ss = "U"
    Case 86
        ss = "V"
    Case 87
        ss = "W"
    Case 88
        ss = "X"
    Case 89
        ss = "Y"
    Case 90
        ss = "Z"
    Case 186
        ss = ";"
    Case 187
        ss = "="
    Case 188
        ss = ","
    Case 189
        ss = "-"
    Case 190
        ss = "."
    Case 191
        ss = "/"
    Case 192
        ss = "~"
    Case 219
        ss = "["
    Case 220
        ss = "\"
    Case 221
        ss = "]"
    Case 222
        ss = """"
Case Else
    ss = "無"
End Select
End Sub

