VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "得知CPU頻率的網路上的範例"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'這個程式需要一個Command
Option Explicit
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Const REALTIME_PRIORITY_CLASS = &H100
Private Const THREAD_BASE_PRIORITY_LOWRT = 15
Private Const THREAD_BASE_PRIORITY_MAX = 2
Private Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Private Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long

Private Type LONG_LONG
    HiP As Long
    LoP As Long
End Type

Public opIndex As Long '寫入位置
Dim OpCode() As Byte 'Assembly 的OPCODE

Public Function GetCPUMHZ() As Double
Dim i As Long, CodeStar As Long
ReDim OpCode(600)  '保留用來寫OPCODE
Dim hModule As Long, AddressOfSleep As Long
Dim hProcess As Long, hThread As Long
Dim ProcessPriority As Long, ThreadPriority As Long

'取得Process,Thread的Handle
hProcess = GetCurrentProcess()
hThread = GetCurrentThread()

'取得原來的優先序
ProcessPriority = GetPriorityClass(hProcess)
ThreadPriority = GetThreadPriority(hThread)

'設成最高
SetPriorityClass hProcess, REALTIME_PRIORITY_CLASS
SetThreadPriority hThread, THREAD_BASE_PRIORITY_LOWRT

Dim TimeStamp As LONG_LONG
Dim adrTimeLo As Long, adrTimeHi As Long
adrTimeLo = VarPtr(TimeStamp.HiP)
adrTimeHi = VarPtr(TimeStamp.LoP)

Dim DelyTime As Long
DelyTime = 500 '延遲時間0.5秒鐘

'讀取模組
hModule = LoadLibrary(ByVal "kernel32")
If hModule = 0 Then
    MsgBox "Library讀取失敗"
    Exit Function
End If

'取得Sleep函數位址
AddressOfSleep = GetProcAddress(hModule, ByVal "Sleep")
If AddressOfSleep = 0 Then
   MsgBox "函數讀取失敗", vbCritical
   FreeLibrary hModule
   Exit Function
End If

'起始陣列
    '程式起始位址必須是16的倍數
    CodeStar = (VarPtr(OpCode(0)) Or &HF) + 1
    
    opIndex = CodeStar - VarPtr(OpCode(0)) '程式開始的元素位置
    
    '前端部份以中斷點填滿
    For i = 0 To opIndex - 1
        OpCode(i) = &HCC 'int 3
    Next
'Assembly Code
    'rdtsc (read Time-Stamp counter)
        AddByteToCode &HF
        AddByteToCode &H31
    
    'mov timeLo ,eax
        AddByteToCode &HA3
        AddLongToCode adrTimeLo
    
    'mov TimeHi,edx
        AddByteToCode &H89
        AddByteToCode &H15
        AddLongToCode adrTimeHi
    
    
    'Call Sleep(DelyTime)
        'push DelyTime
            AddByteToCode &H68 'push
            AddLongToCode DelyTime
        'Call Sleep
            AddByteToCode &HE8 'call
            '函數位址 用call的定址
            AddLongToCode AddressOfSleep - VarPtr(OpCode(opIndex)) - 4
        
    'rdtsc
        AddByteToCode &HF
        AddByteToCode &H31
    
    'sub eax, TimerLo
        AddByteToCode &H2B
        AddByteToCode &H5
        AddLongToCode adrTimeLo
    
    'sbb edx, TimerHi
        AddByteToCode &H1B
        AddByteToCode &H15
        AddLongToCode adrTimeHi
    
    'mov TimeLo ,eax
        AddByteToCode &HA3
        AddLongToCode adrTimeLo
    
    'mov TimeHi,edx
        AddByteToCode &H89
        AddByteToCode &H15
        AddLongToCode adrTimeHi
    
    '返回呼叫函數
    'ret 10h
        AddByteToCode &HC2 'ret 10h
        AddByteToCode &H10
        AddByteToCode &H0
'End Assembly Code

'執行剛剛寫完的Assembly Code
    Call CallWindowProc(CodeStar, 0, 1, 2, 3)

FreeLibrary hModule '釋放模組

'還原程序
SetPriorityClass hProcess, ProcessPriority
SetThreadPriority hThread, ThreadPriority

Dim TimeSt As Currency
CopyMemory TimeSt, TimeStamp, 8

'計算速度
GetCPUMHZ = 10# * TimeSt / CDbl(DelyTime)

End Function

'將Long型態的變數寫到OpCode種
Public Sub AddLongToCode(lData As Long)
CopyMemory OpCode(opIndex), lData, 4
opIndex = opIndex + 4
End Sub

'將Byte型態的變數寫到OpCode種
Public Sub AddByteToCode(bData As Byte)
OpCode(opIndex) = bData
opIndex = opIndex + 1
End Sub

Private Sub Form_Load()
clock = Int(GetCPUMHZ)
Unload Me
Form1.Show
End Sub
