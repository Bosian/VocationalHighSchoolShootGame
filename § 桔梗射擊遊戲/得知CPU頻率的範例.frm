VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "�o��CPU�W�v�������W���d��"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�o�ӵ{���ݭn�@��Command
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

Public opIndex As Long '�g�J��m
Dim OpCode() As Byte 'Assembly ��OPCODE

Public Function GetCPUMHZ() As Double
Dim i As Long, CodeStar As Long
ReDim OpCode(600)  '�O�d�ΨӼgOPCODE
Dim hModule As Long, AddressOfSleep As Long
Dim hProcess As Long, hThread As Long
Dim ProcessPriority As Long, ThreadPriority As Long

'���oProcess,Thread��Handle
hProcess = GetCurrentProcess()
hThread = GetCurrentThread()

'���o��Ӫ��u����
ProcessPriority = GetPriorityClass(hProcess)
ThreadPriority = GetThreadPriority(hThread)

'�]���̰�
SetPriorityClass hProcess, REALTIME_PRIORITY_CLASS
SetThreadPriority hThread, THREAD_BASE_PRIORITY_LOWRT

Dim TimeStamp As LONG_LONG
Dim adrTimeLo As Long, adrTimeHi As Long
adrTimeLo = VarPtr(TimeStamp.HiP)
adrTimeHi = VarPtr(TimeStamp.LoP)

Dim DelyTime As Long
DelyTime = 500 '����ɶ�0.5����

'Ū���Ҳ�
hModule = LoadLibrary(ByVal "kernel32")
If hModule = 0 Then
    MsgBox "LibraryŪ������"
    Exit Function
End If

'���oSleep��Ʀ�}
AddressOfSleep = GetProcAddress(hModule, ByVal "Sleep")
If AddressOfSleep = 0 Then
   MsgBox "���Ū������", vbCritical
   FreeLibrary hModule
   Exit Function
End If

'�_�l�}�C
    '�{���_�l��}�����O16������
    CodeStar = (VarPtr(OpCode(0)) Or &HF) + 1
    
    opIndex = CodeStar - VarPtr(OpCode(0)) '�{���}�l��������m
    
    '�e�ݳ����H���_�I��
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
            '��Ʀ�} ��call���w�}
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
    
    '��^�I�s���
    'ret 10h
        AddByteToCode &HC2 'ret 10h
        AddByteToCode &H10
        AddByteToCode &H0
'End Assembly Code

'������g����Assembly Code
    Call CallWindowProc(CodeStar, 0, 1, 2, 3)

FreeLibrary hModule '����Ҳ�

'�٭�{��
SetPriorityClass hProcess, ProcessPriority
SetThreadPriority hThread, ThreadPriority

Dim TimeSt As Currency
CopyMemory TimeSt, TimeStamp, 8

'�p��t��
GetCPUMHZ = 10# * TimeSt / CDbl(DelyTime)

End Function

'�NLong���A���ܼƼg��OpCode��
Public Sub AddLongToCode(lData As Long)
CopyMemory OpCode(opIndex), lData, 4
opIndex = opIndex + 4
End Sub

'�NByte���A���ܼƼg��OpCode��
Public Sub AddByteToCode(bData As Byte)
OpCode(opIndex) = bData
opIndex = opIndex + 1
End Sub

Private Sub Form_Load()
clock = Int(GetCPUMHZ)
Unload Me
Form1.Show
End Sub
