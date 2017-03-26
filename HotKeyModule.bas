Attribute VB_Name = "HotKeyModule"
Option Explicit
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal id As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_HOTKEY = &H312
Public PrevWndProc As Long
    Public RegisterHotKeyID1 As Long, RegisterHotKeyID2 As Long, RegisterHotKeyID3 As Long

Function WndProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If uMsg = WM_HOTKEY Then
        If wParam = RegisterHotKeyID1 Then
        Call setwintop
        ElseIf wParam = RegisterHotKeyID2 Then
        Call SetFormAlpha
        ElseIf wParam = RegisterHotKeyID3 Then
        Call hidewin
        End If
Else
        WndProc = CallWindowProc(PrevWndProc, Hwnd, uMsg, wParam, lParam)
End If
End Function
