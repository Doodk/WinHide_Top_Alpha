Attribute VB_Name = "Customize"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOPMOST = &H8
'alphastart
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000
'alphaend
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, wndRect As RECT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'^is GetTaskbarHeight
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0

Public hideyesno As Boolean
Public hWnd1 As Long, hWnd2 As Long, hidehwnd1 As Long, exitname As String, RECTfanhui As Long
Public hide1 As Long, f2hidetips As Boolean, ActiveWindowRECT As RECT
Public MoveScreen As Boolean, MousX As Integer, MousY As Integer, CurrX As Integer, CurrY As Integer
Private oldhwnd1 As Long, oldhwnd2 As Long, alphaback As Long


Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = Screen.Height - rectVal.Bottom * Screen.TwipsPerPixelY
End Function

Public Sub moveform2()
    If Form2.Left < 0 Then Form2.Left = 0
    If Form2.Top < 0 Then Form2.Top = 0
    If Form2.Top > Screen.Height - GetTaskbarHeight - Form2.Height Then _
    Form2.Top = Screen.Height - GetTaskbarHeight - Form2.Height
    If Form2.Left > Screen.Width - Form2.Width Then _
    Form2.Left = Screen.Width - Form2.Width
End Sub

Public Sub SetFormAlpha()
    hWnd2 = GetForegroundWindow()
    Do
        If Form2.Hwnd <> hWnd2 Then
        oldhwnd2 = hWnd2
        Call SetForm2Move
            If ActiveWindowRECT.Top * Screen.TwipsPerPixelY < Form2.Height Then
            Form2.Top = 600
            ElseIf ActiveWindowRECT.Top * Screen.TwipsPerPixelY >= Form2.Height Then
            Form2.Top = ActiveWindowRECT.Top * Screen.TwipsPerPixelY - Form2.Height
            End If
        Form2.Left = ActiveWindowRECT.Left * Screen.TwipsPerPixelX
        Form2.Show
        Form2.Timer1.Enabled = True
        Call moveform2
        Dim TmpInfo As Long
        TmpInfo = GetWindowLong(hWnd2, GWL_EXSTYLE)
        alphaback = SetWindowLong(hWnd2, GWL_EXSTYLE, TmpInfo Or WS_EX_LAYERED)
        If alphaback = 0 Then
        traymu = False
        MsgBox "This window cannot set transparency.", vbCritical, "Fail£¡"
        traymu = True
        Else
        Form2.Label3.Caption = "Window:" & getCaption(hWnd2)
        Form2.Label3.ToolTipText = "You are going to change window<" & getCaption(hWnd2) & ">'s transparency. Drag here to move the tool."
        SetWindowPos Form2.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        End If
        Exit Do
        Else
        hWnd2 = oldhwnd2
        End If
    Loop
End Sub
Public Sub SetForm2Move()
    RECTfanhui = GetWindowRect(hWnd2, ActiveWindowRECT)
    If RECTfanhui = 0 Then Unload Form2
End Sub


Public Sub hidewin()
    hWnd1 = GetForegroundWindow()
    If Form2.Hwnd = hWnd1 Then
    hWnd1 = hWnd2
    End If
    If hideyesno Then
    hide1 = ShowWindow(hidehwnd1, SW_SHOW)
        If hide1 = 0 Then
        hideyesno = False
        Form1.m1hide.Visible = False
        End If
    Else
    hide1 = ShowWindow(hWnd1, SW_HIDE)
        If hide1 <> 0 Then
        hideyesno = True
        hidehwnd1 = hWnd1
        Form1.m1hide.Visible = True
        exitname = getCaption(hWnd1)
        Form1.m1hide.Caption = "Show<" & exitname & ">"
        End If
    End If
End Sub

Public Sub setwintop()
    hWnd1 = GetForegroundWindow()
    If Form2.Hwnd = hWnd1 Then
    hWnd1 = hWnd2
    End If
    Dim top1 As Long
    If IsTopmost(hWnd1) Then
    top1 = SetWindowPos(hWnd1, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
        If top1 = 0 Then
        traymu = False
        MsgBox "Fail to Cancel this window's Topmost!", vbCritical, "Fail!"
        traymu = True
        Else
        SetWindowText hWnd1, Replace(getCaption(hWnd1), " [Topmost]", "")
        traymu = False
        'MsgBox "Successfully Cancel window<" & getCaption(hWnd1) & ">'s Topmost!", vbInformation + vbSystemModal, "Cancel Topmost!"
        traymu = True
        End If
    Else
    top1 = SetWindowPos(hWnd1, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
        If top1 = 0 Then
        traymu = False
        MsgBox "Fail to Set this window's Topmost!", vbCritical, "Fail!"
        traymu = True
        Else
        traymu = False
        'MsgBox "Successfully Set window<" & getCaption(hWnd1) & ">'s Topmost!", vbInformation + vbSystemModal, "Set Topmost!"
        traymu = True
        SetWindowText hWnd1, Replace(getCaption(hWnd1), " [Topmost]", "") & " [Topmost]"
        End If
    End If
End Sub

Function getCaption(Hwnd As Long)
    Dim hWndlength As Long, hWndTitle As String, A As Long
    hWndlength = GetWindowTextLength(Hwnd)
    hWndTitle = String$(hWndlength, 0)
    A = GetWindowText(Hwnd, hWndTitle, (hWndlength + 1))
    getCaption = hWndTitle
End Function

Function IsTopmost(Hwnd As Long) As Boolean
    Dim Ret As Long, t1 As Long, t2 As Long
    Ret = GetWindowLong(Hwnd, GWL_EXSTYLE)
    t1 = Ret Or WS_EX_TOPMOST
    IsTopmost = (Ret = t1)
End Function

