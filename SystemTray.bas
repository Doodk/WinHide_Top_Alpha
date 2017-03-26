Attribute VB_Name = "SystemTray"
Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Public Type NOTIFYICONDATA
  cbSize As Long
  Hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type
Public Enum TrayRetunEventEnum
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    Rightup = &H205
    Rightdown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
End Enum

Public traymu As Boolean

