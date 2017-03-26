VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Top,Hide,AlphaWindow-By LJL [Topmost]"
   ClientHeight    =   1590
   ClientLeft      =   1560
   ClientTop       =   1980
   ClientWidth     =   3750
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a window then press Ctrl+SHIFT+   to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   7
      Top             =   465
      Width           =   3675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a window then press Ctrl+Alt+         to set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   3675
   End
   Begin VB.Label lb2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3510
      TabIndex        =   9
      Top             =   465
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "set alpha."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   225
      TabIndex        =   8
      Top             =   690
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "Click HERE to minimize it to system tray."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   225
      Left            =   60
      TabIndex        =   5
      ToolTipText     =   "You can resume the window by doubleclick the icon in system tray or right click to choose options."
      Top             =   1350
      Width           =   3645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "hide a window or resume."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   225
      TabIndex        =   4
      Top             =   1125
      Width           =   2190
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "topmost or canel topmost."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   570
      TabIndex        =   2
      Top             =   255
      Width           =   2235
   End
   Begin VB.Label lb1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3210
      TabIndex        =   1
      Top             =   30
      Width           =   105
   End
   Begin VB.Label lb3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3510
      TabIndex        =   0
      Top             =   900
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a window then press Ctrl+SHIFT+  to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   30
      TabIndex        =   6
      Top             =   900
      Width           =   3675
   End
   Begin VB.Menu m1 
      Caption         =   "Meum"
      Visible         =   0   'False
      Begin VB.Menu m1show 
         Caption         =   "Show this tool"
      End
      Begin VB.Menu m1hide 
         Caption         =   "show some window"
         Visible         =   0   'False
      End
      Begin VB.Menu m1about 
         Caption         =   "About"
      End
      Begin VB.Menu m1nil1 
         Caption         =   "-"
      End
      Begin VB.Menu m1exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tray1 As NOTIFYICONDATA
Dim extlop As Integer


Private Sub Form_Load()
If App.PrevInstance Then
MsgBox "This program has already run, click ok to exit.", vbInformation, "Already Load."
End
End If
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Dim lResult As Long
    RegisterHotKeyID1 = GlobalAddAtom("SetForHotKey1")
    RegisterHotKeyID2 = GlobalAddAtom("SetForHotKey2")
    RegisterHotKeyID3 = GlobalAddAtom("SetForHotKey3")
    '--hotkey
    lResult = RegisterHotKey(Form1.Hwnd, RegisterHotKeyID1, MOD_CONTROL Or MOD_ALT, vbKeyT)
        If lResult = 0 Then
            Do
                If extlop = 25 Then
                MsgBox "Ctrl+Alt cannot register hotkey with any letter (A-Z), thus this program cannot run. Press ok to exit.", vbCritical, "ERROR on hotkey for topmost."
                Exit Do
                Unload Me
                End If
            lResult = RegisterHotKey(Form1.Hwnd, RegisterHotKeyID1, MOD_CONTROL Or MOD_ALT, Int(extlop + 65))
            extlop = extlop + 1
            Loop Until lResult <> 0
        MsgBox "Fail to registe hotkey <CTRL+ALT+T>, use <CTRL+ALT+" & Chr(Int(extlop + 64)) & ">as new hotkey for setting topmost!", vbExclamation, "Use <CTRL+ALT+" & Chr(Int(extlop + 64)) & "> as hotkey for topmost!"
        lb1.Caption = Chr(Int(extlop + 64))
        extlop = 0
        End If
    lResult = RegisterHotKey(Form1.Hwnd, RegisterHotKeyID2, MOD_CONTROL Or MOD_SHIFT, vbKeyW)
        If lResult = 0 Then
            Do
                If extlop = 25 Then
               MsgBox "Ctrl+SHIFT cannot registerhotkey with any letter (A-Z), thus this exe cannot run. Press ok to exit.", vbCritical, "ERROR on notkey to set transparency."
                Exit Do
                Unload Me
                End If
            lResult = RegisterHotKey(Form1.Hwnd, RegisterHotKeyID2, MOD_CONTROL Or MOD_SHIFT, Int(extlop + 65))
            extlop = extlop + 1
            Loop Until lResult <> 0
        MsgBox "Registe hotkey <CTRL+SHIFT+Q> fail, use <CTRL+SHIFT+" & Chr(Int(extlop + 64)) & ">as new hotkey to set transparency!", vbExclamation, "Use <CTRL+ALT+" & Chr(Int(extlop + 64)) & "> as hotkey to set transparency!"
        lb2.Caption = Chr(Int(extlop + 64))
        extlop = 0
        End If
    lResult = RegisterHotKey(Form1.Hwnd, RegisterHotKeyID3, MOD_CONTROL Or MOD_SHIFT, vbKeyQ)
        If lResult = 0 Then
            Do
                If extlop = 25 Then
               MsgBox "Ctrl+SHIFT cannot registerhotkey with any letter (A-Z), thus this exe cannot run. Press ok to exit.", vbCritical, "ERROR on notkey for hiding window."
                Exit Do
                Unload Me
                End If
            lResult = RegisterHotKey(Form1.Hwnd, RegisterHotKeyID3, MOD_CONTROL Or MOD_SHIFT, Int(extlop + 65))
            extlop = extlop + 1
            Loop Until lResult <> 0
        MsgBox "Registe hotkey <CTRL+SHIFT+Q> fail, use <CTRL+SHIFT+" & Chr(Int(extlop + 64)) & ">as new hotkey for hiding window!", vbExclamation, "Use <CTRL+ALT+" & Chr(Int(extlop + 64)) & "> as hotkey for hiding window!"
        lb3.Caption = Chr(Int(extlop + 64))
        End If
    PrevWndProc = SetWindowLong(Me.Hwnd, GWL_WNDPROC, AddressOf WndProc)
'-----------------hotleyend
    With tray1
        .cbSize = Len(tray1)
        .Hwnd = Me.Hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.Handle
        .szTip = "Hide & Topmost & transparency  v2.0EN_Alpha" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, tray1
'-----------------trayend
traymu = True
Label6.ToolTipText = "Choose a window which you want to change transparency, then press Ctrl+SHIFT+" & lb2.Caption & ", there will be a new window showd for setting."
Label3.ToolTipText = "Choose a window which you want to hide, then press Ctrl+SHIFT+" & lb3.Caption & ". You can press the same hotkey Ctrl+Shift+" & lb3.Caption & " to resume hidden window anywhere."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If traymu Then
    Dim cEvent As Single
    cEvent = x / Screen.TwipsPerPixelX
    
        Select Case cEvent
        Case LeftDbClick
        m1show_Click
        Case Rightup
        Form1.PopupMenu m1
        End Select
    End If
'-----------------trayend
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If hideyesno Then
    Dim exitexe As Integer
    traymu = False
    exitexe = MsgBox("You have hiden<" & exitname & ">. If exit, the program will automatically resume this window.", vbOKCancel + vbQuestion, "Exit & Resume the window?")
    traymu = True
        If exitexe = 1 Then
        m1hide_Click
        Else
        Cancel = 1
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim lResult As Long
    lResult = SetWindowLong(Me.Hwnd, GWL_WNDPROC, PrevWndProc)
    lResult = UnregisterHotKey(Me.Hwnd, RegisterHotKeyID1)
    lResult = UnregisterHotKey(Me.Hwnd, RegisterHotKeyID2)
    lResult = UnregisterHotKey(Me.Hwnd, RegisterHotKeyID3)
'-----------------hotkeyend
    tray1.uFlags = 0
    Shell_NotifyIcon NIM_DELETE, tray1
'-----------------trayend
Unload Form2
End Sub

Private Sub Label2_Click()
    Form1.Visible = False
End Sub
Private Sub m1about_Click()
    traymu = False
    MsgBox "If you have any questions or suggestions, " & Chr(10) & "please connect: ljl0722@hotmail.com", vbInformation, "About... - v1.2EN (Beta)"
    traymu = True
End Sub
Private Sub m1exit_Click()
    Unload Me
End Sub
Private Sub m1hide_Click()
    hide1 = ShowWindow(hidehwnd1, SW_SHOW)
    If hide1 = 0 Then
    hideyesno = False
    m1hide.Visible = False
    End If
End Sub
Private Sub m1show_Click()
    Form1.Visible = True
    Form1.WindowState = 0
    Form1.Show
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    MoveScreen = True
    MousX = x
    MousY = y
    End If
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveScreen Then
    CurrX = Me.Left - MousX + x
    CurrY = Me.Top - MousY + y
    Me.Move CurrX, CurrY
    End If
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = False
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    MoveScreen = True
    MousX = x
    MousY = y
    End If
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveScreen Then
    CurrX = Me.Left - MousX + x
    CurrY = Me.Top - MousY + y
    Me.Move CurrX, CurrY
    End If
End Sub
Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = False
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    MoveScreen = True
    MousX = x
    MousY = y
    End If
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveScreen Then
    CurrX = Me.Left - MousX + x
    CurrY = Me.Top - MousY + y
    Me.Move CurrX, CurrY
    End If
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = False
End Sub
