VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   705
   ClientLeft      =   6435
   ClientTop       =   4800
   ClientWidth     =   3060
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   705
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Tips"
      Height          =   645
      Left            =   2475
      TabIndex        =   6
      Top             =   30
      Width           =   570
   End
   Begin VB.Frame frm1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Drag here to move the window."
      Top             =   30
      Width           =   1350
      Begin VB.HScrollBar f2scr1 
         Height          =   150
         LargeChange     =   5
         Left            =   0
         Max             =   40
         Min             =   255
         SmallChange     =   10
         TabIndex        =   4
         Top             =   0
         Value           =   255
         Width           =   1350
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Caption         =   "alpha %"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   $"SetAlpha.frx":0000
         Top             =   165
         Width           =   840
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "ËÎÌו"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   870
         TabIndex        =   1
         ToolTipText     =   "Click here to exit.You can press the hotkey of <alpha> to show this window again."
         Top             =   165
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   15
         TabIndex        =   3
         Top             =   450
         Width           =   405
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tips:Drag to move the window."
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   1455
      TabIndex        =   5
      Top             =   60
      Width           =   1005
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ckvl1 As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        If ckvl1 = 0 Then
        traymu = False
        ckvl1 = MsgBox("Are you sure to cancel the limit, thus the windows can be set into totally invisible (as dfault the lowest transparency is 40,opaque is 255), and this may cause that you cannot find the window which you set or resume the transparency.", 4132, "Sure to cancel limit?") 'vbYesNo 4 // vbSystemModal 4096 // vbQuestion 32
        traymu = True
            If ckvl1 = 6 Then
            f2scr1.Max = 0
            Else
            Check1.Value = 0
            End If
        Else
        f2scr1.Max = 0
        End If
    Else
    f2scr1.Max = 40
    End If
End Sub

Private Sub Command1_Click()
f2scr1_Change
Unload Me
End Sub


Private Sub Command2_Click()
f2hidetips = True
Command2.Visible = False
Me.Width = 1410
traymu = False
MsgBox "You can leave you mouse on a botton or a label to see detail.", vbInformation, "TIPs.."
traymu = True
End Sub

Private Sub f2scr1_Change()
    SetLayeredWindowAttributes hWnd2, 0, f2scr1.Value, LWA_ALPHA
End Sub


Private Sub Form_Load()
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    If f2hidetips Then
    Command2.Visible = False
    Me.Width = 1410
    End If
Call moveform2
End Sub

Private Sub frm1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    MoveScreen = True
    MousX = x
    MousY = y
    End If
End Sub
Private Sub frm1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveScreen Then
    CurrX = Me.Left - MousX + x
    CurrY = Me.Top - MousY + y
    Me.Move CurrX, CurrY
    End If
End Sub
Private Sub frm1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = False
Call moveform2
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
Call moveform2
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
Call moveform2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    MoveScreen = True
    MousX = x
    MousY = y
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveScreen Then
    CurrX = Me.Left - MousX + x
    CurrY = Me.Top - MousY + y
    Me.Move CurrX, CurrY
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveScreen = False
Call moveform2
End Sub

Private Sub Timer1_Timer()
Call SetForm2Move
End Sub
