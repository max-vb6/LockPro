VERSION 5.00
Begin VB.Form frmLock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   ControlBox      =   0   'False
   Icon            =   "frmLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picScr 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      MouseIcon       =   "frmLock.frx":000C
      MousePointer    =   99  'Custom
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer tmrScr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1920
   End
   Begin LockPro.ucKeyboard Keyboard 
      Height          =   4380
      Left            =   9840
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7726
   End
   Begin VB.Timer tmrSKb 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1440
   End
   Begin VB.PictureBox picKeyPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrSMe 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox picUSB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      ScaleHeight     =   495
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   7560
      Width           =   5295
      Begin VB.Label lblUSB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "请插入指定的USB设备，然后点击我"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   5295
      End
   End
   Begin VB.TextBox txtPsw 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   26.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "l"
      TabIndex        =   1
      Text            =   "s"
      Top             =   7560
      Width           =   4935
   End
   Begin VB.Timer tmrTop 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picKeyboard 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2160
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   7560
      Width           =   495
      Begin LockPro.PngImage pngKeyboard 
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   926
         Image           =   "frmLock.frx":0156
         Opacity         =   70
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin VB.PictureBox picShowPsw 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   7560
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   7560
      Width           =   495
      Begin LockPro.PngImage pngShowPsw 
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   926
         Image           =   "frmLock.frx":0896
         Opacity         =   70
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   5
      Top             =   8040
      Width           =   1065
   End
   Begin LockPro.PngImage pngShut 
      Height          =   1080
      Left            =   9360
      Top             =   7680
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   1905
      Image           =   "frmLock.frx":0E08
      Opacity         =   50
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngUnlock 
      Height          =   3975
      Left            =   5280
      Top             =   -120
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   7011
      Image           =   "frmLock.frx":1C6C
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "如需解锁，请输入密码"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   6600
      Width           =   5895
   End
   Begin LockPro.PngImage pngShdw 
      Height          =   1485
      Left            =   1680
      Top             =   7080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2619
      Image           =   "frmLock.frx":901A
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "您的计算机已被锁定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   10620
   End
   Begin LockPro.PngImage pngLkPr 
      Height          =   3840
      Left            =   1440
      Top             =   0
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   6773
      Image           =   "frmLock.frx":A578
   End
   Begin LockPro.PngImage pngWhite 
      Height          =   1260
      Index           =   0
      Left            =   -240
      Top             =   4440
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   2223
      Image           =   "frmLock.frx":11A4D
      Scaler          =   1
      Opacity         =   80
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngArrow 
      Height          =   960
      Left            =   4560
      Top             =   6360
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "frmLock.frx":1B1BF
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngWhite 
      Height          =   780
      Index           =   2
      Left            =   -240
      Top             =   8040
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1376
      Image           =   "frmLock.frx":1BB20
      Scaler          =   1
      Opacity         =   70
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngWhite 
      Height          =   780
      Index           =   1
      Left            =   3000
      Top             =   6480
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   1376
      Image           =   "frmLock.frx":25292
      Scaler          =   1
      Opacity         =   70
      OLEdrop         =   1
      Props           =   5
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PswErr As Long, PswLarge As Long, Waited As Long, sPsw As String
Dim lBtm As Long, lHookWin As Long
Dim ScrPos As POINTAPI, LastPos As POINTAPI, ScrCnt As Long, ScrWait As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    If picScr.Visible Then
        picScr.Visible = False
    Else
        ScrCnt = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo OpenErr
    
    ShowWindow FindWindow(vbNullString, "“开始”菜单"), 0    '隐藏开始菜单，防止 Win8/8.1 非桌面模式
    mouse_event &H2, 0, 0, 0, 0                             '模拟鼠标点击，转移焦点
    mouse_event &H4, 0, 0, 0, 0
    KillTaskMgr                                             '杀任管
    
    lHookWin = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf HookKeyboard, App.hInstance, 0)
    
    PswErr = 0
    PswLarge = CLng(ReadCon("PswLarge"))
    sPsw = ReadPsw("Psw")
    lblShow(0).Caption = ReadCon("Txt")
    
    With Me
        'If .Tag = "" Then .Move 0, 0
        '.Width = Screen.Width
        '.Height = Screen.Height
        .Move 0, 0, Screen.Width, Screen.Height
        picScr.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        pngShut.Enabled = True
        pngShut.Opacity = 50
        .PaintPicture LoadPicture(MyPath & "LockPicture\" & ReadCon("BGPic")), 0, 0, .ScaleWidth, .ScaleHeight
        Keyboard.Move 0, .ScaleHeight - Keyboard.Height, .ScaleWidth
        picKeyPic.Move 0, 0, .ScaleWidth, Keyboard.Height
        picKeyPic.PaintPicture .Image, 0, 0, .ScaleWidth, Keyboard.Height, 0, .ScaleHeight - Keyboard.Height, .ScaleWidth, Keyboard.Height
        Keyboard.SetKeyPic picKeyPic.Image          '截取键盘背景图像
        picKeyPic.Cls
    End With
    
    ResetLockControlTop
    picScr.ZOrder 0
    
    If ReadCon("Scr") = 1 Then
        ScrWait = CLng(ReadCon("ScrWait"))
        tmrScr.Enabled = True
    Else
        tmrScr.Enabled = False
    End If
    
    If ReadCon("Psw") = 1 Then
        lblShow(1).Caption = ""
        txtPsw.Enabled = False
        pngWhite(1).Visible = False
    Else
        picUSB.Visible = False
        txtPsw.Enabled = True
        pngWhite(1).Visible = True
    End If
    txtPsw.Text = ""
    tmrTop.Tag = ""
    Locked = True
    
    Exit Sub
OpenErr:
    Me.Hide
    MsgBox "由于配置文件错误，锁定界面无法启动", 48, "重大错误！"
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindowsHookEx lHookWin
    Locked = False
    'Me.Tag = ""
End Sub

Private Sub Keyboard_InputFinish()
    If txtPsw.Enabled = False Then Beep: Exit Sub
    txtPsw_KeyPress 13
End Sub

Private Sub Keyboard_KeyPressed(sKey As String)
    If txtPsw.Enabled = False Then Beep: Exit Sub
    With txtPsw
        If .SelText = "" Then
            .Text = .Text + sKey
        Else
            .SelText = sKey
        End If
    End With
End Sub

Private Sub Keyboard_LetterBacked()
    If txtPsw.Enabled = False Then Beep: Exit Sub
    With txtPsw
        If .Text = "" Then Beep: Exit Sub
        If .SelText = "" Then
            .Text = Left(.Text, Len(.Text) - 1)
        Else
            .SelText = ""
        End If
    End With
End Sub

Private Sub lblUSB_Click()
    DoEvents
    If CheckUSB = 1 Then
        Pass
    Else
        PswErr = PswErr + 1
        lblUSB.ForeColor = vbRed
        If PswLarge - PswErr > 0 Then
            lblUSB.Caption = "验证错误，还有 " & PswLarge - PswErr & " 次机会"
        Else
            lblUSB.Enabled = False
            If ReadCon("PswErr") = 0 Then
                Waited = CLng(ReadCon("PswWait"))
                tmrWait_Timer
                tmrWait.Enabled = True
            ElseIf ReadCon("PswErr") = 1 Then
                lblUSB.Caption = "解锁失败，计算机即将关闭"
                'Shell "shutdown.exe -s -t 0", vbHide
                NtShutdown
            Else
                lblUSB.Caption = "解锁失败，计算机即将重新启动"
                'Shell "shutdown.exe -r -t 0", vbHide
                NtShutdown 1
            End If
        End If
    End If
End Sub

Private Sub picScr_Click()
    picScr.Visible = False
End Sub

Private Sub pngArrow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrSMe.Enabled = False
End Sub

Private Sub pngArrow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Static ox!
    With Me
        If Button = 1 Then
            .Move .Left - ox + x
        Else
            ox = x
        End If
    End With
End Sub

Private Sub pngArrow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrSMe.Enabled = True
End Sub

Private Sub pngKeyboard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngKeyboard.FadeInOut 100, 2
End Sub

Private Sub pngKeyboard_MouseEnter()
    pngKeyboard.FadeInOut 80, 2
End Sub

Private Sub pngKeyboard_MouseExit()
    pngKeyboard.FadeInOut 70, 2
End Sub

Private Sub pngKeyboard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Keyboard.Visible = False Then
        ShowKeyboard
    Else
        CloseKeyboard
    End If
    pngKeyboard.FadeInOut 80, 2
End Sub

Private Sub pngShowPsw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sFont As String
    sFont = IIf(Dir(Environ("Windir") & "\Fonts\msyh.ttf") <> "", "微软雅黑", "Arial")
    With txtPsw
        .PasswordChar = ""
        .Font.Name = sFont
        .Font.Size = "22"
    End With
    pngShowPsw.FadeInOut 100, 2
End Sub

Private Sub pngShowPsw_MouseEnter()
    pngShowPsw.FadeInOut 80, 2
End Sub

Private Sub pngShowPsw_MouseExit()
    pngShowPsw.FadeInOut 70, 2
End Sub

Private Sub pngShowPsw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    With txtPsw
        .PasswordChar = "l"
        .Font.Name = "Wingdings"
        .Font.Size = "26"
        .SetFocus
    End With
    pngShowPsw.FadeInOut 80, 2
End Sub

Private Sub pngShut_DblClick(ByVal Button As Integer)
    'Shell "shutdown.exe -s -t 0", vbHide
    NtShutdown
End Sub

Private Sub pngShut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngShut.FadeInOut 100, 5
End Sub

Private Sub pngShut_MouseEnter()
    pngShut.FadeInOut 80, 5
    tmrTop.Tag = "Shutdown"
End Sub

Private Sub pngShut_MouseExit()
    pngShut.FadeInOut 50, 5
    tmrTop.Tag = ""
End Sub

Private Sub pngShut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngShut.FadeInOut 50, 5
End Sub

Private Sub tmrScr_Timer()
    GetCursorPos ScrPos
    If Abs(LastPos.x - ScrPos.x) > 5 Or Abs(LastPos.y - ScrPos.y) > 5 Then
        ScrCnt = 0
        If picScr.Visible Then picScr.Visible = False
    ElseIf Not picScr.Visible Then
        ScrCnt = ScrCnt + 1
        If ScrCnt >= ScrWait Then
            picScr.Visible = True
            ScrCnt = 0
        End If
    End If
    LastPos = ScrPos
End Sub

Private Sub tmrSKb_Timer()
    If tmrSKb.Tag = "" Then
        SetLockControlTop GetMoveNum(Me.ScaleHeight - Keyboard.Height - lBtm, pngShdw.Top, 3)
        If GetMoveNum(Me.ScaleHeight - Keyboard.Height - lBtm, pngShdw.Top, 3) = 0 Then
            Keyboard.Visible = True
            Keyboard.ShowBd
            pngWhite(0).Visible = True
            pngWhite(1).Visible = True
            tmrSKb.Enabled = False
        End If
    Else
        SetLockControlTop GetMoveNum(Me.ScaleHeight - lBtm, pngShdw.Top, 3)
        If GetMoveNum(Me.ScaleHeight - lBtm, pngShdw.Top, 3) = 0 Then
            pngWhite(0).Visible = True
            pngWhite(1).Visible = True
            tmrSKb.Enabled = False
        End If
    End If
End Sub

Private Sub tmrSMe_Timer()
    If Me.Left > Screen.Width / 4 Then
        Me.Left = Me.Left + GetMoveNum(Screen.Width, Me.Left, 10)
        If GetMoveNum(Screen.Width, Me.Left, 10) = 0 Then Me.Left = Screen.Width: Unload Me
    Else
        Me.Left = Me.Left + GetMoveNum(0, Me.Left, 10)
        If GetMoveNum(0, Me.Left, 10) = 0 Then Me.Left = 0: tmrSMe.Enabled = False
    End If
End Sub

Private Sub tmrTop_Timer()
    If tmrTop.Tag <> "Pass" Then FormOnTop Me, True: Me.ZOrder 0
    Me.SetFocus
    If tmrTop.Tag <> "Shutdown" Then
        lblTime.Caption = Format(time, "hh:mm")
    Else
        lblTime.Caption = "双击按钮来关机"
    End If
    pngWhite(2).Width = lblTime.Width + 720
End Sub

Private Sub tmrWait_Timer()
    If picUSB.Visible = False Then
        lblShow(1).Caption = "你可以在 " & Waited & " 秒后重新输入密码"
        Waited = Waited - 1
        If Waited < 0 Then
            lblShow(1).ForeColor = vbBlack
            lblShow(1).Caption = "如需解锁，请输入密码"
            txtPsw.Enabled = True
            txtPsw.SetFocus
            PswErr = 0
            tmrWait.Enabled = False
        End If
    Else
        lblUSB.Caption = "你可以在 " & Waited & " 秒后重新进行验证"
        Waited = Waited - 1
        If Waited < 0 Then
            lblUSB.ForeColor = vbBlack
            lblUSB.Caption = "请插入指定的USB设备，然后点击我"
            lblUSB.Enabled = True
            PswErr = 0
            tmrWait.Enabled = False
        End If
    End If
End Sub

Private Sub txtPsw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If XMD5(txtPsw.Text) = sPsw Then
            If ReadCon("Psw") = 2 Then
                picUSB.Visible = True
                lblShow(1).Visible = False
                CloseKeyboard
                pngWhite(1).Left = Me.ScaleWidth
                Exit Sub
            End If
            Pass
        Else
            PswErr = PswErr + 1
            lblShow(1).ForeColor = vbRed
            txtPsw.Text = ""
            If PswLarge - PswErr > 0 Then
                lblShow(1).Caption = "密码错误，还有 " & PswLarge - PswErr & " 次机会"
            Else
                txtPsw.Enabled = False
                If ReadCon("PswErr") = 0 Then
                    Waited = CLng(ReadCon("PswWait"))
                    tmrWait_Timer
                    tmrWait.Enabled = True
                ElseIf ReadCon("PswErr") = 1 Then
                    lblShow(1).Caption = "解锁失败，计算机即将关闭"
                    'Shell "shutdown.exe -s -t 0", vbHide
                    NtShutdown
                Else
                    lblShow(1).Caption = "解锁失败，计算机即将重新启动"
                    'Shell "shutdown.exe -r -t 0", vbHide
                    NtShutdown 1
                End If
            End If
        End If
    End If
End Sub

Sub Pass()
    DoEvents
    PlaySoundResource 101
    DoEvents
    CloseKeyboard
    pngLkPr.Visible = False
    pngUnlock.Visible = True
    lblShow(0).Caption = "请向右拖动箭头"
    lblShow(1).Visible = False
    pngShdw.Visible = False
    txtPsw.Visible = False
    picShowPsw.Visible = False
    picKeyboard.Visible = False
    picUSB.Visible = False
    pngArrow.Visible = True
    pngWhite(1).Visible = True
    pngWhite(1).Move pngArrow.Left, pngShdw.Top, pngArrow.Width, pngArrow.Height
    pngShut.Enabled = False
    pngShut.FadeInOut 0, 5
    tmrTop.Tag = "Pass"
End Sub

Sub ShowKeyboard()
    pngWhite(0).Visible = False
    pngWhite(1).Visible = False
    tmrSKb.Tag = ""
    tmrSKb.Enabled = True
End Sub

Sub CloseKeyboard()
    pngWhite(0).Visible = False
    pngWhite(1).Visible = False
    Keyboard.HideBd
    Keyboard.Visible = False
    tmrSKb.Tag = "Cl"
    tmrSKb.Enabled = True
End Sub

Sub SetLockControlTop(lOffset As Long)
    pngLkPr.Top = pngLkPr.Top + lOffset
    pngUnlock.Top = pngUnlock.Top + lOffset
    lblShow(0).Top = lblShow(0).Top + lOffset
    pngWhite(0).Top = pngWhite(0).Top + lOffset
    lblShow(1).Top = lblShow(1).Top + lOffset
    pngWhite(1).Top = pngWhite(1).Top + lOffset
    pngShdw.Top = pngShdw.Top + lOffset
    txtPsw.Top = txtPsw.Top + lOffset
    picKeyboard.Top = picKeyboard.Top + lOffset
    picShowPsw.Top = picShowPsw.Top + lOffset
    picUSB.Top = picUSB.Top + lOffset
End Sub

Sub ResetLockControlTop()
    Dim lHgt As Long
    With Me
        lHgt = (.ScaleHeight - 8565) / 2                          '8565: 锁定区域控件总高度
        pngLkPr.Move (.ScaleWidth - pngLkPr.Width) / 2, lHgt
        pngUnlock.Move pngLkPr.Left, pngLkPr.Top - 120
        lblShow(0).Move (.ScaleWidth - lblShow(0).Width) / 2, lblShow(0).Top + lHgt
        pngWhite(0).Move (.ScaleWidth - pngWhite(0).Width) / 2, lblShow(0).Top - 120
        lblShow(1).Move (.ScaleWidth - lblShow(1).Width) / 2, lblShow(1).Top + lHgt
        pngWhite(1).Move (.ScaleWidth - pngWhite(1).Width) / 2, lblShow(1).Top - 120
        pngShdw.Move (.ScaleWidth - pngShdw.Width) / 2, pngShdw.Top + lHgt
        lBtm = .ScaleHeight - pngShdw.Top
        pngArrow.Move (.ScaleWidth - pngArrow.Width) / 2, pngShdw.Top
        txtPsw.Move pngShdw.Left + 480 + picKeyboard.Width, pngShdw.Top + 480, pngShdw.Width - 1024 - pngShowPsw.Width * 2
        picKeyboard.Move pngShdw.Left + 480, txtPsw.Top, pngKeyboard.Width, txtPsw.Height
        picShowPsw.Move txtPsw.Left + txtPsw.Width, txtPsw.Top, pngShowPsw.Width, picKeyboard.Height
        picUSB.Move picKeyboard.Left, txtPsw.Top, txtPsw.Width + picKeyboard.Width + picShowPsw.Width, txtPsw.Height
        picUSB.ZOrder 0
        lblUSB.Move 0, (picUSB.Height - lblUSB.Height) / 2, picUSB.ScaleWidth
        pngShut.Move .ScaleWidth - pngShut.Width - 600, .ScaleHeight - pngShut.Height - 600
        lblTime.Move 600, .ScaleHeight - lblTime.Height - 600
        pngWhite(2).Move lblTime.Left - 360, lblTime.Top
    End With
End Sub
