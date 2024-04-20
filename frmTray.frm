VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmTray 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   645
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   2235
   ControlBox      =   0   'False
   FillColor       =   &H80000012&
   Icon            =   "frmTray.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "LockPro"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock sckHtp 
      Index           =   0
      Left            =   2040
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Timer tmrCntDwn 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   480
   End
   Begin LockPro.Downloader Updater 
      Left            =   1440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer tmrSUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrSMe 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.PictureBox picDDE 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrKey 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblMore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1905
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      ToolTipText     =   "忽略更新"
      Top             =   720
      Width           =   225
   End
   Begin VB.Label lblUpd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发现新版本了哦~"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   165
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      ToolTipText     =   "点击查看更新"
      Top             =   720
      Width           =   1395
   End
   Begin LockPro.PngImage pngMenu 
      Height          =   645
      Left            =   1800
      ToolTipText     =   "菜单"
      Top             =   0
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   1138
      Image           =   "frmTray.frx":4781A
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngLock 
      Height          =   645
      Left            =   0
      ToolTipText     =   "双击立即锁定计算机"
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1138
      Image           =   "frmTray.frx":47CC9
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuLock 
         Caption         =   "立即锁定计算机"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "打开使用向导"
      End
      Begin VB.Menu mnuTimer 
         Caption         =   "定时器"
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Lock Pro 设置"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于 Lock Pro"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Shdw As cShadow
Dim lpPoint As POINTAPI

Sub CloseLP()
    If ReadCon("UNR") = 1 Then RtlSetProcessIsCritical 0, 0, 0
    SaveCon "FormLeft", CStr(Me.Left)
    UnHook Me.hwnd
    Set Shdw = Nothing
    Dim frms As Form
    For Each frms In VB.Forms
        Unload frms
    Next
End Sub

Sub SetSocket()
    On Error GoTo SSErr
    
    Dim i As Long
    If ReadCon("Remolock") = "1" Then
        ResetAllSocket
        With sckHtp(0)
            .LocalPort = CLng(ReadCon("Port"))
            .Protocol = sckTCPProtocol
            .Listen
        End With
    Else
        ResetAllSocket
    End If
    
    Exit Sub
SSErr:
    MsgBox "初始化 Remolock 远程服务时出现错误 " & Err.Number & "！" & vbCrLf & Err.Description, 48, "初始化 Remolock 错误"
    SaveCon "Remolock", "0"
End Sub

Sub ResetAllSocket()
    On Error GoTo rsErr
    
    Dim i As Long
    sckHtp(0).Close
    If sckHtp.Count > 1 Then
        For i = 1 To sckHtp.UBound
            sckHtp(i).Close
            Unload sckHtp(i)
        Next i
    End If
    
    Exit Sub
rsErr:
    If Err.Number = 340 Then
        i = i + 1
        Resume
    End If
End Sub

Private Sub sckHtp_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Load sckHtp(requestID)
    With sckHtp(requestID)
        .Tag = 0
        .Accept requestID
    End With
End Sub

Private Sub sckHtp_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim sTmp As String
    sckHtp(Index).GetData sTmp, vbString
    HTTPRespond sckHtp(Index), sTmp, "Remolock"
End Sub

Private Sub sckHtp_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
End Sub

Private Sub sckHtp_SendComplete(Index As Integer)
    HTTPSendCheck sckHtp(Index)
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    Select Case LCase(CmdStr)
        Case "lock"
            frmLock.Show
    End Select
    Cancel = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    If App.PrevInstance Then
        If ReadCon("First") = 0 Then
            On Error Resume Next
            Dim t As Long
            With picDDE
                .LinkMode = 0
                .LinkTopic = "LockPro|LockPro"
                .LinkMode = 2
                .LinkExecute "lock"
                t = .LinkTimeout
                .LinkTimeout = 1
                .LinkMode = 0
                .LinkTimeout = t
            End With
            On Error GoTo 0
        End If
        End
    End If
    
    If LCase(Command) = "lock" And ReadCon("First") = 0 Then frmLock.Show 1
    If ReadCon("UNR") = 1 Then
        RtlAdjustPrivilege 20, 1, 0, 0
        RtlSetProcessIsCritical 0, 0, 1
    End If
    
    FormOnTop Me
    If ReadCon("First") = 1 Then
        Me.Move (Screen.Width - Me.Width) / 2, -20, pngMenu.Left + pngMenu.Width, pngLock.Height
    Else
        Me.Move CSng(ReadCon("FormLeft")), _
        -20, pngMenu.Left + pngMenu.Width, pngLock.Height
    End If
    
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    setBorderColor Me.hwnd, RGB(0, 170, 255)
    
    Me.Show
    If ReadCon("First") = 1 Then
        Me.Enabled = False
        frmGuide.Show
    Else
        lblUpd.Move 120, pngLock.Height + 75
        lblMore.Move Me.ScaleWidth - lblMore.Width - 75, pngLock.Height + (lblUpd.Height + 150 - lblMore.Height) / 2
        Updater.BeginDownload "http://maxxsoft.net/Update/LockPro/update.html", MyPath & "ver.log"  '检查更新
        SetSocket
    End If
End Sub

Private Sub lblMore_Click()
    tmrSUp.Tag = ""
    tmrSUp.Enabled = True
End Sub

Private Sub lblUpd_Click()
    If Dir(MyPath & "update.exe") = "" Then
        Shell "rundll32.exe url.dll,FileProtocolHandler http://maxxsoft.net/lockpro.html", vbNormalFocus
    Else
        ShellExecute 0, "open", MyPath & "update.exe", "", "", 1
    End If
    tmrSUp.Tag = ""
    tmrSUp.Enabled = True
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
    If ReadCon("ExitPsw") = 1 Then
        frmExit.ChangeMode 0
        frmExit.Show 1
    Else
        CloseLP
    End If
End Sub

Private Sub mnuGuide_Click()
    frmGuide.Show 1
End Sub

Private Sub mnuLock_Click()
    On Error Resume Next
    frmLock.Show
End Sub

Private Sub mnuSet_Click()
    Me.Enabled = False
    frmSettings.Show ' 1
End Sub

Private Sub mnuTimer_Click()
    frmTimer.Show
End Sub

Private Sub pngLock_DblClick(ByVal Button As Integer)
    mnuLock_Click
End Sub

Private Sub pngLock_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngLock.FadeInOut 80, 5
End Sub

Private Sub pngLock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Static ox!
    With Me
        If Button = 1 Then
            .Left = .Left - ox + x
            .Top = -20
        Else
            ox = x
        End If
    End With
End Sub

Private Sub pngLock_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngLock.FadeInOut 100, 2
End Sub

Private Sub pngMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngMenu.FadeInOut 80, 5
End Sub

Private Sub pngMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngMenu.FadeInOut 100, 2
    PopupMenu mnuMain, 0, pngMenu.Left, pngMenu.Height
End Sub

Private Sub tmrCntDwn_Timer()
    lCdTime = lCdTime - 1
    If lCdTime <= 0 Then
        Unload frmExit
        mnuLock_Click
        If TimerShowed Then Unload frmTimer
        tmrCntDwn.Enabled = False
    End If
    If TimerShowed Then
        With frmTimer
            .txtTime.Text = CStr(Int(lCdTime / (2 * 60)) + 1)
            If .btnStop.Caption = "　　　　　　　　　　　　　" Then
                .btnStop.Caption = "正在计时，再次点按停止计时"
            Else
                .btnStop.Caption = "　　　　　　　　　　　　　"
            End If
        End With
    End If
End Sub

Private Sub tmrKey_Timer()
    GetCursorPos lpPoint
    If lpPoint.y < ScaleY(Me.Top + Me.Height, vbTwips, vbPixels) Then
        If Me.Top < -20 Then
            tmrSMe.Tag = ""
            tmrSMe.Enabled = True
        End If
    Else
        If Me.Top > -pngLock.Height + 40 Then
            tmrSMe.Tag = "Up"
            tmrSMe.Enabled = True
        End If
    End If
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(Asc(ReadCon("Key"))) Then
        If ReadCon("First") = 0 Then frmLock.Show
    End If
    If Locked And GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyDelete) Then
        'Shell "shutdown.exe -s -t 0"
        NtShutdown
    End If
    If Locked And GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyEscape) Then
        'Shell "shutdown.exe -s -t 0"
        NtShutdown
    End If
End Sub

Private Sub tmrSMe_Timer()
    With Me
        If tmrSMe.Tag = "" Then
            .Top = .Top + GetMoveNum(-20, .Top, 3)
            If GetMoveNum(-20, .Top, 3) = 0 Then
                .Top = -20
                tmrSMe.Enabled = False
            End If
        Else
            If .Left < 0 Then .Left = .Left + GetMoveNum(0, .Left, 3)
            If .Left > Screen.Width - .Width Then .Left = .Left + GetMoveNum(Screen.Width - .Width, .Left, 3)
            .Top = .Top + GetMoveNum(-pngLock.Height + 40, .Top, 3)
            If GetMoveNum(-pngLock.Height + 40, .Top, 3) = 0 Then
                .Top = -pngLock.Height + 40
                tmrSMe.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub tmrSUp_Timer()
    With Me
        If tmrSUp.Tag = "" Then
            .Height = .Height + GetMoveNum(pngLock.Height, .Height, 3)
            If GetMoveNum(pngLock.Height, .Height, 3) = 0 Then
                .Height = pngLock.Height
                tmrSUp.Enabled = False
            End If
        Else
            .Height = .Height + GetMoveNum(lblUpd.Top + lblUpd.Height + 75, .Height, 3)
            If GetMoveNum(lblUpd.Top + lblUpd.Height + 75, .Height, 3) = 0 Then
                .Height = lblUpd.Top + lblUpd.Height + 75
                tmrSUp.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub Updater_DownloadComplete(MaxBytes As Long, SaveFile As String)
    On Error GoTo ChkErr
    Dim sTmp As String, sVer As String, sVers() As String, nVer As Long
    Open MyPath & "ver.log" For Input As #2
        Do While Not EOF(2)
            Line Input #2, sTmp
            sVer = sVer & sTmp & vbCrLf
        Loop
    Close #2
    sVers = Split(sVer, "@@")
    nVer = CLng(Format(sVers(1), "0000") & Format(sVers(2), "0000") & Format(sVers(3), "0000"))
    If nVer > CLng(Format(App.Major, "0000") & Format(App.Minor, "0000") & Format(App.Revision, "0000")) Then
        tmrSUp.Tag = "1"
        tmrSUp.Enabled = True
    Else
        tmrSUp.Tag = ""
        tmrSUp.Enabled = True
    End If
    If Dir(MyPath & "ver.log") <> "" Then Kill MyPath & "ver.log"
    Exit Sub
ChkErr:
    tmrSUp.Tag = ""
    tmrSUp.Enabled = True
End Sub
