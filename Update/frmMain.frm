VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "更新 Lock Pro"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5415
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrSFrm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   360
   End
   Begin LP_Updt.Downloader dwnMain 
      Left            =   480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin LP_Updt.Downloader dwnChk 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   480
      Width           =   5415
      Begin VB.PictureBox picUpd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   360
         ScaleHeight     =   1455
         ScaleWidth      =   4695
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   4695
         Begin LP_Updt.ucProBar prbUpd 
            Height          =   135
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   238
         End
         Begin VB.Label lblShow 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "已下载 0KB/0MB"
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
            Index           =   5
            Left            =   2970
            TabIndex        =   11
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "正在更新 ..."
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Index           =   4
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox picBtn 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   5415
         TabIndex        =   3
         Top             =   4080
         Width           =   5415
         Begin LP_Updt.ucBtn btnUpd 
            Height          =   612
            Left            =   1680
            TabIndex        =   4
            Top             =   120
            Width           =   2052
            _ExtentX        =   3625
            _ExtentY        =   1085
            Caption         =   "立即更新"
            FontSize        =   10.5
         End
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "更新内容："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   4680
      End
      Begin VB.Label lblShow 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本 0.0.0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   6
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MaxXSoft Lock Pro"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   3210
      End
      Begin LP_Updt.PngImage pngIcon 
         Height          =   1800
         Left            =   240
         Top             =   240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   3175
         Image           =   "frmMain.frx":000C
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin LP_Updt.PngImage pngUpd 
      Height          =   3870
      Index           =   2
      Left            =   0
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6826
      Image           =   "frmMain.frx":78CD
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin LP_Updt.PngImage pngUpd 
      Height          =   3870
      Index           =   1
      Left            =   0
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6826
      Image           =   "frmMain.frx":96DC
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "您的 Lock Pro 为最新版本，点击重新检查"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Index           =   0
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   4800
      Width           =   5340
   End
   Begin LP_Updt.PngImage pngUpd 
      Height          =   3870
      Index           =   0
      Left            =   0
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6826
      Image           =   "frmMain.frx":B153
      OLEdrop         =   1
      Props           =   5
   End
   Begin LP_Updt.PngImage pngCtrl 
      Height          =   495
      Index           =   1
      Left            =   4440
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Image           =   "frmMain.frx":C985
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin LP_Updt.PngImage pngCtrl 
      Height          =   495
      Index           =   0
      Left            =   4920
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Image           =   "frmMain.frx":CE1B
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   4920
      Picture         =   "frmMain.frx":D2B1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "frmMain.frx":D323
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "更新 Lock Pro"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sUpdateUrl As String = "http://maxxsoft.net/update/lockpro/update.html"

Dim lBlink As Long, sUrl As String
Dim Shdw As cShadow

Private Sub btnUpd_Click()
    lblShow(5).Caption = ""
    Select Case btnUpd.Caption
        Case "立即更新", "重试"
            picUpd.Visible = True
            If Dir(MyPath & "lpupdate.exe") <> "" Then Kill MyPath & "lpupdate.exe"
            dwnMain.CancelAllDownload
            dwnMain.BeginDownload sUrl, _
            MyPath & "lpupdate.exe"
            lblShow(4).Caption = "正在下载更新 ..."
            btnUpd.Caption = "取消更新"
        Case "取消更新"
            dwnMain.CancelAllDownload
            prbUpd.ChangePro 0, 1
            picUpd.Visible = False
            lblShow(4).Caption = ""
            btnUpd.Caption = "立即更新"
    End Select
End Sub

Private Sub dwnChk_DownloadComplete(MaxBytes As Long, SaveFile As String)
    On Error GoTo ChkErr
    Dim sTmp As String, sVer As String, sVers() As String, nVer As Long
    Open MyPath & "chk.log" For Input As #2
        Do While Not EOF(2)
            Line Input #2, sTmp
            sVer = sVer & sTmp & vbCrLf
        Loop
    Close #2
    sVers = Split(sVer, "@@")
    nVer = CLng(Format(sVers(1), "0000") & Format(sVers(2), "0000") & Format(sVers(3), "0000"))
    If nVer > CLng(Format(App.Major, "0000") & Format(App.Minor, "0000") & Format(App.Revision, "0000")) Then
        lblShow(0).Caption = "发现新版本 " & CLng(sVers(1)) & "." & CLng(sVers(2)) & "." & CLng(sVers(3))
        lblShow(2).Caption = "版本 " & CLng(sVers(1)) & "." & CLng(sVers(2)) & "." & CLng(sVers(3))
        lblShow(3).Caption = "更新内容：" & sVers(4)
        sUrl = sVers(5)
        pngUpd(1).FadeInOut 100, 5
    Else
        lblShow(0).Caption = "您的 Lock Pro 为最新版本，点击重新检查"
    End If
    If Dir(MyPath & "chk.log") <> "" Then Kill MyPath & "chk.log"
    Exit Sub
ChkErr:
    lblShow(0).Caption = "您的 Lock Pro 为最新版本，点击重新检查"
End Sub

Private Sub dwnMain_DownloadComplete(MaxBytes As Long, SaveFile As String)
    If SaveFile = MyPath & "lpupdate.exe" Then
        prbUpd.ChangePro 0, 1
        If Dir(MyPath & "lpupdate.exe") <> "" Then
            ShellExecute 0, "open", MyPath & "lpupdate.exe", MyPath & "@@" & MyPath & App.EXEName & ".exe", "", 1
            Unload Me
        Else
            lblShow(4).Caption = "更新出错！请稍后重试更新"
            lblShow(5).Caption = "更新文件被意外删除"
            btnUpd.Caption = "重试"
        End If
    End If
End Sub

Private Sub dwnMain_DownloadError(SaveFile As String)
    If SaveFile = MyPath & "lpupdate.exe" Then
        lblShow(4).Caption = "更新出错！请检查您的网络连接"
        lblShow(5).Caption = "下载错误"
        prbUpd.ChangePro 0, 1
        btnUpd.Caption = "重试"
    End If
End Sub

Private Sub dwnMain_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    If SaveFile = MyPath & "lpupdate.exe" Then
        prbUpd.ChangePro CurBytes, MaxBytes
        lblShow(5).Caption = "已下载 " & NumToByte(CurBytes, 1) & "/" & NumToByte(MaxBytes, 1)
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then End
    lblCap.Top = (480 - lblCap.Height) / 2
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    lblShow(0).Top = (pngUpd(0).Top + pngUpd(0).Height) / 2 + (Me.ScaleHeight - lblShow(0).Height) / 2
    picFrm.Left = Me.ScaleWidth
    dwnChk.BeginDownload sUpdateUrl, MyPath & "chk.log"            'Check Version
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Static ox!, oy!
    With Me
        If Button = 1 Then
            .Move .Left - ox + x, .Top - oy + y
        Else
            ox = x
            oy = y
        End If
    End With
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub lblShow_Click(Index As Integer)
    If Index = 0 And InStr(lblShow(0).Caption, "检查") Then
        dwnChk.CancelAllDownload
        dwnChk.BeginDownload sUpdateUrl, MyPath & "chk.log"            'Check Version
    End If
End Sub

Private Sub pngCtrl_Click(Index As Integer, ByVal Button As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            Me.WindowState = 1
    End Select
End Sub

Private Sub pngCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    pngCtrl(Index).FadeInOut 80, 5
End Sub

Private Sub pngCtrl_MouseEnter(Index As Integer)
    pngCtrl(Index).FadeInOut 50, 5
End Sub

Private Sub pngCtrl_MouseExit(Index As Integer)
    pngCtrl(Index).FadeInOut 0, 5
End Sub

Private Sub pngCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    pngCtrl(Index).FadeInOut 0, 5
End Sub

Private Sub pngUpd_FadeTerminated(Index As Integer, ByVal CurrentOpacity As Long)
    If Index = 1 Then
        pngUpd(2).FadeInOut 100
    ElseIf Index = 2 Then
        If lBlink Mod 2 = 0 Then
            pngUpd(2).FadeInOut 0
        Else
            pngUpd(2).FadeInOut 100
        End If
        lBlink = lBlink + 1
        If lBlink >= 3 Then
            lBlink = 0
            Sleep 500
            tmrSFrm.Enabled = True
            Exit Sub
        End If
    End If
End Sub

Private Sub tmrSFrm_Timer()
    With picFrm
        .Left = .Left + GetMoveNum(0, .Left, 4)
        If GetMoveNum(0, .Left, 4) = 0 Then
            .Left = 0
            tmrSFrm.Enabled = False
        End If
    End With
End Sub
