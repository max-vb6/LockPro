VERSION 5.00
Begin VB.Form frmTimer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Lock Pro 定时器"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5415
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrErr 
      Interval        =   3000
      Left            =   0
      Top             =   480
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      BackColor       =   &H00595959&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   65.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "1"
      Top             =   1155
      Width           =   2535
   End
   Begin LockPro.ucBtn btnStop 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1720
      Caption         =   "开始计时"
      FontSize        =   18
   End
   Begin LockPro.PngImage pngCtrl 
      Height          =   495
      Index           =   1
      Left            =   4440
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmTimer.frx":000C
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngCtrl 
      Height          =   495
      Index           =   0
      Left            =   4920
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmTimer.frx":04A2
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "计时结束后将锁定屏幕"
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
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   4830
      Width           =   5400
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "单位：分钟"
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
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   5400
   End
   Begin LockPro.PngImage pngTimer 
      Height          =   2700
      Left            =   360
      Top             =   600
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   4763
      Image           =   "frmTimer.frx":0938
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Pro 定时器"
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
      Width           =   1560
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   4920
      Picture         =   "frmTimer.frx":28E4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "frmTimer.frx":2956
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shdw As cShadow

Sub StopTimer()
    frmTray.tmrCntDwn.Enabled = False
    btnStop.Caption = "开始计时"
    txtTime.Enabled = True
End Sub

Private Sub btnStop_Click()
    With txtTime
        If btnStop.Caption = "开始计时" Then
            If Not IsNumeric(.Text) Then
ShowTimerErr:
                lblShow(0).Caption = "请输入正确的数字！"
                lblShow(0).ForeColor = vbRed
                tmrErr.Enabled = True
                Exit Sub
            ElseIf CInt(.Text) < 1 Or CInt(.Text) > 999 Then
                GoTo ShowTimerErr
            End If
            btnStop.Caption = "　　　　　　　　　　　　　"
            .Enabled = False
            lCdTime = CLng(.Text) * 60 * 2     '每500毫秒记一次，所以乘二
            frmTray.tmrCntDwn.Enabled = True
        Else
            If ReadCon("ExitPsw") = 1 Then
                frmExit.ChangeMode 1
                frmExit.Show 1
            Else
                StopTimer
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    lblCap.Top = (480 - lblCap.Height) / 2
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    TimerShowed = True
    If frmTray.tmrCntDwn.Enabled Then
        txtTime.Enabled = False
        btnStop.Caption = "　　　　　　　　　　　　　"
    End If
    txtTime.Text = CStr(Int(lCdTime / (2 * 60)) + 1)
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

Private Sub Form_Unload(Cancel As Integer)
    TimerShowed = False
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub lblShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
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

Private Sub pngTimer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub tmrErr_Timer()
    With lblShow(0)
        .Caption = "单位：分钟"
        .ForeColor = &HC0C0C0
    End With
    tmrErr.Enabled = False
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnStop_Click
    End If
End Sub
