VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "关于 Lock Pro"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   7095
      TabIndex        =   1
      Top             =   480
      Width           =   7095
      Begin LockPro.PngImage pngLk 
         Height          =   1920
         Left            =   660
         Top             =   120
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   3387
         Image           =   "frmAbout.frx":000C
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin LockPro.PngImage pngCls 
      Height          =   495
      Left            =   6600
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmAbout.frx":3B10
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关于 Lock Pro"
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
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   80
      Width           =   1350
   End
   Begin VB.Image imgCls 
      Height          =   480
      Left            =   6600
      Picture         =   "frmAbout.frx":3FA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MaxXSoft Lock Pro　　　　　版本 1.xx.00xx"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   7065
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":4018
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1140
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   7065
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "版权所有：2015 MaxXSoft 曼软工作室"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "点击进入 MaxXSoft 主页"
      Top             =   4800
      Width           =   7065
   End
   Begin VB.Label lblEgg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":40CF
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   6135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DblCnt As Long, hHandCur As Long
Dim Shdw As cShadow

Private Sub Form_Load()
    lblShow(4).Top = (480 - lblShow(4).Height) / 2
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    lblShow(0).Caption = "MaxXSoft Lock Pro　　　　　版本 " & _
        App.Major & "." & App.Minor & "." & Format(App.Revision, "0000")
    pngLk.Opacity = 0
    pngLk.FadeInOut 100, 1
    lblEgg.Move (Me.ScaleWidth - lblEgg.Width) / 2, (Me.ScaleHeight - lblEgg.Height) / 2
    hHandCur = LoadCursor(0&, IDC_HAND)       '载入默认手形鼠标
    DblCnt = 0
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

Private Sub lblLink_Click()
    Shell "rundll32.exe url.dll,FileProtocolHandler http://maxxsoft.net/", vbNormalFocus
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor hHandCur
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor hHandCur
End Sub

Private Sub lblShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub pngCls_Click(ByVal Button As Integer)
    Unload Me
End Sub

Private Sub pngCls_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngCls.FadeInOut 80, 5
End Sub

Private Sub pngCls_MouseEnter()
    pngCls.FadeInOut 50, 5
End Sub

Private Sub pngCls_MouseExit()
    pngCls.FadeInOut 0, 5
End Sub

Private Sub pngCls_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngCls.FadeInOut 0, 5
End Sub

Private Sub pngLk_DblClick(ByVal Button As Integer)
    DblCnt = DblCnt + 1
    If DblCnt > 10 Then
        picFrm.Visible = False
        lblShow(0).Visible = False
        lblShow(2).Visible = False
        lblShow(4).Visible = False
        lblLink.Visible = False
        Me.BackColor = vbBlack
        lblEgg.Visible = True
    End If
End Sub
