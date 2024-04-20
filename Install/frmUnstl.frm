VERSION 5.00
Begin VB.Form frmUnstl 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "卸载 Lock Pro"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   Icon            =   "frmUnstl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7215
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   0
      Left            =   2400
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   480
      Width           =   4815
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "在卸载之前请确定您的 Lock Pro 没有设置为开机启动。"
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
         Height          =   1395
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   4185
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "您确定要卸载 Lock Pro 吗？"
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
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   3060
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "不忍心，和你说再见"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2835
      End
   End
   Begin LP_Unstl.ucBtn btnUnstl 
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "卸载 Lock Pro"
      FontSize        =   10.5
   End
   Begin VB.PictureBox picLogoBG 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      Begin LP_Unstl.PngImage pngLogo 
         Height          =   2040
         Left            =   240
         Top             =   960
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   3598
         Image           =   "frmUnstl.frx":000C
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin LP_Unstl.PngImage pngCtrl 
      Height          =   495
      Index           =   1
      Left            =   6240
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Image           =   "frmUnstl.frx":7B94
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin LP_Unstl.PngImage pngCtrl 
      Height          =   495
      Index           =   0
      Left            =   6720
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Image           =   "frmUnstl.frx":802A
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   6240
      Picture         =   "frmUnstl.frx":84C0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   6720
      Picture         =   "frmUnstl.frx":851F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卸载 Lock Pro"
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
      TabIndex        =   1
      Top             =   80
      Width           =   1350
   End
End
Attribute VB_Name = "frmUnstl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shdw As cShadow

Private Sub btnUnstl_Click()
    On Error GoTo UErr
    If btnUnstl.Caption = "卸载 Lock Pro" Then
        pngCtrl(0).Enabled = False
        btnUnstl.Caption = "卸载中..."
        lblShow(0).Caption = "正在卸载 Lock Pro"
        lblShow(1).Caption = ""
        lblShow(2).Caption = "准备中..."
        If CheckExeIsRun("LockPro.exe") Then
            MsgBox "卸载程序检测到您的 Lock Pro 正在运行。" & vbCrLf & "为了卸载进程的顺利进行，请先手动关闭 Lock Pro。" & _
                vbCrLf & "若您不手动关闭 Lock Pro，卸载程序将在您点击“确定”后强制结束 Lock Pro。", 48, "Lock Pro 正在运行"
        End If
        If CheckExeIsRun("LockPro.exe") Then Shell "taskkill /f /im LockPro.exe", 0
        Sleep 1000
        Dim lTry As Long
        lTry = 0
StartUni:
        lblShow(2).Caption = "删除开始菜单项..."                    '删除文件和目录
        If Dir(GetStartMenuPath & "MaxXSoft Lock Pro\", vbDirectory) <> "" Then DeleteFolder GetStartMenuPath & "MaxXSoft Lock Pro\"
        lblShow(2).Caption = "删除桌面快捷方式..."
        If Dir(GetStartMenuPath(0) & "Lock Pro.lnk") <> "" Then Kill GetStartMenuPath(0) & "Lock Pro.lnk"
        If Dir(GetStartMenuPath(0) & "立即使用 Lock Pro 锁定计算机.lnk") <> "" Then Kill GetStartMenuPath(0) & "立即使用 Lock Pro 锁定计算机.lnk"
        If Dir(App.Path & "\mswinsck.ocx") <> "" Then
            lblShow(2).Caption = "解除注册组件..."
            Shell "regsvr32 /s /u """ & App.Path & "\" & "mswinsck.ocx""", vbHide
        End If
        lblShow(2).Caption = "删除程序文件 ..."
        If Dir(App.Path & "\", vbDirectory) <> "" Then DeleteFolder App.Path & "\"
'       lblShow(2).Caption = "删除注册表数据..."
'        ChangeReg HKEY_LOCAL_MACHINE, "SOFTWARE\" & _
            "Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", Environ("Windir") & "\system32\userinit.exe"
        pngCtrl(0).Enabled = True
        btnUnstl.Caption = "关闭"
        lblShow(0).Caption = "Lock Pro 已从您的电脑上移除"
        lblShow(2).Caption = "感谢您使用曼软的软件，我们未来会做得更好！" & vbCrLf & vbCrLf & "点击“关闭”完成卸载"
    ElseIf btnUnstl.Caption = "关闭" Then
        DeleteFolder App.Path & "\"
        Unload Me
    End If
    Exit Sub
UErr:
    If lTry < 2 Then
        lTry = lTry + 1
        Sleep 1000
        GoTo StartUni
    Else
        MsgBox "卸载过程中出现错误，请稍后重新运行卸载程序！", 48, "出错啦！"
        End
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
    pngLogo.Opacity = 0
    pngLogo.FadeInOut 100, 3
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

Private Sub pngCtrl_Click(Index As Integer, ByVal Button As Integer)
    Select Case Index
        Case 0
            If btnUnstl.Caption = "关闭" Then btnUnstl_Click: Exit Sub
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
