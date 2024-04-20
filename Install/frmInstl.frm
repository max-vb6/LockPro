VERSION 5.00
Begin VB.Form frmInstl 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "安装 Lock Pro"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   Icon            =   "frmInstl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7335
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSel 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   375
      TabIndex        =   22
      Top             =   4320
      Width           =   375
   End
   Begin VB.Timer tmrSel 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin LP_Instl.ucBtn btnUnl 
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Caption         =   "下一步"
      FontSize        =   10.5
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      Begin LP_Instl.PngImage pngLogo 
         Height          =   2160
         Left            =   240
         Top             =   840
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   3810
         Image           =   "frmInstl.frx":000C
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   4
      Left            =   2520
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   19
      Top             =   480
      Width           =   4815
      Begin LP_Instl.ucSwitch swiRun 
         Height          =   375
         Left            =   3360
         TabIndex        =   29
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   -1  'True
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "如果 Lock Pro 未正常运行，请尝试使用兼容模式"
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
         Height          =   915
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   4020
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "立即运行 Lock Pro"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "点击“完成”退出安装向导"
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
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   2880
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Pro 已安装在您的计算机上"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   3
      Left            =   2520
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   10
      Top             =   480
      Width           =   4815
      Begin VB.Label lblPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "准备中..."
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
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在安装 Lock Pro"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2535
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   2
      Left            =   2520
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   13
      Top             =   480
      Width           =   4815
      Begin LP_Instl.ucSwitch swiOpt 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   14
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   -1  'True
      End
      Begin LP_Instl.ucSwitch swiOpt 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   -1  'True
      End
      Begin LP_Instl.ucSwitch swiOpt 
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   -1  'True
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保留之前的设置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "如果您发现快捷方式没有正常创建，请尝试关闭系统防护软件并重新安装"
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
         Height          =   915
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   4200
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "创建桌面快捷方式"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1920
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "创建开始菜单快捷方式"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2400
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   2520
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   6
      Top             =   480
      Width           =   4815
      Begin VB.PictureBox picHide 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   30
         Top             =   1800
         Width           =   495
      End
      Begin LP_Instl.ucBtn btnSel 
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "..."
         FontSize        =   10.5
      End
      Begin VB.Label lblSpace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可用空间未计算"
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
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   1470
      End
      Begin VB.Label lblDir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C:\Program Files\Lock Pro\"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3090
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择 Lock Pro 的安装位置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3825
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   0
      Left            =   2520
      ScaleHeight     =   3855
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   480
      Width           =   4815
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInstl.frx":7B94
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
         Height          =   1515
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "欢迎进入 Lock Pro 2 安装向导"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "完成"
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
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   27
      Top             =   4750
      Width           =   360
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安装"
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
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   26
      Top             =   4750
      Width           =   360
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "杂项"
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
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   25
      Top             =   4750
      Width           =   360
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设定位置"
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
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   24
      Top             =   4750
      Width           =   720
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎"
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
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   4750
      Width           =   360
   End
   Begin LP_Instl.PngImage pngCtrl 
      Height          =   495
      Index           =   1
      Left            =   6360
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Image           =   "frmInstl.frx":7BF8
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin LP_Instl.PngImage pngCtrl 
      Height          =   495
      Index           =   0
      Left            =   6840
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Image           =   "frmInstl.frx":808E
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   6360
      Picture         =   "frmInstl.frx":8524
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   6840
      Picture         =   "frmInstl.frx":8583
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安装 Lock Pro"
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
      Top             =   75
      Width           =   1350
   End
End
Attribute VB_Name = "frmInstl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LP_MinMB As Long = 15
Dim lNow As Long, isReged As Boolean ', isGrnVer As Boolean
Dim Shdw As cShadow

Private Sub btnSel_Click()
    Dim sDir As String
    sDir = GetDirectory(Me.hwnd, "选择 Lock Pro 的安装位置")
    If sDir = "" Then Exit Sub
    lblDir.Caption = sDir & "Lock Pro\"
    lblDir.ToolTipText = lblDir.Caption
    lblSpace.Caption = GetDriveSpaceString(lblDir.Caption, LP_MinMB)
End Sub

Private Sub btnUnl_Click()
    On Error GoTo InsErr
    If btnUnl.Caption = "下一步" Then
        If lNow = 1 And InStr(lblSpace.Caption, "不足") <> 0 Then Beep: Exit Sub
        lNow = lNow + 1
        picFrm(lNow).ZOrder 0
        picSel.Width = lblStep(lNow).Width
        tmrSel.Enabled = True
        If lNow = 2 Then
            btnUnl.Caption = "安装"
            If Dir(lblDir.Caption & "LockProCfg.ini") <> "" And _
            Dir(lblDir.Caption & "LockPro.xm5") <> "" Then
                lblShow(10).Visible = True
                swiOpt(2).Visible = True
                'If Dir(lblDir.Caption & "Uninstall.exe") = "" Then
                    'isGrnVer = True
                'Else
                    'isGrnVer = False
                'End If
            End If
        End If
    ElseIf btnUnl.Caption = "安装" Then
        lNow = 3
        picSel.Width = lblStep(lNow).Width
        tmrSel.Enabled = True
        pngCtrl(0).Enabled = False
        picFrm(3).ZOrder 0
        btnUnl.Caption = "安装中..."
Instl:
        If CheckExeIsRun("LockPro.exe") Then
            MsgBox "安装程序检测到您的 Lock Pro 正在运行。" & vbCrLf & "为了安装进程的顺利进行，请先手动关闭 Lock Pro。" & _
                vbCrLf & "若您不手动关闭 Lock Pro，安装程序将在您点击“确定”后强制结束 Lock Pro。", 48, "Lock Pro 正在运行"
        End If
        If CheckExeIsRun("LockPro.exe") Then Shell "taskkill /f /im LockPro.exe", 0
        Sleep 1000
        'Install Start
        lblPro.Caption = "正在创建文件夹..."
        If Dir(lblDir.Caption, vbDirectory) = "" Then MkDir lblDir.Caption
        Sleep 100
        If swiOpt(2).Visible And swiOpt(2).Checked Then
            lblPro.Caption = "正在备份设置..."
            If Dir(lblDir.Caption & "LockProCfg.ini") <> "" And _
            Dir(lblDir.Caption & "LockPro.xm5") <> "" Then
                If Dir(lblDir.Caption & "LockProCfg.ini.bak") <> "" Then Kill lblDir.Caption & "LockProCfg.ini.bak"
                If Dir(lblDir.Caption & "LockPro.xm5.bak") <> "" Then Kill lblDir.Caption & "LockPro.xm5.bak"
                Name lblDir.Caption & "LockProCfg.ini" As lblDir.Caption & "LockProCfg.ini.bak"
                Name lblDir.Caption & "LockPro.xm5" As lblDir.Caption & "LockPro.xm5.bak"
            End If
            Sleep 100
        End If
        isReged = (Dir(lblDir.Caption & "mswinsck.ocx") <> "")                                  '记录安装前注册状态，防止误删除 OCX
        lblPro.Caption = "正在提取文件..."
        SaveFileFromRes 101, "CUSTOM", lblDir.Caption & "7z.exe"
        SaveFileFromRes 102, "CUSTOM", lblDir.Caption & "LockPro.7z"
        SaveFileFromRes 103, "CUSTOM", lblDir.Caption & "7z.dll"
        Sleep 500
        lblPro.Caption = "正在展开文件..."
        Shell lblDir.Caption & "7z.exe x """ & lblDir.Caption & "LockPro.7z"" -y -o""" & _
        lblDir.Caption & """", 0
        Sleep 2000
        lblPro.Caption = "正在删除临时文件..."
        Kill lblDir.Caption & "7z.exe"
        Kill lblDir.Caption & "LockPro.7z"
        Kill lblDir.Caption & "7z.dll"
        'If isGrnVer Then Kill lblDir.Caption & "Uninstall.exe"
        If CheckWinsockOCX And Not isReged Then Kill lblDir.Caption & "mswinsck.ocx"            '若已经在系统其他地方注册 OCX 则删除本目录的 OCX，防止二次注册
        Sleep 500
        If Dir(lblDir.Caption & "mswinsck.ocx") <> "" Then
            lblPro.Caption = "正在注册组件..."
            Shell "regsvr32.exe /s """ & lblDir.Caption & "mswinsck.ocx""", vbHide
            Shell lblDir.Caption & "LockPro.exe Reg@@" & CStr(&H80000000) & "@@Licenses\2c49f800-c2dd-11cf-9ad6-0080c7e7b78d@@@@mlrljgrlhltlngjlthrligklpkrhllglqlrk"
            Sleep 500
        End If
        If swiOpt(2).Visible And swiOpt(2).Checked Then
            lblPro.Caption = "正在还原设置..."
            If Dir(lblDir.Caption & "LockProCfg.ini.bak") <> "" And _
            Dir(lblDir.Caption & "LockPro.xm5.bak") <> "" Then
                MoveCfg lblDir.Caption & "LockProCfg.ini.bak", lblDir.Caption & "LockProCfg.ini"
                MoveCfg lblDir.Caption & "LockPro.xm5.bak", lblDir.Caption & "LockPro.xm5"
                Kill lblDir.Caption & "LockProCfg.ini.bak"
                Kill lblDir.Caption & "LockPro.xm5.bak"
            End If
            Sleep 500
        End If
        If swiOpt(0).Checked Or swiOpt(1).Checked Then
            lblPro.Caption = "正在创建快捷方式..."
            If swiOpt(0).Checked Then                '桌面快捷方式
                mShellLnk "Lock Pro", lblDir.Caption & "LockPro.exe", lblDir.Caption, _
                "LockPro.exe", "", "", "启动 Lock Pro"
            End If
            If swiOpt(1).Checked Then                '开始菜单快捷方式
                If Dir(GetStartMenuPath & "MaxXSoft Lock Pro\", vbDirectory) = "" Then MkDir GetStartMenuPath & "MaxXSoft Lock Pro\"
                Sleep 100
                mShellLnk "Lock Pro", lblDir.Caption & "LockPro.exe", lblDir.Caption, _
                "LockPro.exe", "", "", "启动 Lock Pro", GetStartMenuPath & "MaxXSoft Lock Pro"
                mShellLnk "卸载 Lock Pro", lblDir.Caption & "Uninstall.exe", lblDir.Caption, _
                "Uninstall.exe", "", "", "卸载 Lock Pro", GetStartMenuPath & "MaxXSoft Lock Pro"
            End If
            Sleep 100
        End If
        'Install End
        lNow = 4
        picSel.Width = lblStep(lNow).Width
        tmrSel.Enabled = True
        pngCtrl(0).Enabled = True
        picFrm(4).ZOrder 0
        btnUnl.Caption = "完成"
        RefreshShell
    ElseIf btnUnl.Caption = "完成" Then
        If swiRun.Checked And Dir(lblDir.Caption & "LockPro.exe") <> "" Then
            Shell lblDir.Caption & "LockPro.exe", 1
        End If
        If InStr(lblCap.Caption, "更新") <> 0 Then
            Shell "cmd.exe /c del /f /q " & """" & App.Path & "\" & App.exeName & ".exe""", vbHide
        End If
        Unload Me
    End If
    Exit Sub
InsErr:
    If Dir(lblDir.Caption, vbDirectory) <> "" Then DeleteFolder lblDir.Caption
    If MsgBox("安装过程中出现错误！" & _
        vbCrLf & "请稍后运行安装程序或者关闭系统防护软件重试", _
        48 + vbRetryCancel, "出错啦！") = vbRetry Then
        lblPro.Caption = "正在重试安装..."
        GoTo Instl
    End If
    End
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
    
    If Command <> "" And InStr(Command, "@@") <> 0 Then                                 'Auto install.
        Dim sCmds() As String
        sCmds = Split(Command, "@@")
        If Mid(sCmds(0), 2, 2) = ":\" And Mid(sCmds(0), Len(sCmds(0)), 1) = "\" Then
            Me.Hide
            If InStr(GetDriveSpaceString(sCmds(0), LP_MinMB), "不足") <> 0 Then
                MsgBox "您的磁盘空间不足，无法完成此次更新！", 48, "磁盘空间不足"
                Shell sCmds(1), vbNormalFocus
                End
            End If
            lblDir.Caption = sCmds(0)
            swiOpt(0).Checked = False
            swiOpt(1).Checked = False
            If Dir(sCmds(0) & "LockProCfg.ini") <> "" And _
            Dir(sCmds(0) & "LockPro.xm5") <> "" Then
                swiOpt(2).Visible = True
                swiOpt(2).Checked = True
                'If Dir(sCmds(0) & "Uninstall.exe") = "" Then
                    'isGrnVer = True
                'Else
                    'isGrnVer = False
                'End If
            End If
            lblCap.Caption = "更新 Lock Pro"
            lblShow(6).Caption = "Lock Pro 已完成更新"
            lblShow(7).Caption = "点击“完成”结束更新"
            btnUnl.Caption = "安装"
            Me.Show
            btnUnl_Click
            Exit Sub
        End If
    End If
    
    lNow = 0
    picFrm(0).ZOrder 0
    If Dir(Environ("Windir") & "\SysWOW64\", vbDirectory) = "" Then
        lblDir.Caption = "C:\Program Files\Lock Pro\"
    Else
        lblDir.Caption = "C:\Program Files (x86)\Lock Pro\"
    End If
    lblDir.ToolTipText = lblDir.Caption
    lblSpace.Caption = GetDriveSpaceString(lblDir.Caption, LP_MinMB)
    'isGrnVer = False
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

Private Sub tmrSel_Timer()
    With picSel
        .Left = .Left + GetMoveNum(lblStep(lNow).Left, .Left, 5)
        If GetMoveNum(lblStep(lNow).Left, .Left, 5) = 0 Then
            tmrSel.Enabled = False
        End If
    End With
End Sub
