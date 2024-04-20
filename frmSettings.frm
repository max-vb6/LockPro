VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Lock Pro 设置"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8535
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   6
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   73
      Top             =   480
      Width           =   6375
      Begin LockPro.ucBtn btnClear 
         Height          =   375
         Left            =   4080
         TabIndex        =   81
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "清空授权数据"
         FontSize        =   9
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   74
         Text            =   "80"
         Top             =   1800
         Width           =   735
      End
      Begin LockPro.ucSwitch swiRemo 
         Height          =   375
         Left            =   4680
         TabIndex        =   75
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "为了防止恶意使用 Remolock，企图使用 Remolock 远程功能的任何设备在未经授权的情况下将无权远程锁定计算机。"
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
         Height          =   1020
         Index           =   28
         Left            =   600
         TabIndex        =   82
         Top             =   3720
         Width           =   5055
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 个设备已被授权"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   33
         Left            =   600
         TabIndex        =   80
         Top             =   3240
         Width           =   1650
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "若要使用 Remolock 远程功能，请在 PC 或手机浏览器访问：http://192.168.0.xxx/ 。"
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
         Height          =   780
         Index           =   29
         Left            =   600
         TabIndex        =   79
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用端口"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   30
         Left            =   600
         TabIndex        =   78
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "该项默认值为 80。若您的 80 端口被其他软件占据，请设置其他端口。"
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
         Height          =   780
         Index           =   31
         Left            =   600
         TabIndex        =   77
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remolock 状态"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   32
         Left            =   600
         TabIndex        =   76
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Timer tmrSLst 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox picLst 
      BackColor       =   &H00CDCDCD&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5295
      ScaleWidth      =   2175
      TabIndex        =   3
      Top             =   480
      Width           =   2175
      Begin VB.PictureBox picSel 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   530
         Left            =   2040
         ScaleHeight     =   525
         ScaleWidth      =   135
         TabIndex        =   10
         Tag             =   "0"
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "杂项"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   72
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "便捷设置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remolock"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "个性化"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "错误捕获"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "USB设置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "密码设置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblLst 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "解锁选项"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin LockPro.PngImage pngMouse 
         Height          =   4785
         Left            =   0
         Top             =   240
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   8440
         Image           =   "frmSettings.frx":000C
         Scaler          =   1
         OLEdrop         =   1
         Props           =   0
      End
   End
   Begin LockPro.ucBtn btnOK 
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "保存设置"
      FontSize        =   10.5
   End
   Begin LockPro.ucBtn btnCancel 
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "关闭"
      FontSize        =   10.5
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   7
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   63
      Top             =   480
      Width           =   6375
      Begin VB.TextBox txtScr 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   68
         Text            =   "60"
         Top             =   960
         Width           =   735
      End
      Begin LockPro.ucSwitch swiScr 
         Height          =   375
         Left            =   4680
         TabIndex        =   65
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin LockPro.ucSwitch swiExit 
         Height          =   375
         Left            =   4680
         TabIndex        =   70
         Top             =   2760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSettings.frx":0024
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
         Height          =   1380
         Index           =   27
         Left            =   600
         TabIndex        =   71
         Top             =   3480
         Width           =   5055
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退出或停止计时时要求验证密码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   24
         Left            =   600
         TabIndex        =   69
         Top             =   2760
         Width           =   2940
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "等待时间　　　　　　秒 （至少 10 秒）"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   26
         Left            =   600
         TabIndex        =   67
         Top             =   960
         Width           =   3780
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "在您使用 Lock Pro 锁定计算机时，如果您长时间没有操作计算机，屏幕将会自动变暗以节能"
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
         Height          =   900
         Index           =   25
         Left            =   600
         TabIndex        =   66
         Top             =   1560
         Width           =   5055
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "锁屏屏幕保护"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   23
         Left            =   600
         TabIndex        =   64
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   1
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   16
      Top             =   480
      Width           =   6375
      Begin VB.PictureBox picOldPsw 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   720
         ScaleHeight     =   1695
         ScaleWidth      =   4215
         TabIndex        =   24
         Top             =   3240
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtPsw 
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   240
            PasswordChar    =   "l"
            TabIndex        =   25
            Top             =   960
            Width           =   3375
         End
         Begin VB.Label lblShow 
            BackStyle       =   0  'Transparent
            Caption         =   "如果您需要改变设置，请输入之前的密码进行确认"
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
            Height          =   825
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   3540
         End
      End
      Begin VB.TextBox txtPsw 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   960
         PasswordChar    =   "l"
         TabIndex        =   23
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtPsw 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   960
         PasswordChar    =   "l"
         TabIndex        =   21
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确认密码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   8
         Left            =   840
         TabIndex        =   22
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新密码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "为了防止输入错误，您需要输入两次密码"
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
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   3780
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   0
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   480
      Width           =   6375
      Begin LockPro.ucSwitch swiUnl 
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   14
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   -1  'True
      End
      Begin LockPro.ucSwitch swiUnl 
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   15
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "无需记下密码，解锁时只需插入 USB，有效防止密码破解。不会在 USB 上产生额外文件"
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
         Height          =   1380
         Index           =   4
         Left            =   600
         TabIndex        =   18
         Top             =   2040
         Width           =   3690
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "传统又熟悉的解锁方式"
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
         Index           =   2
         Left            =   600
         TabIndex        =   17
         Top             =   840
         Width           =   2100
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "您可以同时开启以上两个选项来提高锁定的安全程度"
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
         Index           =   3
         Left            =   600
         TabIndex        =   13
         Top             =   4440
         Width           =   4830
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "插入 USB 设备解锁"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入密码解锁"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   5
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   48
      Top             =   480
      Width           =   6375
      Begin LockPro.ucBtn btnCelAuto 
         Height          =   375
         Left            =   4200
         TabIndex        =   57
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "取消启动时锁屏"
         FontSize        =   9
      End
      Begin LockPro.ucSwitch swiDesktop 
         Height          =   375
         Left            =   4680
         TabIndex        =   51
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin LockPro.ucSwitch swiAuto 
         Height          =   375
         Left            =   4680
         TabIndex        =   52
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin LockPro.ucSwitch swiBlue 
         Height          =   375
         Left            =   4680
         TabIndex        =   53
         Top             =   3000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Checked         =   0   'False
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSettings.frx":00E4
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
         Height          =   1455
         Index           =   21
         Left            =   720
         TabIndex        =   58
         Top             =   3480
         Width           =   4740
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "如果您之前已经设置了启动时锁屏，您可以"
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
         Index           =   20
         Left            =   720
         TabIndex        =   56
         Top             =   2520
         Width           =   3420
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSettings.frx":01D1
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
         Height          =   855
         Index           =   19
         Left            =   720
         TabIndex        =   55
         Top             =   1560
         Width           =   4740
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开启无敌模式（慎用！需要管理员权限）"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   18
         Left            =   600
         TabIndex        =   54
         Top             =   3000
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "启动时立即运行 Lock Pro 锁屏"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   17
         Left            =   600
         TabIndex        =   50
         Top             =   1080
         Width           =   2880
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "在桌面创建快捷方式以便快速锁屏"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   16
         Left            =   600
         TabIndex        =   49
         Top             =   360
         Width           =   3150
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   4
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   41
      Top             =   480
      Width           =   6375
      Begin VB.TextBox txtLTxt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   61
         Text            =   "您的计算机已被锁定"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.FileListBox filPic 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   960
         Pattern         =   "*.jpg"
         TabIndex        =   47
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtKey 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   44
         Text            =   "L"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "锁定屏幕显示文字"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   22
         Left            =   600
         TabIndex        =   60
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Image imgPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   3480
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "如果您找到了心仪的图片，可以将图片放到程序的“LockPicture”文件夹内。点击文字打开壁纸文件夹"
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
         Height          =   1020
         Index           =   15
         Left            =   600
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   46
         Top             =   3840
         Width           =   4725
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置 Lock Pro 的锁屏背景"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   14
         Left            =   600
         TabIndex        =   45
         Top             =   1800
         Width           =   2460
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + "
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   13
         Left            =   4200
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "激活 Lock Pro 锁屏的快捷键"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   12
         Left            =   600
         TabIndex        =   42
         Top             =   360
         Width           =   2670
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   3
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   32
      Top             =   480
      Width           =   6375
      Begin VB.OptionButton optUnl 
         BackColor       =   &H00F0F0F0&
         Caption         =   "重新启动计算机"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   39
         Top             =   3120
         Width           =   3735
      End
      Begin VB.OptionButton optUnl 
         BackColor       =   &H00F0F0F0&
         Caption         =   "关闭计算机"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   38
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtSec 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   37
         Text            =   "30"
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton optUnl 
         BackColor       =   &H00F0F0F0&
         Caption         =   "等待　　　　　　秒"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   36
         Top             =   1920
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.TextBox txtCnt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "5"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "达到设定次数后"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   11
         Left            =   600
         TabIndex        =   35
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许输入错误解锁密码的次数"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   10
         Left            =   600
         TabIndex        =   33
         Top             =   360
         Width           =   2730
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   2
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   27
      Top             =   480
      Width           =   6375
      Begin LockPro.ucBtn btnRfsh 
         Height          =   495
         Left            =   3840
         TabIndex        =   29
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "刷新 USB 列表"
         FontSize        =   10.5
      End
      Begin VB.ComboBox cboUSB 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label lblUSB 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Pro 检测到您曾经设置过 USB 解锁设备。如需改变设置，请插入原先设置的 USB 设备然后点击我方可设置新的 USB 解锁设备。"
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
         Height          =   1260
         Left            =   720
         TabIndex        =   31
         Top             =   3000
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置解锁 Lock Pro 的 USB 设备"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   9
         Left            =   600
         TabIndex        =   30
         Top             =   360
         Width           =   2985
      End
   End
   Begin LockPro.PngImage pngCtrl 
      Height          =   495
      Index           =   0
      Left            =   8040
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmSettings.frx":0267
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "设置错误"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   59
      Top             =   6120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Pro 设置"
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
   Begin LockPro.PngImage pngCtrl 
      Height          =   495
      Index           =   1
      Left            =   7560
      ToolTipText     =   "最小化"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmSettings.frx":06FD
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   7560
      Picture         =   "frmSettings.frx":0B93
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   8040
      Picture         =   "frmSettings.frx":0BF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shdw As cShadow, bClear As Boolean

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCelAuto_Click()
    ChangeReg HKEY_LOCAL_MACHINE, "SOFTWARE\" & _
        "Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", Environ("Windir") & "\system32\userinit.exe"
    lblShow(20).Visible = False
    btnCelAuto.Visible = False
End Sub

Private Sub btnClear_Click()
    If InStr(lblShow(33).Caption, "0 个") <> 0 Then
        Beep
        Exit Sub
    End If
    bClear = (MsgBox("您将清空所有的授权数据，" & _
                    "这会导致已被授权的所有设备无法继续使用 Remolock，除非重新获得授权。" & vbCrLf & _
                    "此操作会在设置保存时生效，确认清空授权数据吗？", 48 + vbOKCancel, "注意") = vbOK)
    btnClear.Caption = IIf(bClear, "待清空...", "清空授权数据")
End Sub

Private Sub btnOK_Click()
    On Error Resume Next
    '========CheckSettings========
    If swiUnl(0).Checked Then
        If picOldPsw.Visible And XMD5(txtPsw(0).Text) <> ReadPsw("Psw") Then
            lblLst_Click (1)
            Beep
            ShowErr "请输入正确的旧密码！"
            txtPsw(0).Text = ""
            txtPsw(0).SetFocus
            Exit Sub
        End If
        If txtPsw(1).Text <> "" And txtPsw(2).Text <> txtPsw(1).Text Then
            lblLst_Click (1)
            Beep
            ShowErr "两次输入的密码不相符，请重新输入！"
            txtPsw(2).Text = ""
            txtPsw(2).SetFocus
            Exit Sub
        End If
        If picOldPsw.Visible = False And txtPsw(1).Text = "" Then
            lblLst_Click (1)
            Beep
            ShowErr "您还没有设置过密码，请务必设置您的解锁密码！"
            txtPsw(1).SetFocus
            Exit Sub
        End If
    End If
    If swiUnl(1).Checked And cboUSB.ListIndex = -1 Then
        lblLst_Click (2)
        Beep
        ShowErr "请选择正确的 USB 解锁设备！"
        cboUSB.SetFocus
        Exit Sub
    End If
    If Not (IsNumeric(txtCnt.Text)) Then
        lblLst_Click (3)
        Beep
        ShowErr "请输入正确的次数限制！"
        txtCnt.Text = ""
        txtCnt.SetFocus
        Exit Sub
    End If
    If txtSec.Enabled Then
        If Not (IsNumeric(txtSec.Text)) Then
            lblLst_Click (3)
            Beep
            ShowErr "请输入正确的等待秒数！"
            txtSec.Text = ""
            txtSec.SetFocus
            Exit Sub
        End If
    End If
    If txtKey.Text = "" Then
        lblLst_Click (4)
        Beep
        ShowErr "请输入正确的快捷键！"
        txtKey.SetFocus
        Exit Sub
    End If
    If txtLTxt.Text = "" Then
        lblLst_Click (4)
        Beep
        ShowErr "请输入正确的锁屏文字！"
        txtLTxt.SetFocus
        Exit Sub
    End If
    If filPic.FileName = "" Then
        lblLst_Click (4)
        Beep
        ShowErr "请选择锁屏背景！"
        filPic.SetFocus
        Exit Sub
    End If
    If Not (IsNumeric(txtPort.Text)) Or Len(txtPort.Text) < 2 Then
        lblLst_Click (6)
        Beep
        ShowErr "请输入正确的端口号！"
        txtPort.Text = ""
        txtPort.SetFocus
        Exit Sub
    End If
    If txtScr.Enabled Then
        If Not (IsNumeric(txtScr.Text)) Or Len(txtScr.Text) < 2 Then
            lblLst_Click (7)
            Beep
            ShowErr "请输入正确的屏幕保护等待秒数！"
            txtScr.Text = ""
            txtScr.SetFocus
            Exit Sub
        End If
    End If
    If swiExit.Checked Then
        If picOldPsw.Visible = False And txtPsw(1).Text = "" Then
            lblLst_Click (1)
            Beep
            ShowErr "您还没有设置过密码，请务必设置您的解锁密码！"
            txtPsw(1).SetFocus
            Exit Sub
        End If
    End If
    '===========EndCheck==========
    '=========SaveSettings========
    SaveCon "First", "0"
    If swiUnl(0).Checked And swiUnl(1).Checked Then
        SaveCon "Psw", "2"
        If txtPsw(1).Text <> "" Then SavePsw "Psw", XMD5(txtPsw(1).Text)
        SavePsw "USB", XMD5(GetUSBSerial(cboUSB.List(cboUSB.ListIndex)))
    ElseIf Not (swiUnl(1).Checked) Then
        SaveCon "Psw", "0"
        If txtPsw(1).Text <> "" Then SavePsw "Psw", XMD5(txtPsw(1).Text)
    Else
        SaveCon "Psw", "1"
        SavePsw "USB", XMD5(GetUSBSerial(cboUSB.List(cboUSB.ListIndex)))
    End If
    Dim i As Long
    For i = 0 To 2
        If optUnl(i).value = True Then Exit For
    Next i
    SaveCon "PswErr", CStr(i)
    SaveCon "PswLarge", txtCnt.Text
    If txtSec.Enabled Then SaveCon "PswWait", txtSec.Text
    SaveCon "Key", txtKey.Text
    SaveCon "Txt", txtLTxt.Text
    SaveCon "BGPic", filPic.FileName
    SaveCon "Scr", CStr(Abs(CInt(swiScr.Checked)))
    SaveCon "ScrWait", txtScr.Text
    SaveCon "ExitPsw", CStr(Abs(CInt(swiExit.Checked)))
    If swiDesktop.Checked Then
        mShellLnk "立即使用 Lock Pro 锁定计算机", MyPath & App.EXEName & ".exe", MyPath, App.EXEName & ".exe", "Lock", "", "立即使用 Lock Pro 锁定你的计算机"
    End If
    If swiAuto.Checked Then
        ChangeReg HKEY_LOCAL_MACHINE, "SOFTWARE\" & _
            "Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", Environ("Windir") & "\system32\userinit.exe," & _
            MyPath & App.EXEName & ".exe Lock"
    End If
    SaveCon "UNR", CStr(Abs(CInt(swiBlue.Checked)))
    Dim sLastRemo As String                                     '防止重复开关
    sLastRemo = ReadCon("Remolock") & ReadCon("Port")
    SaveCon "Remolock", CStr(Abs(CInt(swiRemo.Checked)))
    SaveCon "Port", txtPort.Text
    If ReadCon("Remolock") & ReadCon("Port") <> sLastRemo Then frmTray.SetSocket
    If bClear Then SavePsw "Remolock", ""
    '===========EndSave===========
    Unload Me
End Sub

Private Sub btnRfsh_Click()
    On Error Resume Next
    Dim fso As Object, disks As Object, disk As Object, ID As Object
    cboUSB.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set disks = fso.Drives
    For Each disk In disks
        Set ID = fso.GetDrive(fso.GetDriveName(disk))
        If ID.drivetype = 1 And disk.IsReady = True Then
            cboUSB.AddItem ID.DriveLetter & ":\"
        End If
    Next
    cboUSB.ListIndex = 0
End Sub

Private Sub filPic_Click()
    On Error GoTo PicErr
    imgPic.Picture = LoadPicture(MyPath & "LockPicture\" & filPic.FileName)
    Exit Sub
PicErr:
    imgPic.Picture = LoadPicture()
End Sub

Private Sub Form_Load()
    On Error Resume Next
    lblCap.Top = (480 - lblCap.Height) / 2
    picFrm(0).ZOrder 0
    filPic.Path = MyPath & "LockPicture\"
    btnRfsh_Click
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    '========LoadSettings========
    If ReadPsw("Psw") <> "" Then picOldPsw.Visible = True
    If ReadPsw("USB") <> "" Then
        cboUSB.Enabled = False
        lblUSB.Visible = True
    End If
    
    If ReadCon("First") = 0 Then
        swiUnl(0).Checked = Not (Int(ReadCon("Psw")) = 1)
        swiUnl(1).Checked = Not (Int(ReadCon("Psw")) = 0)
        txtSec.Text = ReadCon("PswWait")
        optUnl_Click Int(ReadCon("PswErr"))
        optUnl(Int(ReadCon("PswErr"))).value = True
        txtCnt.Text = ReadCon("PswLarge")
        Dim i As Long
        For i = 0 To filPic.ListCount - 1
            If filPic.List(i) = ReadCon("BGPic") Then filPic.ListIndex = i: Exit For
        Next i
        filPic_Click
        txtKey.Text = ReadCon("Key")
        txtLTxt.Text = ReadCon("Txt")
        swiScr.Checked = Int(ReadCon("Scr"))
        txtScr.Text = ReadCon("ScrWait")
        swiExit.Checked = Int(ReadCon("ExitPsw"))
        swiBlue.Checked = Int(ReadCon("UNR"))
        swiRemo.Checked = Int(ReadCon("Remolock"))
        txtPort.Text = ReadCon("Port")
        lblShow(33).Caption = GetLicenseNum & " 个设备已被授权"
    End If
    '=============End============
    txtPort_Change
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
    If ReadCon("First") = 1 Then
        frmTray.CloseLP
    Else
        If frmTray.Enabled = False Then frmTray.Enabled = True
    End If
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub lblLst_Click(Index As Integer)
    picFrm(Index).ZOrder 0
    picSel.Tag = Index
    With tmrSLst
        .Tag = Index
        .Enabled = True
    End With
End Sub

Private Sub lblLst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    With tmrSLst
        .Tag = Index
        .Enabled = True
    End With
End Sub

Private Sub lblShow_Click(Index As Integer)
    If Index = 15 Then
        ShellExecute 0, "open", MyPath & "LockPicture\", "", "", 1
    End If
End Sub

Private Sub lblUSB_Click()
    If CheckUSB = 1 Then cboUSB.Enabled = True: lblUSB.Visible = False
End Sub

Private Sub optUnl_Click(Index As Integer)
    If Index = 0 Then txtSec.Enabled = True Else txtSec.Enabled = False
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

Private Sub pngMouse_MouseExit()
    With picSel
        If .Top <> lblLst(CInt(.Tag)).Top - 120 Then
            tmrSLst.Tag = .Tag
            tmrSLst.Enabled = True
        End If
    End With
End Sub

Private Sub swiScr_Switched()
    txtScr.Enabled = swiScr.Checked
End Sub

Private Sub swiUnl_Switched(Index As Integer)
    If swiUnl(0).Checked = False And swiUnl(1).Checked = False Then
        swiUnl(Abs(Index - 1)).Checked = True
    End If
End Sub

Private Sub tmrSLst_Timer()
    With picSel
        .Top = .Top + GetMoveNum(lblLst(CInt(tmrSLst.Tag)).Top - 120, .Top, 8)
        If GetMoveNum(lblLst(CInt(tmrSLst.Tag)).Top - 120, .Top, 8) = 0 Then
            .Top = lblLst(CInt(tmrSLst.Tag)).Top - 120
            tmrSLst.Enabled = False
        End If
    End With
End Sub

Sub ShowErr(sErr As String)
    lblErr.Caption = sErr
    lblErr.Visible = True
End Sub

Sub ChangeReg(lKey As Long, sPath As String, sVal As String, sData As String)
    On Error GoTo crErr
    Dim lpVer As OSVERSIONONFO
    lpVer.dwOSVersionInfoSize = Len(lpVer)
    GetVersionEx lpVer
    If lpVer.dwMajorVersion >= 6 Then
        ShellExecute 0, "runas", MyPath & App.EXEName & ".exe", "Reg@@" & _
            CStr(lKey) & "@@" & sPath & "@@" & sVal & "@@" & sData, "", 1
    Else
        Shell MyPath & App.EXEName & ".exe Reg@@" & _
            CStr(lKey) & "@@" & sPath & "@@" & sVal & "@@" & sData, 1
    End If
    Exit Sub
crErr:
End Sub

Private Sub txtLTxt_Change()
    If Len(txtLTxt.Text) > 14 Then
        txtLTxt.Text = Left(txtLTxt.Text, 14)
        txtLTxt.SelStart = 14
        Beep
    End If
End Sub

Private Sub txtPort_Change()
    lblShow(29).Caption = "若要使用 Remolock 远程功能，请在 PC 或手机浏览器访问：" & _
            "http://" & GetMyIP & IIf(txtPort.Text = "80", "", ":" & txtPort.Text) & "/ 。"
End Sub
