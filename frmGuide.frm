VERSION 5.00
Begin VB.Form frmGuide 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Lock Pro ��"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picStep 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   6120
      ScaleHeight     =   855
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   480
      Width           =   1095
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/9"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   690
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   6
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   30
      Top             =   480
      Width           =   7215
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remolock"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   17
         Left            =   3120
         TabIndex        =   32
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGuide.frx":0000
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1455
         Index           =   16
         Left            =   3120
         TabIndex        =   31
         Top             =   1680
         Width           =   3780
      End
      Begin LockPro.PngImage pngPic 
         Height          =   1815
         Index           =   6
         Left            =   240
         Top             =   1080
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   3201
         Image           =   "frmGuide.frx":008D
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin VB.Timer tmrSFrm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   0
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   0
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Pro ��һ�������������������򵼽���������ʶ Lock Pro �Ļ�������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   915
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Top             =   1800
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ã��ҽ� Lock Pro��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   1200
         Width           =   3735
      End
      Begin LockPro.PngImage pngLogo 
         Height          =   2040
         Index           =   0
         Left            =   480
         Top             =   960
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   3598
         Image           =   "frmGuide.frx":CCCB
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   7
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   26
      Top             =   480
      Width           =   7215
      Begin LockPro.PngImage pngPic 
         Height          =   1560
         Index           =   5
         Left            =   240
         Top             =   1200
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   2752
         Image           =   "frmGuide.frx":141A0
         OLEdrop         =   1
         Props           =   5
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGuide.frx":277EA
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1455
         Index           =   15
         Left            =   3120
         TabIndex        =   28
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݵ���"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   14
         Left            =   3120
         TabIndex        =   27
         Top             =   1080
         Width           =   1440
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   5
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   20
      Top             =   480
      Width           =   7215
      Begin LockPro.PngImage pngPic 
         Height          =   2115
         Index           =   4
         Left            =   360
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3731
         Image           =   "frmGuide.frx":2787A
         OLEdrop         =   1
         Props           =   5
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ϊ��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   11
         Left            =   3120
         TabIndex        =   22
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Pro �Դ������������������Ż�������İ�ť����ʹ���Ĵ�����������ʡ�����ḻ���л�Ч�������� Naree ������ÿһ�ε��������Ѥ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1455
         Index           =   10
         Left            =   3120
         TabIndex        =   21
         Top             =   1680
         Width           =   3780
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   4
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   17
      Top             =   480
      Width           =   7215
      Begin LockPro.PngImage pngPic 
         Height          =   2040
         Index           =   3
         Left            =   360
         Top             =   960
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   3598
         Image           =   "frmGuide.frx":2F5BD
         OLEdrop         =   1
         Props           =   5
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "ͨ���˵������Կ��� Lock Pro ���á�����ṩ�˷ḻ��ѡ����������������ĸ�����������ѡ�����ò�����Lock Pro ���ڴ������½���ʾ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1455
         Index           =   9
         Left            =   3120
         TabIndex        =   19
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ������� Lock Pro"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   8
         Left            =   3120
         TabIndex        =   18
         Top             =   1080
         Width           =   3375
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   3
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   14
      Top             =   480
      Width           =   7215
      Begin LockPro.PngImage pngPic 
         Height          =   1530
         Index           =   2
         Left            =   360
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2699
         Image           =   "frmGuide.frx":3E631
         OLEdrop         =   1
         Props           =   5
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmGuide.frx":7CC3F
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1455
         Index           =   7
         Left            =   3120
         TabIndex        =   16
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   6
         Left            =   3120
         TabIndex        =   15
         Top             =   1080
         Width           =   1440
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   2
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   11
      Top             =   480
      Width           =   7215
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "˫�� Lock Pro ������ɫ��ť��������������Ļ����Ҳ���Ե���Ҳ�Ĳ˵���ť�������úõĿ�ݼ�������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1515
         Index           =   5
         Left            =   3000
         TabIndex        =   13
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Top             =   1080
         Width           =   1440
      End
      Begin LockPro.PngImage pngPic 
         Height          =   1665
         Index           =   1
         Left            =   240
         Top             =   1200
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   2937
         Image           =   "frmGuide.frx":7CCCD
         OLEdrop         =   1
         Props           =   5
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   8
      Top             =   480
      Width           =   7215
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "Lock Pro ����ʱ������Ļ������ʾ Lock Pro ����Lock Pro �����Զ����أ��������ק��ɫ�İ�ť���Ըı� Lock Pro ����λ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1275
         Index           =   3
         Left            =   3000
         TabIndex        =   10
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ Lock Pro ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   2
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   2760
      End
      Begin LockPro.PngImage pngPic 
         Height          =   2220
         Index           =   0
         Left            =   240
         Top             =   840
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   3916
         Image           =   "frmGuide.frx":90980
         OLEdrop         =   1
      End
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   8
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7215
      TabIndex        =   23
      Top             =   480
      Width           =   7215
      Begin LockPro.PngImage pngLogo 
         Height          =   2040
         Index           =   2
         Left            =   480
         Top             =   960
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   3598
         Image           =   "frmGuide.frx":9A8FD
      End
      Begin VB.Label lblFirst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˳���֮����Զ���������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   2880
         TabIndex        =   29
         Top             =   2880
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ܸ�����ʶ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   13
         Left            =   2880
         TabIndex        =   25
         Top             =   1080
         Width           =   2160
      End
      Begin VB.Label lblShow 
         BackStyle       =   0  'Transparent
         Caption         =   "���Ѿ�������ʶ�� Lock Pro�����������Ծ������� Lock Pro ���ڶ๦����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1335
         Index           =   12
         Left            =   2880
         TabIndex        =   24
         Top             =   1680
         Width           =   3540
      End
   End
   Begin LockPro.PngImage pngCls 
      Height          =   495
      Left            =   6720
      ToolTipText     =   "�ر�"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmGuide.frx":A1DD2
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblCnl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˳���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3000
      TabIndex        =   2
      Top             =   4560
      Width           =   1200
   End
   Begin LockPro.PngImage pngBack 
      Height          =   585
      Left            =   240
      Top             =   4480
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1032
      Image           =   "frmGuide.frx":A2268
      Mirror          =   1
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngNext 
      Height          =   585
      Left            =   6360
      Top             =   4485
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1032
      Image           =   "frmGuide.frx":A32C5
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Pro ʹ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Width           =   1770
   End
   Begin VB.Image imgCls 
      Height          =   480
      Left            =   6720
      Picture         =   "frmGuide.frx":A4322
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lNow As Long
Dim Shdw As cShadow

Private Sub Form_Load()
    lblCap.Top = (480 - lblCap.Height) / 2
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    lNow = 0
    SetZOrder picFrm(0)
    lblFirst.Visible = CBool(ReadCon("First"))
    pngLogo(0).Opacity = 0
    pngLogo(0).FadeInOut 100, 3
    pngBack.Opacity = 0
    pngNext.Opacity = 0
    pngNext.FadeInOut 100, 3
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
    If ReadCon("First") = "1" Then frmSettings.Show
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub

Private Sub lblCnl_Click()
    Unload Me
End Sub

Private Sub pngBack_Click(ByVal Button As Integer)
    If pngBack.Opacity <> 0 And lNow <= 1 Then pngBack.FadeInOut 0, 3
    If pngNext.Opacity < 100 And lNow = 8 Then pngNext.FadeInOut 100, 3
    If lNow = 0 Then
        Exit Sub
    End If
    lNow = lNow - 1
    lblStep(0).Caption = lNow + 1
    picFrm(lNow).Left = -picFrm(0).Width
    SetZOrder picFrm(lNow)
    tmrSFrm.Tag = "Back"
    tmrSFrm.Enabled = True
End Sub

Private Sub pngBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lNow = 0 Then Exit Sub
    pngBack.FadeInOut 80, 5
End Sub

Private Sub pngBack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lNow = 0 Then Exit Sub
    pngBack.FadeInOut 100, 2
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

Private Sub pngNext_Click(ByVal Button As Integer)
    If pngNext.Opacity <> 0 And lNow >= 7 Then pngNext.FadeInOut 0, 3
    If pngBack.Opacity < 100 And lNow >= 0 Then pngBack.FadeInOut 100, 3
    If lNow = 8 Then
        Exit Sub
    End If
    lNow = lNow + 1
    lblStep(0).Caption = lNow + 1
    picFrm(lNow).Left = picFrm(0).Width
    SetZOrder picFrm(lNow)
    tmrSFrm.Tag = ""
    tmrSFrm.Enabled = True
End Sub

Private Sub pngNext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lNow >= 7 Then Exit Sub
    pngNext.FadeInOut 80, 5
End Sub

Private Sub pngNext_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lNow >= 7 Then Exit Sub
    pngNext.FadeInOut 100, 2
End Sub

Sub SetZOrder(zCtrl As Control)
    zCtrl.ZOrder 0
    picStep.ZOrder 0
End Sub

Private Sub tmrSFrm_Timer()
    With picFrm(CInt(lNow))
        .Left = .Left + GetMoveNum(0, .Left, 5)
        If tmrSFrm.Tag = "" Then
            picFrm(CInt(lNow) - 1).Left = .Left - .Width
        Else
            picFrm(CInt(lNow) + 1).Left = .Left + .Width
        End If
        If GetMoveNum(0, .Left, 5) = 0 Then
            .Left = 0
            tmrSFrm.Enabled = False
        End If
    End With
End Sub
