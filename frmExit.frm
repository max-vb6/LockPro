VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   " ‰»Î√‹¬Î"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.Timer tmrFocus 
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin LockPro.ucBtn btnCancel 
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "»°œ˚"
      FontSize        =   10.5
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.PictureBox picPsw 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   2040
         ScaleHeight     =   2535
         ScaleWidth      =   3615
         TabIndex        =   4
         Top             =   0
         Width           =   3615
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
            Left            =   120
            PasswordChar    =   "l"
            TabIndex        =   6
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "√‹¬Î¥ÌŒÛ£¨«Î÷ÿ ‘"
            BeginProperty Font 
               Name            =   "Œ¢»Ì—≈∫⁄"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "»Ù“™Ω‚À¯£¨«Î ‰»ÎΩ‚À¯√‹¬Î"
            BeginProperty Font 
               Name            =   "Œ¢»Ì—≈∫⁄"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   2880
         End
      End
      Begin LockPro.PngImage pngLogo 
         Height          =   1800
         Left            =   240
         Top             =   360
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   3175
         Image           =   "frmExit.frx":000C
      End
   End
   Begin LockPro.ucBtn btnOK 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "»∑∂®"
      FontSize        =   10.5
   End
   Begin LockPro.PngImage pngCls 
      Height          =   495
      Left            =   5160
      ToolTipText     =   "πÿ±’"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   -2147483633
      Image           =   "frmExit.frx":74E1
      Opacity         =   0
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ‰»Î√‹¬Î"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
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
      Top             =   80
      Width           =   840
   End
   Begin VB.Image imgCls 
      Height          =   480
      Left            =   5160
      Picture         =   "frmExit.frx":7977
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shdw As cShadow

Sub ChangeMode(lMode As Long)
    Select Case lMode
        Case 0
            lblShow(0).Caption = "»Ù“™ÕÀ≥ˆ£¨«Î ‰»ÎΩ‚À¯√‹¬Î"
        Case 1
            lblShow(0).Caption = "Õ£÷πº∆ ±£¨«Î ‰»ÎΩ‚À¯√‹¬Î"
    End Select
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    If XMD5(txtPsw.Text) = ReadPsw("Psw") Then
        If InStr(lblShow(0).Caption, "ÕÀ≥ˆ") <> 0 Then
            frmTray.CloseLP
        Else
            frmTimer.StopTimer
            Unload Me
        End If
    Else
        lblShow(1).Visible = True
        Beep
    End If
End Sub

Private Sub Form_Load()
    lblCap.Top = (480 - lblCap.Height) / 2
    Set Shdw = New cShadow
    With Shdw
        .Transparency = 120
        .Depth = 10
        .Shadow Me
    End With
    txtPsw.Text = ""
    lblShow(0).ForeColor = vbBlack
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

Private Sub tmrFocus_Timer()
    On Error Resume Next
    txtPsw.SetFocus
    tmrFocus.Enabled = False
End Sub

Private Sub txtPsw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnOK_Click
    End If
End Sub
