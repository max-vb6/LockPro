VERSION 5.00
Begin VB.UserControl ucSwitch 
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   ScaleHeight     =   990
   ScaleWidth      =   1425
   Begin VB.PictureBox picSwi 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Width           =   495
      Begin VB.Shape shpBrd 
         BorderColor     =   &H00C0C0C0&
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Timer tmrSwi 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox picPro 
      Appearance      =   0  'Flat
      BackColor       =   &H00D7D7D7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.Shape shpBrd 
         BorderColor     =   &H00C0C0C0&
         Height          =   375
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   975
      End
      Begin LockPro.PngImage pngShdw 
         Height          =   615
         Index           =   1
         Left            =   120
         Top             =   0
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   1085
         Image           =   "ucSwitch.ctx":0000
         Mirror          =   1
         OLEdrop         =   1
         Props           =   5
      End
      Begin LockPro.PngImage pngShdw 
         Height          =   615
         Index           =   0
         Left            =   480
         Top             =   0
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   1085
         Image           =   "ucSwitch.ctx":0BA8
         OLEdrop         =   1
         Props           =   5
      End
      Begin VB.Label lblPro 
         BackColor       =   &H00FAC8A5&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "ucSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Switched()

Dim SwiSta As Boolean

Public Property Get Checked() As Boolean
    Checked = SwiSta
End Property

Public Property Let Checked(ByVal nChk As Boolean)
    PropertyChanged "Checked"
    Switcher nChk
End Property

Private Sub lblPro_Click()
    Switcher False
End Sub

Private Sub picPro_Click()
    Switcher True
End Sub

Private Sub picSwi_Click()
    With picSwi
        If .Left = 0 Or .Left = UserControl.ScaleWidth - .Width Then Switcher (Not SwiSta): Exit Sub
    End With
End Sub

Private Sub picSwi_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Static ox!
    With picSwi
        If Button = 1 Then
            If .Left - ox + x < 0 Then
                .Left = 0
                picPro.Refresh
                UserControl_Resize
                Exit Sub
            ElseIf .Left - ox + x > UserControl.ScaleWidth - picSwi.Width Then
                .Left = UserControl.ScaleWidth - picSwi.Width
                picPro.Refresh
                UserControl_Resize
                Exit Sub
            Else
                .Move .Left - ox + x
                UserControl_Resize
            End If
        Else
            ox = x
        End If
    End With
End Sub

Private Sub picSwi_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With picSwi
        If .Left <= picPro.Width / 4 Then
            Switcher False
        ElseIf .Left >= picPro.Width / 4 Then
            Switcher True
        End If
    End With
End Sub

Private Sub pngShdw_Click(Index As Integer, ByVal Button As Integer)
    If Index = 0 Then
        Switcher True
    Else
        Switcher False
    End If
End Sub

Private Sub tmrSwi_Timer()
    With picSwi
        If tmrSwi.Tag = "" Then
            .Left = .Left + GetMoveNum(0, .Left, 5)
            UserControl_Resize
            If GetMoveNum(0, .Left, 5) = 0 Then .Left = 0: tmrSwi.Enabled = False
        Else
            .Left = .Left + GetMoveNum(UserControl.ScaleWidth - picSwi.Width, .Left, 5)
            UserControl_Resize
            If GetMoveNum(UserControl.ScaleWidth - picSwi.Width, .Left, 5) = 0 Then
                .Left = UserControl.ScaleWidth - picSwi.Width
                tmrSwi.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Switcher PropBag.ReadProperty("Checked", "")
End Sub

Private Sub UserControl_Resize()
    On Error GoTo rsErr
    With UserControl
        .Width = picPro.Width
        .Height = picPro.Height
        pngShdw(0).Left = picSwi.Left + picSwi.Width - 10
        pngShdw(1).Left = picSwi.Left - pngShdw(1).Width
        lblPro.Width = picSwi.Left + 120
    End With
    Exit Sub
rsErr:
    picSwi.Left = 10
    Switcher False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Checked", SwiSta
End Sub

Sub Switcher(bSta As Boolean)
    With tmrSwi
        If bSta Then
            .Tag = "On"
        Else
            .Tag = ""
        End If
        .Enabled = True
    End With
    UserControl_Resize
    SwiSta = bSta
    RaiseEvent Switched
End Sub

