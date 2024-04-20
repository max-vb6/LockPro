VERSION 5.00
Begin VB.UserControl ucBtn 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ScaleHeight     =   885
   ScaleWidth      =   1830
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00808080&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
   Begin LP_Updt.PngImage pngMs 
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   847
      Image           =   "ucBtn.ctx":0000
      Scaler          =   1
      Opacity         =   0
      OLEdrop         =   1
      Props           =   4
   End
   Begin LP_Updt.PngImage pngBtn 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1058
      Image           =   "ucBtn.ctx":04C7
      Scaler          =   1
      OLEdrop         =   1
      Props           =   4
   End
   Begin LP_Updt.PngImage pngBg 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1058
      Image           =   "ucBtn.ctx":098E
      Scaler          =   1
      OLEdrop         =   1
      Props           =   5
   End
End
Attribute VB_Name = "ucBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Click()
Event DblClick()

Private Sub lblCap_Click()
    pngMs_Click 1
End Sub

Private Sub lblCap_DblClick()
    pngMs_DblClick 1
End Sub

Private Sub lblCap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngMs_MouseDown Button, Shift, x, y
End Sub

Private Sub lblCap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngMs_MouseUp Button, Shift, x, y
End Sub

Private Sub pngMs_Click(ByVal Button As Integer)
    RaiseEvent Click
End Sub

Private Sub pngMs_DblClick(ByVal Button As Integer)
    RaiseEvent DblClick
End Sub

Private Sub pngMs_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngMs.FadeInOut 0, 30
    pngBtn.FadeInOut 0, 30
End Sub

Private Sub pngMs_MouseEnter()
    pngMs.FadeInOut 100, 15
End Sub

Private Sub pngMs_MouseExit()
    pngMs.FadeInOut 0
End Sub

Private Sub pngMs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pngBtn.FadeInOut 100
End Sub

Private Sub UserControl_Resize()
    pngBtn.Move 0, 0, UserControl.ScaleWidth + 20, UserControl.ScaleHeight
    pngMs.Move 0, 0, pngBtn.Width, pngBtn.Height
    pngBg.Move 0, 0, pngBtn.Width, pngBtn.Height
    shpBorder.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblCap.Move (pngBtn.Width - lblCap.Width) / 2, (pngBtn.Height - lblCap.Height) / 2
End Sub

Public Property Get Caption() As String
    Caption = lblCap.Caption
End Property

Public Property Let Caption(ByVal nCap As String)
    PropertyChanged "Caption"
    lblCap.Caption = nCap
    UserControl_Resize
End Property

Public Property Get FontSize() As Single
    FontSize = lblCap.FontSize
End Property

Public Property Let FontSize(ByVal lSiz As Single)
    PropertyChanged "FontSize"
    lblCap.FontSize = lSiz
    UserControl_Resize
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblCap.Caption = PropBag.ReadProperty("Caption", "")
    lblCap.FontSize = PropBag.ReadProperty("FontSize", lblCap.FontSize)
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", lblCap.Caption
    PropBag.WriteProperty "FontSize", lblCap.FontSize
End Sub

