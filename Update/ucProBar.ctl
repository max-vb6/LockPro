VERSION 5.00
Begin VB.UserControl ucProBar 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ScaleHeight     =   1890
   ScaleWidth      =   5235
   Begin VB.Timer tmrSPro 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblPro 
      BackColor       =   &H00008000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ucProBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrSPro_Timer()
    If tmrSPro.Tag = "" Then tmrSPro.Enabled = False
    With lblPro
        .Width = .Width + GetMoveNum(CSng(tmrSPro.Tag), .Width, 3)
        If GetMoveNum(CSng(tmrSPro.Tag), .Width, 3) = 0 Then
            .Width = CSng(tmrSPro.Tag)
            tmrSPro.Tag = ""
            tmrSPro.Enabled = False
        End If
    End With
End Sub

Private Sub UserControl_Initialize()
    lblPro.Left = -50
    lblPro.Width = 0
End Sub

Sub ChangePro(cPro As Long, cMax As Long)
    If cMax <= 0 Or cPro <= 0 Then
        tmrSPro.Tag = "50"
    Else
        tmrSPro.Tag = CStr((cPro / cMax) * UserControl.ScaleWidth + 50)
    End If
    tmrSPro.Enabled = True
End Sub
