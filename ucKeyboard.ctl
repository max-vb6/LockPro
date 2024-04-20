VERSION 5.00
Begin VB.UserControl ucKeyboard 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   ScaleHeight     =   4935
   ScaleWidth      =   11025
   Begin VB.Timer tmrCnt 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrBk 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin LockPro.PngImage pngShift 
      Height          =   960
      Index           =   1
      Left            =   240
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":0000
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngEnd 
      Height          =   960
      Left            =   8040
      Top             =   3240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":0F04
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Êý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   26
      Left            =   0
      TabIndex        =   26
      Top             =   3480
      Width           =   975
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   26
      Left            =   0
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":1E91
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   25
      Left            =   6960
      TabIndex        =   25
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   24
      Left            =   6000
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   23
      Left            =   5040
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   22
      Left            =   4080
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   21
      Left            =   3120
      TabIndex        =   21
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   20
      Left            =   2160
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   19
      Left            =   1200
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   18
      Left            =   8160
      TabIndex        =   18
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   17
      Left            =   7200
      TabIndex        =   17
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   16
      Left            =   6240
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   15
      Left            =   5280
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   14
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   13
      Left            =   3360
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   12
      Left            =   2400
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   11
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   9
      Left            =   8640
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   8
      Left            =   7680
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   7
      Left            =   6720
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   6
      Left            =   5760
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   4
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblAsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin LockPro.PngImage pngSpc 
      Height          =   960
      Left            =   1200
      Top             =   3240
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":2B3C
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngBk 
      Height          =   960
      Left            =   8040
      Top             =   2280
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":3886
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngShift 
      Height          =   960
      Index           =   0
      Left            =   240
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":4833
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   25
      Left            =   6960
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":578B
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   24
      Left            =   6000
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":6436
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   23
      Left            =   5040
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":70E1
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   22
      Left            =   4080
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":7D8C
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   21
      Left            =   3120
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":8A37
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   20
      Left            =   2160
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":96E2
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   19
      Left            =   1200
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":A38D
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   18
      Left            =   8160
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":B038
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   17
      Left            =   7200
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":BCE3
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   16
      Left            =   6240
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":C98E
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   15
      Left            =   5280
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":D639
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   14
      Left            =   4320
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":E2E4
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   13
      Left            =   3360
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":EF8F
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   12
      Left            =   2400
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":FC3A
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   11
      Left            =   1440
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":108E5
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   10
      Left            =   480
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":11590
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngShdw 
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   503
      Image           =   "ucKeyboard.ctx":1223B
      Scaler          =   1
      OLEdrop         =   1
      Props           =   4
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   9
      Left            =   8640
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":12DA9
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   8
      Left            =   7680
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":13A54
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   7
      Left            =   6720
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":146FF
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   6
      Left            =   5760
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":153AA
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   5
      Left            =   4800
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":16055
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   4
      Left            =   3840
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":16D00
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   3
      Left            =   2880
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":179AB
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   2
      Left            =   1920
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":18656
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   1
      Left            =   960
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":19301
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngKey 
      Height          =   960
      Index           =   0
      Left            =   0
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1693
      Image           =   "ucKeyboard.ctx":19FAC
      OLEdrop         =   1
      Props           =   5
   End
   Begin LockPro.PngImage pngBg 
      Height          =   4095
      Left            =   0
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7223
      Image           =   "ucKeyboard.ctx":1AC57
      Scaler          =   1
      Opacity         =   50
      OLEdrop         =   1
      Props           =   5
   End
   Begin VB.Image imgBg 
      Height          =   4575
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "ucKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lBgTop As Long

Event KeyPressed(sKey As String)
Event LetterBacked()
Event InputFinish()

Private Sub lblAsc_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index < 26 Then
        tmrBk.Tag = Index
        tmrCnt.Enabled = True
    End If
    pngKey(Index).FadeInOut 70, 5
End Sub

Private Sub lblAsc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrCnt.Enabled = False
    tmrBk.Tag = ""
    tmrBk.Enabled = False
    If Index < 26 Then
        If pngShift(0).Visible = True Then
            RaiseEvent KeyPressed(LCase(lblAsc(Index).Caption))
        Else
            RaiseEvent KeyPressed(lblAsc(Index).Caption)
        End If
    Else
        If lblAsc(26).Caption = "Êý" Then
            pngShift(1).Visible = False
            pngShift(0).Visible = True
            lblAsc(0).Caption = "1"
            lblAsc(1).Caption = "2"
            lblAsc(2).Caption = "3"
            lblAsc(3).Caption = "4"
            lblAsc(4).Caption = "5"
            lblAsc(5).Caption = "6"
            lblAsc(6).Caption = "7"
            lblAsc(7).Caption = "8"
            lblAsc(8).Caption = "9"
            lblAsc(9).Caption = "0"
            lblAsc(10).Caption = "`"
            lblAsc(11).Caption = "-"
            lblAsc(12).Caption = "="
            lblAsc(13).Caption = "["
            lblAsc(14).Caption = "]"
            lblAsc(15).Caption = "\"
            lblAsc(16).Caption = ";"
            lblAsc(17).Caption = "'"
            lblAsc(18).Caption = ","
            lblAsc(19).Caption = ""
            lblAsc(20).Caption = ""
            lblAsc(21).Caption = "."
            lblAsc(22).Caption = ""
            lblAsc(23).Caption = "/"
            lblAsc(24).Caption = ""
            lblAsc(25).Caption = ""
            lblAsc(26).Caption = "Ó¢"
        Else
            pngShift(1).Visible = False
            pngShift(0).Visible = True
            lblAsc(0).Caption = "Q"
            lblAsc(1).Caption = "W"
            lblAsc(2).Caption = "E"
            lblAsc(3).Caption = "R"
            lblAsc(4).Caption = "T"
            lblAsc(5).Caption = "Y"
            lblAsc(6).Caption = "U"
            lblAsc(7).Caption = "I"
            lblAsc(8).Caption = "O"
            lblAsc(9).Caption = "P"
            lblAsc(10).Caption = "A"
            lblAsc(11).Caption = "S"
            lblAsc(12).Caption = "D"
            lblAsc(13).Caption = "F"
            lblAsc(14).Caption = "G"
            lblAsc(15).Caption = "H"
            lblAsc(16).Caption = "J"
            lblAsc(17).Caption = "K"
            lblAsc(18).Caption = "L"
            lblAsc(19).Caption = "Z"
            lblAsc(20).Caption = "X"
            lblAsc(21).Caption = "C"
            lblAsc(22).Caption = "V"
            lblAsc(23).Caption = "B"
            lblAsc(24).Caption = "N"
            lblAsc(25).Caption = "M"
            lblAsc(26).Caption = "Êý"
        End If
    End If
    
    pngKey(Index).FadeInOut 100, 5
End Sub

Private Sub pngBk_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
tmrBk.Tag = ""
tmrCnt.Enabled = True
pngBk.FadeInOut 70, 5
End Sub

Private Sub pngBk_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If pngShift(0).Visible = False Then pngShift_MouseUp 1, 1, 0, 0, 0
tmrCnt.Enabled = False
tmrBk.Tag = ""
tmrBk.Enabled = False
RaiseEvent LetterBacked
pngBk.FadeInOut 100, 5
End Sub

Private Sub pngEnd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
pngEnd.FadeInOut 70, 5
End Sub

Private Sub pngEnd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If pngShift(0).Visible = False Then pngShift_MouseUp 1, 1, 0, 0, 0
RaiseEvent InputFinish
pngEnd.FadeInOut 100, 5
End Sub

Private Sub pngKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
pngKey(Index).FadeInOut 70, 5
End Sub

Private Sub pngKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
lblAsc_MouseUp Index, 1, 0, 0, 0
End Sub

Private Sub pngShift_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
pngShift(Index).FadeInOut 70, 5
End Sub

Private Sub pngShift_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 0 Then
pngShift(0).Visible = False
pngShift(1).Visible = True
If lblAsc(26).Caption = "Ó¢" Then
lblAsc(0).Caption = "!"
lblAsc(1).Caption = "@"
lblAsc(2).Caption = "#"
lblAsc(3).Caption = "$"
lblAsc(4).Caption = "%"
lblAsc(5).Caption = "^"
lblAsc(6).Caption = "&"
lblAsc(7).Caption = "*"
lblAsc(8).Caption = "("
lblAsc(9).Caption = ")"
lblAsc(10).Caption = "~"
lblAsc(11).Caption = "_"
lblAsc(12).Caption = "+"
lblAsc(13).Caption = "{"
lblAsc(14).Caption = "}"
lblAsc(15).Caption = "|"
lblAsc(16).Caption = ":"
lblAsc(17).Caption = """"
lblAsc(18).Caption = "<"
lblAsc(19).Caption = ""
lblAsc(20).Caption = ""
lblAsc(21).Caption = ">"
lblAsc(22).Caption = ""
lblAsc(23).Caption = "?"
lblAsc(24).Caption = ""
lblAsc(25).Caption = ""
End If
Else
pngShift(1).Visible = False
pngShift(0).Visible = True
If lblAsc(26).Caption = "Ó¢" Then
lblAsc(0).Caption = "1"
lblAsc(1).Caption = "2"
lblAsc(2).Caption = "3"
lblAsc(3).Caption = "4"
lblAsc(4).Caption = "5"
lblAsc(5).Caption = "6"
lblAsc(6).Caption = "7"
lblAsc(7).Caption = "8"
lblAsc(8).Caption = "9"
lblAsc(9).Caption = "0"
lblAsc(10).Caption = "`"
lblAsc(11).Caption = "-"
lblAsc(12).Caption = "="
lblAsc(13).Caption = "["
lblAsc(14).Caption = "]"
lblAsc(15).Caption = "\"
lblAsc(16).Caption = ";"
lblAsc(17).Caption = "'"
lblAsc(18).Caption = ","
lblAsc(19).Caption = ""
lblAsc(20).Caption = ""
lblAsc(21).Caption = "."
lblAsc(22).Caption = ""
lblAsc(23).Caption = "/"
lblAsc(24).Caption = ""
lblAsc(25).Caption = ""
End If
End If

pngShift(Index).FadeInOut 100, 5
End Sub

Private Sub pngSpc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
pngSpc.FadeInOut 70, 5
End Sub

Private Sub pngSpc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If pngShift(0).Visible = False Then pngShift_MouseUp 1, 1, 0, 0, 0
RaiseEvent KeyPressed(" ")
pngSpc.FadeInOut 100, 5
End Sub

Private Sub tmrBk_Timer()
If tmrBk.Tag = "" Then
RaiseEvent LetterBacked
Else
If pngShift(0).Visible = True Then
RaiseEvent KeyPressed(LCase(lblAsc(CInt(tmrBk.Tag)).Caption))
Else
RaiseEvent KeyPressed(lblAsc(CInt(tmrBk.Tag)).Caption)
End If
End If
End Sub

Private Sub tmrCnt_Timer()
tmrBk.Enabled = True
tmrCnt.Enabled = False
End Sub

Private Sub UserControl_Resize()
Dim lWdt As Long, i As Long
With UserControl
imgBg.Move 0, lBgTop, .ScaleWidth, .ScaleHeight
pngShdw.Width = .ScaleWidth + 600
pngBg.Move 0, pngShdw.Height, .ScaleWidth + 600
.Height = pngShdw.Height + pngBg.Height
lWdt = (.ScaleWidth - 9615) / 2
For i = 0 To 9
pngKey(i).Left = i * 975 + lWdt
lblAsc(i).Left = pngKey(i).Left
Next i
For i = 10 To 18
pngKey(i).Left = (i - 10) * 975 + lWdt + 480
lblAsc(i).Left = pngKey(i).Left
Next i
For i = 19 To 25
pngKey(i).Left = (i - 19) * 975 + lWdt + 1200
lblAsc(i).Left = pngKey(i).Left
Next i
pngShift(0).Left = 240 + lWdt
pngShift(1).Left = 240 + lWdt
pngBk.Left = pngKey(25).Left + pngKey(25).Width + 120
pngKey(26).Left = lWdt
lblAsc(26).Left = lWdt
pngSpc.Left = pngKey(19).Left
pngEnd.Left = pngBk.Left
End With
End Sub

Sub SetKeyPic(sPic As StdPicture)
On Error Resume Next
imgBg.Picture = LoadPicture()
Set imgBg.Picture = sPic
End Sub

Sub ShowBd()
Dim lCnt As Long
For lCnt = 0 To 26
pngKey(lCnt).Opacity = 0
Next lCnt
pngBg.Opacity = 0
pngBk.Opacity = 0
pngEnd.Opacity = 0
pngSpc.Opacity = 0
pngShift(0).Opacity = 0
pngShdw.Opacity = 0

Do Until pngShdw.Opacity = 100
Sleep 10
For lCnt = 0 To 26
pngKey(lCnt).Opacity = pngKey(lCnt).Opacity + 10
Next lCnt
pngBk.Opacity = pngBk.Opacity + 10
pngBg.Opacity = IIf(pngBk.Opacity > 30, pngBk.Opacity - 30, 0)
pngEnd.Opacity = pngBk.Opacity
pngSpc.Opacity = pngBk.Opacity
pngShift(0).Opacity = pngBk.Opacity
pngShdw.Opacity = pngBk.Opacity
Loop
End Sub

Sub HideBd()
Dim lCnt As Long
Do Until pngShdw.Opacity = 0
Sleep 10
For lCnt = 0 To 26
pngKey(lCnt).Opacity = pngKey(lCnt).Opacity - 10
Next lCnt
pngBg.Opacity = pngBg.Opacity - 10
pngBk.Opacity = pngBg.Opacity
pngEnd.Opacity = pngBg.Opacity
pngSpc.Opacity = pngBg.Opacity
pngShift(0).Opacity = pngBg.Opacity
pngShdw.Opacity = pngBg.Opacity
Loop
End Sub

