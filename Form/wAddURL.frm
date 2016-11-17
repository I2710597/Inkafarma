VERSION 5.00
Begin VB.Form wAddURL 
   BorderStyle     =   0  'None
   Caption         =   "Ingresar URL"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wAddURL.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton BtnCancel 
      Height          =   360
      Left            =   2985
      TabIndex        =   5
      Top             =   2280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Picture         =   "wAddURL.frx":3958A
   End
   Begin SistemaInkaFarma.mButton BtnOK 
      Height          =   360
      Left            =   4305
      TabIndex        =   4
      Top             =   2280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Picture         =   "wAddURL.frx":40DDC
   End
   Begin SistemaInkaFarma.mButton BtnClose 
      Height          =   255
      Left            =   5220
      TabIndex        =   3
      ToolTipText     =   "Cerrar"
      Top             =   15
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   450
      Picture         =   "wAddURL.frx":4862E
   End
   Begin VB.Timer tDetect 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   345
      Top             =   2235
   End
   Begin VB.TextBox txtURL 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      HideSelection   =   0   'False
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   0
      Top             =   1290
      Width           =   4035
   End
   Begin SistemaInkaFarma.ucImage iconLock 
      Height          =   720
      Left            =   300
      Top             =   1050
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      bvData          =   "wAddURL.frx":4B320
      bData           =   -1  'True
      Filename        =   "Icon_IE.png"
      eScale          =   0
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
   Begin VB.Label TopTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "%FormTitle%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   5025
   End
   Begin VB.Label lbText 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese la direeción URL que desea abrir."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   780
      Width           =   4320
   End
End
Attribute VB_Name = "wAddURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    tDetect = False
    Me.Hide
    OpenEXE txtURL.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = 397 * 15: Me.Height = 197 * 15
    Redondear Me, 9
    TopTitle.Caption = Me.Caption
    tDetect = True
End Sub

Private Sub BtnCancel_Click()
    tDetect = False
    Unload Me
End Sub

Private Sub tDetect_Timer()
    If txtURL.Text <> "" Then
        BtnOK.Enabled = True
    Else
        BtnOK.Enabled = False
    End If
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
    If BtnOK.Visible = True Then
        If KeyAscii = vbKeyReturn Then
            btnOK_Click
        End If
    End If
End Sub

Private Sub TopTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MoverObjeto Me
    End If
End Sub
