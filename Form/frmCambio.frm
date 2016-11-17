VERSION 5.00
Begin VB.Form frmCambio 
   BorderStyle     =   0  'None
   Caption         =   "Tipo de cambio"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCambio.frx":0000
   ScaleHeight     =   3645
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton BClose 
      Height          =   255
      Left            =   5220
      TabIndex        =   6
      ToolTipText     =   "Cerrar"
      Top             =   15
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   450
      Picture         =   "frmCambio.frx":4C648
   End
   Begin SistemaInkaFarma.mButton cmdCancelar 
      Height          =   360
      Left            =   4200
      TabIndex        =   5
      Top             =   2295
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Picture         =   "frmCambio.frx":4F33A
   End
   Begin SistemaInkaFarma.mButton cmdAceptar 
      Height          =   360
      Left            =   2820
      TabIndex        =   4
      Top             =   2295
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Picture         =   "frmCambio.frx":56B8C
   End
   Begin VB.TextBox txtCambio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1485
      Width           =   1335
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   765
      Width           =   1335
   End
   Begin SistemaInkaFarma.ucImage Conversion 
      Height          =   1590
      Left            =   240
      Top             =   480
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   2805
      bvData          =   "frmCambio.frx":5E3DE
      bData           =   -1  'True
      Filename        =   "cdfv.png"
      eScale          =   1
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   435
      TabIndex        =   7
      Top             =   75
      Width           =   1380
   End
   Begin SistemaInkaFarma.ucImage ucImage1 
      Height          =   270
      Left            =   135
      Top             =   75
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   476
      bvData          =   "frmCambio.frx":6628B
      bData           =   -1  'True
      Filename        =   "Accounting.png"
      eScale          =   1
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de cambio"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2190
      TabIndex        =   2
      Top             =   1485
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   2220
      X2              =   5325
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Cotización"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2205
      TabIndex        =   0
      Top             =   765
      Width           =   1710
   End
End
Attribute VB_Name = "frmCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BClose_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    If Val(txtCambio.Text) = 0 Then MsgBox "Tipo de cambio incorrecto...?", vbOKOnly + vbCritical, _
    "Información al usuario...!": txtCambio.Text = "": txtCambio.SetFocus: Exit Sub
    cmdCancelar_Click
End Sub

Private Sub cmdCancelar_Click()
    If Val(txtCambio.Text) = 0 Then MsgBox "Tipo de cambio incorrecto...?", vbOKOnly + vbCritical, _
    "Información al usuario...!": txtCambio.Text = "": txtCambio.SetFocus: Exit Sub
        cambio = txtCambio.Text
        Unload frmCambio
        Load wMain
        wMain.Show
End Sub

Private Sub Form_Load()
    txtFecha.Text = Format(Date, "dd/mm/yyyy")
    If cambio = 0 Then
        txtCambio.Text = "3.50"
    Else
        txtCambio.Text = cambio
    End If
    '////////////////////////////////////////
    Me.Width = 397 * 15: Me.Height = 197 * 15
    Redondear Me, 9
End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar.SetFocus
End Sub
