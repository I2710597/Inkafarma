VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de los Integrantes."
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton BtnAceptar 
      Height          =   360
      Left            =   4320
      TabIndex        =   14
      Top             =   5880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      Picture         =   "About.frx":6852
   End
   Begin VB.TextBox InfoAuthor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3480
      Width           =   3750
   End
   Begin VB.TextBox InfoMailWeb 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   3750
   End
   Begin VB.TextBox InfoDesc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4200
      Width           =   3750
   End
   Begin VB.TextBox InfoCredits 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4920
      Width           =   3750
   End
   Begin VB.ComboBox Control 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3120
      Width           =   3750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vasquez Cueva, Jose Miguel."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serna Huaman, Viviana Patricia."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   16
      Top             =   1800
      Width           =   2550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coleto Tadeo, Flor Maximina."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   15
      Top             =   1320
      Width           =   2370
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Créditos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1020
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mail / WEB:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control / Módulo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1410
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Información de Controles y Módulos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2400
      X2              =   5520
      Y1              =   2640
      Y2              =   2640
   End
   Begin SistemaInkaFarma.ucImage ucImage2 
      Height          =   585
      Left            =   2400
      Top             =   480
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1032
      bvData          =   "About.frx":E0A4
      bData           =   -1  'True
      Filename        =   "Users 2.png"
      eScale          =   1
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cerna Huerta, Ever."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valverde Dulanto, Adam Carl."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Integrantes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   1065
   End
   Begin SistemaInkaFarma.ucImage ucImage1 
      Height          =   2130
      Left            =   45
      Top             =   480
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3757
      bvData          =   "About.frx":11144
      bData           =   -1  'True
      Filename        =   "LOG.png"
      eScale          =   0
      lContrast       =   0
      lBrightness     =   0
      lAlpha          =   100
      bGrayScale      =   0   'False
      lAngle          =   0
      bFlipH          =   0   'False
      bFlipV          =   0   'False
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnAceptar_Click()
    Unload Me
End Sub

Private Sub Control_Change()
    Select Case Control.ListIndex

    Case 0
        SetControlsInfo "SONIC88 - Marco A. Olivares Aracena", _
                        "sonic88@live.cl - maoa17@gmail.com", _
                        "Botón simple de uso común mediante el API BitBlt."

    Case 1
        SetControlsInfo "Cobein", _
                        "cobein27@hotmail.com", _
                        "Simple Image control replacement (Beta).", _
                        "LaVolpe, Paul Caton and http://www.activevb.de"
    Case 2
        SetControlsInfo "SONIC88 - Marco A. Olivares Aracena", _
                        "sonic88@live.cl", _
                        "Control para mostrar una animación simple mediante el API BitBlt."

    Case 3
        SetControlsInfo "SONIC88 - Marco A. Olivares Aracena", _
                        "sonic88@live.cl", _
                        "Es un control para simular un objeto o boton del escritorio de windows vista mediante el API BitBlt."

    Case 4
        SetControlsInfo "SONIC88 - Marco A. Olivares Aracena", _
                        "sonic88@live.cl", _
                        "Es un control para simular un boton de comando comun que en este caso se ocupara para el menu, ya que tiene una propiedad especial para eso la propiedad Subcheck."

    End Select
End Sub

Private Sub Control_Click()
    Control_Change
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    Control.AddItem "mButton", 0
    Control.AddItem "ucImage", 1
    Control.AddItem "mAniControl", 2
    Control.AddItem "BtnDesktop", 3
    Control.AddItem "BtnMenu", 4
    Control.ListIndex = 0
End Sub

Private Sub SetControlsInfo( _
    Optional ByVal mAuthor As String, _
    Optional ByVal mMailWeb As String, _
    Optional ByVal mDescription As String, _
    Optional ByVal mCredits As String)

    InfoAuthor.Text = mAuthor
    InfoMailWeb.Text = mMailWeb
    InfoDesc.Text = mDescription
    InfoCredits.Text = mCredits
End Sub

