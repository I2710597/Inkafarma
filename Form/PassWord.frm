VERSION 5.00
Begin VB.Form PassWord 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ContObj 
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   600
      Picture         =   "PassWord.frx":0000
      ScaleHeight     =   3405
      ScaleWidth      =   6360
      TabIndex        =   1
      Top             =   480
      Width           =   6360
      Begin VB.TextBox TxtPassWord 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1725
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2310
         Width           =   2415
      End
      Begin VB.ComboBox cmbUsuario 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "PassWord.frx":46BB8
         Left            =   1680
         List            =   "PassWord.frx":46BC2
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1845
         Width           =   2505
      End
      Begin SistemaInkaFarma.mButton cmdEntrar 
         Height          =   360
         Left            =   2625
         TabIndex        =   6
         Top             =   2850
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   635
         Picture         =   "PassWord.frx":46BDD
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         Height          =   360
         Left            =   1665
         Top             =   1815
         Width           =   2535
      End
      Begin SistemaInkaFarma.ucImage ucImage1 
         Height          =   1515
         Left            =   4500
         Top             =   1455
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   2672
         bvData          =   "PassWord.frx":4E42F
         bData           =   -1  'True
         Filename        =   "Banned User.png"
         eScale          =   1
         lContrast       =   0
         lBrightness     =   0
         lAlpha          =   100
         bGrayScale      =   0   'False
         lAngle          =   0
         bFlipH          =   0   'False
         bFlipV          =   0   'False
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   330
         Index           =   0
         Left            =   1665
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label ingresar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso al Sistema"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   195
         Left            =   4155
         TabIndex        =   5
         Top             =   1185
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1890
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   2310
         Width           =   1065
      End
   End
End
Attribute VB_Name = "PassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEntrar_Click()
    If cmbUsuario.ListIndex < 0 Then MsgBox "Seleccione el Usuario del Sistema...?", vbOKOnly + _
        vbInformation, "Información al usuario...!": cmbUsuario.SetFocus: Exit Sub
            
            '\\\Si Ingresa el Administrador.
            If cmbUsuario.Text = "ADMINISTRADOR" Then
                If UCase(Trim(TxtPassWord.Text)) = "INKAFARMA" Then
                    PassWord.Hide
                    wMain.Show
                    user = cmbUsuario.Text
                Else
                    MsgBox "La contraseña no es Correcta...?", vbOKOnly + vbCritical, _
                    "Información al usuario...!": TxtPassWord.Text = "": TxtPassWord.SetFocus
                End If
                
            Else
                    '\\\Si ingresa el Vendedor.
                    If UCase(Trim(TxtPassWord.Text)) = "VENTAS" Then
                        user = cmbUsuario.Text
                        PassWord.Hide
                        wMain.Show
                  
                    Else
                        MsgBox "La contraseña no es Correcta...?", vbOKOnly + vbCritical, _
                        "Información al usuario...!": TxtPassWord.Text = "": TxtPassWord.SetFocus
                    End If
            End If
    Call wMain.Form_Load
End Sub

Private Sub Form_Load()
    cmbUsuario.Text = "ADMINISTRADOR"
End Sub

Private Sub Form_Resize()
    '\\\Centramos el Picture.
    ContObj.Move (Me.Width - ContObj.Width) / 2, (Me.Height - ContObj.Height) / 2
End Sub

