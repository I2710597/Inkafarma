VERSION 5.00
Begin VB.Form Inicio 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tStart 
      Interval        =   7000
      Left            =   360
      Top             =   1680
   End
   Begin VB.PictureBox contObj 
      BackColor       =   &H00C7722C&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   1680
      ScaleHeight     =   2535
      ScaleWidth      =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   2760
      Begin SistemaInkaFarma.mAniControl aniSlide 
         Height          =   420
         Left            =   720
         TabIndex        =   1
         Top             =   435
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         Picture         =   "Inicio.frx":6852
         Interval        =   100
         LayerHeight     =   420
         LockClick       =   -1  'True
      End
      Begin VB.Label mTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "© Sistema Farmacia"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   60
         TabIndex        =   2
         Top             =   -15
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ContObj.BackColor = vbBlack 'Fondo Negro
    aniSlide.PlayAni
End Sub

Private Sub Form_Resize()
    '\\\Posicionamos el Picture.
    ContObj.Left = Me.Width / 2 - ContObj.Width / 2
    ContObj.Top = Me.ScaleHeight - ContObj.Height
End Sub

Private Sub tStart_Timer()
    ContObj.Visible = False
    aniSlide.StopAni: PassWord.Show: Unload Me: tStart = False
End Sub

