VERSION 5.00
Begin VB.Form wMain 
   BorderStyle     =   0  'None
   Caption         =   "Sistema Farmacia"
   ClientHeight    =   10830
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   19200
   Icon            =   "wMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   1155
   End
   Begin VB.PictureBox ClockAnaloj 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   17025
      ScaleHeight     =   2145
      ScaleWidth      =   2070
      TabIndex        =   57
      Top             =   120
      Width           =   2070
      Begin SistemaInkaFarma.ucImage ucImage1 
         Height          =   2100
         Left            =   -30
         Top             =   45
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   3704
         bvData          =   "wMain.frx":6852
         bData           =   -1  'True
         Filename        =   "10.png"
         eScale          =   1
         lContrast       =   0
         lBrightness     =   0
         lAlpha          =   100
         bGrayScale      =   0   'False
         lAngle          =   0
         bFlipH          =   0   'False
         bFlipV          =   0   'False
      End
   End
   Begin VB.PictureBox contIEMenu 
      BackColor       =   &H00C7722C&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   3120
      Picture         =   "wMain.frx":13EFB
      ScaleHeight     =   1050
      ScaleWidth      =   3135
      TabIndex        =   54
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
      Begin SistemaInkaFarma.BtnMenu ContextIE_menu 
         Height          =   330
         Index           =   1
         Left            =   45
         TabIndex        =   56
         Top             =   375
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   582
         Picture         =   "wMain.frx":1B14D
      End
      Begin SistemaInkaFarma.BtnMenu ContextIE_menu 
         Height          =   330
         Index           =   0
         Left            =   45
         TabIndex        =   55
         Top             =   45
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   582
         Picture         =   "wMain.frx":2B41F
      End
   End
   Begin VB.PictureBox contRepMenu 
      BackColor       =   &H00C7722C&
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   240
      ScaleHeight     =   2955
      ScaleWidth      =   3315
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   3315
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":3B6F1
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   26
         Top             =   330
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":4C783
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   4
         Left            =   0
         TabIndex        =   27
         Top             =   1320
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":5D815
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   5
         Left            =   0
         TabIndex        =   28
         Top             =   1650
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":6E8A7
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   6
         Left            =   0
         TabIndex        =   29
         Top             =   1980
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":7F939
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   7
         Left            =   0
         TabIndex        =   30
         Top             =   2310
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":909CB
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   45
         Top             =   660
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":A1A5D
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuRep 
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   46
         Top             =   990
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         Picture         =   "wMain.frx":B2AEF
      End
   End
   Begin VB.PictureBox contManMenu 
      BackColor       =   &H00C7722C&
      BorderStyle     =   0  'None
      Height          =   1230
      Left            =   3600
      ScaleHeight     =   1230
      ScaleWidth      =   2055
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
      Begin SistemaInkaFarma.BtnMenu SubMenuMan 
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   42
         Top             =   660
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         Picture         =   "wMain.frx":C3B81
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuMan 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         Picture         =   "wMain.frx":CD9F3
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuMan 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   44
         Top             =   330
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         Picture         =   "wMain.frx":D7865
      End
   End
   Begin VB.PictureBox contProcMenu 
      BackColor       =   &H00C7722C&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4440
      ScaleHeight     =   1215
      ScaleWidth      =   2325
      TabIndex        =   37
      Top             =   4920
      Visible         =   0   'False
      Width           =   2325
      Begin SistemaInkaFarma.BtnMenu SubMenuProc 
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   38
         Top             =   660
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
         Picture         =   "wMain.frx":E16D7
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuProc 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
         Picture         =   "wMain.frx":ED229
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuProc 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   40
         Top             =   330
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
         Picture         =   "wMain.frx":F8D7B
      End
   End
   Begin VB.PictureBox contSFMenu 
      BackColor       =   &H00C7722C&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3720
      ScaleHeight     =   1455
      ScaleWidth      =   2205
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   2205
      Begin SistemaInkaFarma.BtnMenu SubMenuSF 
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   48
         Top             =   975
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
         Picture         =   "wMain.frx":1048CD
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuSF 
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   33
         Top             =   660
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
         Picture         =   "wMain.frx":1106DF
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuSF 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
         Picture         =   "wMain.frx":11C4F1
      End
      Begin SistemaInkaFarma.BtnMenu SubMenuSF 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   32
         Top             =   330
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
         Picture         =   "wMain.frx":128303
      End
   End
   Begin VB.Timer tHour 
      Interval        =   500
      Left            =   6120
      Top             =   720
   End
   Begin SistemaInkaFarma.mButton btnStart 
      Height          =   510
      Left            =   1080
      TabIndex        =   1
      Top             =   9960
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   900
      Picture         =   "wMain.frx":134115
   End
   Begin VB.PictureBox imgTaskBar 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   2160
      Picture         =   "wMain.frx":1389A7
      ScaleHeight     =   450
      ScaleWidth      =   1680
      TabIndex        =   0
      Top             =   10080
      Width           =   1680
      Begin VB.Label lbSysTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   900
         TabIndex        =   12
         Top             =   90
         Width           =   660
      End
   End
   Begin VB.PictureBox ContStartMenu 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   7170
      Left            =   7080
      Picture         =   "wMain.frx":139961
      ScaleHeight     =   7170
      ScaleWidth      =   6795
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   6795
      Begin VB.PictureBox contListMenu 
         BorderStyle     =   0  'None
         Height          =   4200
         Left            =   3960
         ScaleHeight     =   4200
         ScaleWidth      =   2775
         TabIndex        =   15
         Top             =   1905
         Width           =   2775
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   585
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":1D8503
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   6
            Left            =   0
            TabIndex        =   17
            Top             =   3240
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":1E9BC5
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   3
            Left            =   0
            TabIndex        =   18
            Top             =   1665
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":1FB287
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   19
            Top             =   2160
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":20C949
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   20
            Top             =   1080
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":21E00B
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   7
            Left            =   0
            TabIndex        =   21
            Top             =   3735
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":22F6CD
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   5
            Left            =   0
            TabIndex        =   22
            Top             =   2655
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Enabled         =   0   'False
            Picture         =   "wMain.frx":240D8F
         End
         Begin SistemaInkaFarma.BtnMenu btnSeparator 
            Height          =   90
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   1575
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   159
            Picture         =   "wMain.frx":252451
         End
         Begin SistemaInkaFarma.BtnMenu btnSeparator 
            Height          =   90
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   3150
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   159
            Picture         =   "wMain.frx":256823
         End
         Begin SistemaInkaFarma.BtnMenu btnSeparator 
            Height          =   90
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   495
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   159
            Picture         =   "wMain.frx":25ABF5
         End
         Begin SistemaInkaFarma.BtnMenu btnMenu 
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   873
            Picture         =   "wMain.frx":25EFC7
         End
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   1
         Left            =   165
         TabIndex        =   5
         Top             =   4320
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":270689
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   165
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":29305B
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   3
         Left            =   165
         TabIndex        =   7
         Top             =   3690
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":2B5A2D
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   2
         Left            =   165
         TabIndex        =   8
         Top             =   2415
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":2D83FF
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   4
         Left            =   165
         TabIndex        =   9
         Top             =   3060
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":2FADD1
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   6
         Left            =   165
         TabIndex        =   10
         Top             =   780
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":31D7A3
      End
      Begin SistemaInkaFarma.BtnMenu listMenu 
         Height          =   540
         Index           =   5
         Left            =   165
         TabIndex        =   11
         Top             =   1425
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   953
         Picture         =   "wMain.frx":340175
      End
      Begin SistemaInkaFarma.mButton btnMenuOFF 
         Height          =   300
         Left            =   4020
         TabIndex        =   34
         ToolTipText     =   "Salir del Sistema."
         Top             =   6735
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Picture         =   "wMain.frx":362B47
      End
      Begin SistemaInkaFarma.mButton btnMenuLOCK 
         Height          =   300
         Left            =   4800
         TabIndex        =   35
         ToolTipText     =   "Cambiar de Usuario."
         Top             =   6735
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Picture         =   "wMain.frx":365C59
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   225
         X2              =   3750
         Y1              =   2190
         Y2              =   2190
      End
      Begin SistemaInkaFarma.ucImage imgLogoSC 
         Height          =   1920
         Left            =   4440
         Top             =   120
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   3387
         bvData          =   "wMain.frx":368D6B
         bData           =   -1  'True
         Filename        =   "LogoSC.png"
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
   Begin VB.PictureBox imgStartB 
      AutoSize        =   -1  'True
      Height          =   2100
      Left            =   6720
      Picture         =   "wMain.frx":36BE26
      ScaleHeight     =   2040
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   7800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox imgStartA 
      AutoSize        =   -1  'True
      Height          =   2100
      Left            =   5880
      Picture         =   "wMain.frx":3706A8
      ScaleHeight     =   2040
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   735
   End
   Begin SistemaInkaFarma.BtnDesktop BtnDesktop 
      Height          =   1080
      Index           =   4
      Left            =   120
      TabIndex        =   53
      ToolTipText     =   "Calculadora"
      Top             =   5400
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1905
      Picture         =   "wMain.frx":374F2A
   End
   Begin SistemaInkaFarma.BtnDesktop BtnDesktop 
      Height          =   1080
      Index           =   3
      Left            =   120
      TabIndex        =   52
      ToolTipText     =   "Blog de Notas"
      Top             =   4080
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1905
      Picture         =   "wMain.frx":38A0FC
   End
   Begin SistemaInkaFarma.BtnDesktop BtnDesktop 
      Height          =   1080
      Index           =   2
      Left            =   120
      TabIndex        =   51
      ToolTipText     =   "Internet Explorer"
      Top             =   2760
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1905
      Picture         =   "wMain.frx":39F2CE
   End
   Begin SistemaInkaFarma.BtnDesktop BtnDesktop 
      Height          =   1080
      Index           =   1
      Left            =   120
      TabIndex        =   50
      ToolTipText     =   "Mis Documentos"
      Top             =   1440
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1905
      Picture         =   "wMain.frx":3B44A0
   End
   Begin SistemaInkaFarma.BtnDesktop BtnDesktop 
      Height          =   1080
      Index           =   0
      Left            =   120
      TabIndex        =   49
      ToolTipText     =   "Mi PC"
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1905
      Picture         =   "wMain.frx":3C9672
   End
   Begin VB.Image imgFondo 
      Height          =   3405
      Left            =   7440
      Picture         =   "wMain.frx":3DE844
      Top             =   6360
      Width           =   6375
   End
End
Attribute VB_Name = "wMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159265
Dim Frm As Form

Private Sub BtnDesktop_DblClick(Index As Integer)
    ResetAll

    Select Case Index

        Case 0
            'Mis documentos
            OpenEXE "::{450D8FBA-AD25-11D0-98A8-0800361B1103}"
        Case 1
            'Mi PC
            OpenEXE "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        Case 2
            'Internet explorer
            OpenEXE "C:\Archivos de programa\Internet Explorer\IEXPLORE.exe"
        Case 3
            'Bloc de notas
            OpenEXE "notepad.exe"
        Case 4
            'Calculadora
            OpenEXE "calc.exe"
    End Select
End Sub

Private Sub BtnDesktop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Select Case Index
            'Botón Internet Explorer
            Case 2
            contIEMenu.Left = BtnDesktop(2).Left + X: contIEMenu.Top = BtnDesktop(2).Top + Y
            contIEMenu.Visible = True
        End Select
    End If
End Sub

Private Sub btnMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            'Botón Sistema Chocano
            ResetAll
        Case 1
            'Contiene submenu así es que no ejerce función
        Case 2
            'Botón Integrantes
            ResetAll
            About.Show
        Case 3
            'Botón Mi PC
            ResetAll
            OpenEXE "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        Case 4
            'Botón Panel de control(en XP)
            ResetAll
            OpenEXE "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}"
        Case 5
            'Desabilitado, ahi se verá, jaja
        Case 6
            'Botón
            ResetAll
        Case 7
            'Botón
            ResetAll
    End Select
End Sub

Private Sub btnMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case 1
            btnMenu(1).SubCheck = True
            contSFMenu.Visible = True ': contJGMenu.Visible = False
        Case 0, 2, 3, 4, 6, 7
            ResetSF
        Case 5
            'btnMenu(5).SubCheck = True
            contSFMenu.Visible = False ': contJGMenu.Visible = True
    End Select
    btnMenu(Index).Refresh
End Sub

Private Sub btnMenuLOCK_Click()
    ResetAll
    PassWord.Show
End Sub

Private Sub btnMenuOFF_Click()
    ResetAll
    If MsgBox("Realmente desea salir del Sistema?", 36, "Sistema Farmacia") = 6 Then
        End
    End If
End Sub

Private Sub btnSeparator_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetSF
End Sub

Private Sub btnStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        If ContStartMenu.Visible = False Then
            ContStartMenu.Visible = True
            ContStartMenu.SetFocus
            Set btnStart.Picture = imgStartB.Picture
        Else
            ContStartMenu.Visible = False
            Set btnStart.Picture = imgStartA.Picture
        End If
    Else
        ContStartMenu.Visible = False
        Set btnStart.Picture = imgStartA.Picture
    End If
    
End Sub

Private Sub ClockAnaloj_Paint()
    Reloj
End Sub

Private Sub ContextIE_menu_Click(Index As Integer)
    ResetAll
    Select Case Index

        Case 0
            'Página principal
            OpenEXE "C:\Archivos de programa\Internet Explorer\IEXPLORE.exe"

        Case 1
            'Diálogo de Agregar URL
            wAddURL.Show vbModal
    End Select
End Sub

'Private Sub contJGMenu_Click()
    'ResetAll
'End Sub

Private Sub contListMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetSF
End Sub

Private Sub ContStartMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetSF
End Sub

Private Sub Form_Initialize()
    ResizeAll
End Sub

Public Sub Form_Load()

    If user = "VENTAS" Then
        With wMain
            .SubMenuProc(2).Enabled = False
            .SubMenuRep(0).Enabled = False
            .SubMenuRep(1).Enabled = False
            .SubMenuRep(2).Enabled = False
            .SubMenuRep(6).Enabled = False
            .SubMenuRep(7).Enabled = False
            .SubMenuMan(1).Enabled = False
            .SubMenuMan(2).Enabled = False
        End With
    Else
         With wMain
            .SubMenuProc(2).Enabled = True
            .SubMenuRep(0).Enabled = True
            .SubMenuRep(1).Enabled = True
            .SubMenuRep(2).Enabled = True
            .SubMenuRep(6).Enabled = True
            .SubMenuRep(7).Enabled = True
            .SubMenuMan(1).Enabled = True
            .SubMenuMan(2).Enabled = True
         End With
    End If
    
    Timer1.Interval = 1000
    ClockAnaloj.BorderStyle = 0
    
    '///////////////////////
    'Conexion ah la Base de Datos.
    Set cnn = New ADODB.Connection
    'ACCESS 2000
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & _
    "\dbFarmacia.mdb;Persist Security Info=False"
    '///////////////////////
    
    Me.BackColor = vbWhite
    contListMenu.BackColor = &H43403C
    StretchPicture imgTaskBar, Horizontal
    tHour_Timer

    '///////////////////////
    'Sistema Farmacia
    contSFMenu.Width = SubMenuSF(0).Width: contSFMenu.Height = SubMenuSF(0).Height * 4
    contProcMenu.Width = SubMenuProc(0).Width: contProcMenu.Height = SubMenuProc(0).Height * 3
    contManMenu.Width = SubMenuMan(0).Width: contManMenu.Height = SubMenuMan(0).Height * 3
    contRepMenu.Width = SubMenuRep(0).Width: contRepMenu.Height = SubMenuRep(0).Height * 8
    'Juegos
    'contJGMenu.Width = SubMenuJG(0).Width: contJGMenu.Height = SubMenuJG(0).Height * 6

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetAll
End Sub

Private Sub Form_Resize()
    ResizeAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For Each Frm In Forms
        Unload Frm
    Next
End Sub

Private Sub imgFondo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetAll
    Set btnStart.Picture = imgStartA.Picture
End Sub

Private Sub imgLogoSC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetSF
End Sub

Private Sub imgTaskBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetAll
    Set btnStart.Picture = imgStartA.Picture
End Sub

Private Sub lbSysTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetAll
End Sub

Private Sub listMenu_Click(Index As Integer)
    ResetAll
        Select Case Index

            Case 0
                'Internet Explorer
                OpenEXE "C:\Archivos de programa\Internet Explorer\IEXPLORE.exe"
            Case 1
                'Microsoft Access(Desabilitado)
                OpenEXE "C:\Archivos de programa\Microsoft Office\OFFICE12\MSACCESS.EXE"
            Case 2
                'Microsoft Word
                OpenEXE "C:\Archivos de programa\Microsoft Office\OFFICE12\WINWORD.exe"
            Case 3
                'Microsoft PowerPoint
                OpenEXE "C:\Archivos de programa\Microsoft Office\OFFICE12\POWERPNT.EXE"
            Case 4
                'Microsoft Excel
                OpenEXE "C:\Archivos de programa\Microsoft Office\OFFICE12\EXCEL.EXE"
            Case 5
                'Bloc de notas
                OpenEXE "notepad.exe"
            Case 6
                'Calculadora
                OpenEXE "calc.exe"
        End Select

End Sub

Private Sub listMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetSF
End Sub

Private Sub SubMenuMan_Click(Index As Integer)
    'Menú Mantenimiento
    Select Case Index
        Case 0 'Clientes
            frmCli.Show
        Case 1 'Empleados
            frmEmp.Show
        Case 2 'Proveedores
            frmProv.Show
    End Select
    ResetAll
End Sub

Private Sub SubMenuProc_Click(Index As Integer)
    'Menú Procedimiento
    ResetAll
    Select Case Index
        Case 0 'Boleta de venta
            frmBole.Show
        Case 1 'Factura de venta
            frmFact.Show
        Case 2 'Guía de remisión
            frmGuia.Show
    End Select

End Sub

Private Sub SubMenuRep_Click(Index As Integer)
    'Menú Reportes
    ResetAll
    Select Case Index

        Case 0 'Empleados
            If dt.rsEmpleado.State = adStateOpen Then
                dt.rsEmpleado.Close
            End If
            dtREmp.Show
        
        Case 1 'Medicamentos
            If dt.rsArticulo.State = adStateOpen Then
                dt.rsArticulo.Close
            End If
            dtRMedi.Show
        
        Case 2 'Proveedores
            If dt.rsProveedor.State = adStateOpen Then
                dt.rsProveedor.Close
            End If
            dtRProv.Show
        
        Case 3 'Boleta de ventas
            If dt.rsBoleta_Grupo.State = adStateOpen Then
                dt.rsBoleta_Grupo.Close
            End If
            dtRBol.Show
        
        Case 4 'Factura de ventas
            If dt.rsFactura_Grupo.State = adStateOpen Then
                dt.rsFactura_Grupo.Close
            End If
            dtRFact.Show
        
        Case 5 'Movimiento de documento
            frmMov.Show
        
        Case 6 'Almacén
            If dt.rsAlmacen.State = adStateOpen Then
                dt.rsAlmacen.Close
            End If
            dtRAlm.Show
        
        Case 7 'Kardex
            If dt.rsKardex_Grupo.State = adStateOpen Then
                dt.rsKardex_Grupo.Close
            End If
            dtRKar.Show
    End Select

End Sub

Private Sub SubMenuSF_Click(Index As Integer)
    ResetAll
    Select Case Index
        Case 3 'Botón Tipo de cambio
            frmCambio.Show
    End Select
End Sub

Private Sub SubMenuSF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'No modificar
    Select Case Index

        Case 0
            SubMenuSF(0).SubCheck = True: SubMenuSF(1).SubCheck = False: SubMenuSF(2).SubCheck = False
            contProcMenu.Visible = True: contManMenu.Visible = False: contRepMenu.Visible = False

        Case 1
            SubMenuSF(0).SubCheck = False: SubMenuSF(1).SubCheck = True: SubMenuSF(2).SubCheck = False
            contProcMenu.Visible = False: contManMenu.Visible = True: contRepMenu.Visible = False

        Case 2
            SubMenuSF(0).SubCheck = False: SubMenuSF(1).SubCheck = False: SubMenuSF(2).SubCheck = True
            contProcMenu.Visible = False: contManMenu.Visible = False: contRepMenu.Visible = True
        
        Case 3
            SubMenuSF(0).SubCheck = False: SubMenuSF(1).SubCheck = False: SubMenuSF(2).SubCheck = False
            contProcMenu.Visible = False: contManMenu.Visible = False: contRepMenu.Visible = False

    End Select
    btnMenu(Index).Refresh

End Sub

Private Sub tHour_Timer()
    'Hora del sistema.
    lbSysTime.Caption = Format(Hour(Time), "0#") & ":" & Format(Minute(Time), "0#")
End Sub

'//////////////////////////////////////////////////////////////////////////////

Private Sub ResetAll()
Set btnStart.Picture = imgStartA.Picture
'contJGMenu.Visible = False

SubMenuSF(0).SubCheck = False: SubMenuSF(1).SubCheck = False: SubMenuSF(2).SubCheck = False
    contProcMenu.Visible = False: contManMenu.Visible = False: contRepMenu.Visible = False
        contSFMenu.Visible = False
            btnMenu(1).SubCheck = False
                ContStartMenu.Visible = False
ResetDesktop
End Sub

Private Sub ResetSF()

    contProcMenu.Visible = False: SubMenuSF(0).SubCheck = False
    contManMenu.Visible = False: SubMenuSF(1).SubCheck = False
    contRepMenu.Visible = False: SubMenuSF(2).SubCheck = False
    contSFMenu.Visible = False: btnMenu(1).SubCheck = False

End Sub

Private Sub ResetDesktop()
    Dim dIndex As BtnDesktop

    For Each dIndex In BtnDesktop
        dIndex.SubCheck = False
    Next
    contIEMenu.Visible = False
End Sub

Public Sub ResizeAll()

    imgFondo.Left = (Me.Width / 2) - (imgFondo.Width / 2)
    imgFondo.Top = (Me.Height / 2) - (imgFondo.Height / 2)

    imgTaskBar.Left = 0: imgTaskBar.Top = Me.ScaleHeight - imgTaskBar.Height
    imgTaskBar.Width = Me.ScaleWidth

    btnStart.Left = 0: btnStart.Top = Me.ScaleHeight - btnStart.Height
    StretchPicture imgTaskBar, Horizontal

    ContStartMenu.Left = 0: ContStartMenu.Top = Me.ScaleHeight - imgTaskBar.Height - ContStartMenu.Height

    lbSysTime.Left = imgTaskBar.ScaleWidth - lbSysTime.Width
    '///////////////////////////////////////////////////////////
    'Menú Sistema Farmacia
    contSFMenu.Left = ContStartMenu.Width - 105: contSFMenu.Top = ContStartMenu.Top + contListMenu.Top + btnMenu(1).Top
    contProcMenu.Left = contSFMenu.Left + contSFMenu.Width - 75: contProcMenu.Top = contSFMenu.Top + SubMenuSF(0).Top
    contManMenu.Left = contSFMenu.Left + contSFMenu.Width - 75: contManMenu.Top = contSFMenu.Top + SubMenuSF(1).Top
    contRepMenu.Left = contSFMenu.Left + contSFMenu.Width - 75: contRepMenu.Top = contSFMenu.Top + SubMenuSF(2).Top
    
    'Menú contextual Internet Explorer
    contIEMenu.AutoSize = True
    
    'contJGMenu.Left = ContStartMenu.Width - 105: contJGMenu.Top = ContStartMenu.Top + contListMenu.Top + btnMenu(5).Top

End Sub

Private Sub Reloj()
Static last_time As Date
Dim cx As Single
Dim cy As Single
Dim num As Single
Dim radius As Single
Dim theta As Single

    If last_time = Now Then Exit Sub
    
    last_time = Now
    ClockAnaloj.Cls
    ClockAnaloj.ForeColor = vbWhite
    cx = ClockAnaloj.ScaleWidth / 2
    cy = ClockAnaloj.ScaleHeight / 2

    ' Horas
    num = 5 * (DatePart("h", last_time) + DatePart("n", last_time) / _
                                60 + DatePart("s", last_time) / 3600)
    theta = MinutesToRadians(num)
    radius = ClockAnaloj.ScaleWidth * 0.25
    ClockAnaloj.DrawWidth = 3
    ClockAnaloj.Line (cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))

    ' Los Minutos
    num = DatePart("n", last_time)
    theta = MinutesToRadians(num)
    radius = ClockAnaloj.ScaleWidth * 0.3
    ClockAnaloj.DrawWidth = 2
    ClockAnaloj.Line (cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))

    ' Los segundos
    num = DatePart("s", last_time)
    theta = MinutesToRadians(num)
    radius = ClockAnaloj.ScaleWidth * 0.33
    ClockAnaloj.DrawWidth = 1
    ClockAnaloj.Line (cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))
End Sub

Private Function MinutesToRadians(ByVal num As Single) As Single
    MinutesToRadians = (15 - num) * 2 * PI / 60
End Function

Private Sub Timer1_Timer()
    ' actualiza
    Reloj
End Sub
