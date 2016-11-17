VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Movimientos"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   ControlBox      =   0   'False
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
   ScaleHeight     =   3120
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Reporte de Movimientos..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   143
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton Option2 
         Caption         =   "Facturas de Ventas..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Boletas de Ventas..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmbImp 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton cmbPan 
         Caption         =   "&Reporte..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mskFechaDos 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaUno 
         Height          =   315
         Left            =   885
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   5
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbImp_Click()
    Unload Me
End Sub
Private Sub cmbPan_Click()
    If Option1 = True Then
        If dt.rsTotalBoletas.State = adStateOpen Then
            dt.rsTotalBoletas.Close
        End If
        dt.TotalBoletas mskFechaUno.Text, mskFechaDos.Text
        dtRMov.Sections("Sección2").Controls("Etiqueta7").Caption = mskFechaUno
        dtRMov.Sections("Sección2").Controls("Etiqueta9").Caption = mskFechaDos
        dtRMov.Show
    Else
        If dt.rsTotalFacturas.State = adStateOpen Then
            dt.rsTotalFacturas.Close
        End If
        dt.TotalFacturas (mskFechaUno.Text), (mskFechaDos.Text)
        dtRMov02.Sections("Sección2").Controls("Etiqueta7").Caption = mskFechaUno
        dtRMov02.Sections("Sección2").Controls("Etiqueta9").Caption = mskFechaDos
        dtRMov02.Show
    End If
End Sub
Private Sub Form_Activate()
    mskFechaUno.Text = Date - 31
    mskFechaDos.Text = Date
End Sub

Private Sub mskFechaDos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsDate(mskFechaDos) Then
            cmbPan.SetFocus
        Else
            MsgBox "La Fecha ingresada no es Valida...?", vbOKOnly + vbQuestion, "Corregir...!"
            mskFechaDos.SetFocus
        End If
    End If
End Sub

Private Sub mskFechaUno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsDate(mskFechaUno) Then
            mskFechaDos.SetFocus
        Else
            MsgBox "La Fecha ingresada no es Valida...?", vbOKOnly + vbQuestion, "Corregir...!"
            mskFechaUno.SetFocus
        End If
    End If
End Sub
