VERSION 5.00
Begin VB.Form frmCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
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
   ScaleHeight     =   4245
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton cmbSal 
      Height          =   870
      Left            =   5730
      TabIndex        =   17
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmCli.frx":0000
   End
   Begin SistemaInkaFarma.mButton cmbBus 
      Height          =   870
      Left            =   4590
      TabIndex        =   16
      ToolTipText     =   "Buscar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmCli.frx":D292
   End
   Begin SistemaInkaFarma.mButton cmbEli 
      Height          =   870
      Left            =   3495
      TabIndex        =   15
      ToolTipText     =   "Eliminar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmCli.frx":1A524
   End
   Begin SistemaInkaFarma.mButton cmbModi 
      Height          =   870
      Left            =   2415
      TabIndex        =   14
      ToolTipText     =   "Editar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmCli.frx":277B6
   End
   Begin SistemaInkaFarma.mButton cmbGra 
      Height          =   870
      Left            =   1380
      TabIndex        =   13
      ToolTipText     =   "Guardar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmCli.frx":34A48
      Enabled         =   0   'False
   End
   Begin SistemaInkaFarma.mButton cmbNue 
      Height          =   870
      Left            =   270
      TabIndex        =   12
      ToolTipText     =   "Nuevo"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmCli.frx":41CDA
   End
   Begin VB.TextBox txtCod 
      Enabled         =   0   'False
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
      Left            =   2385
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtNom 
      Enabled         =   0   'False
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
      Left            =   2385
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1860
      Width           =   3015
   End
   Begin VB.TextBox txtRuc 
      Enabled         =   0   'False
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
      Left            =   2385
      MaxLength       =   11
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1335
   End
   Begin VB.TextBox txtTel 
      Enabled         =   0   'False
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
      Left            =   2385
      MaxLength       =   7
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1335
   End
   Begin VB.TextBox txtDire 
      Enabled         =   0   'False
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
      Left            =   2385
      MaxLength       =   50
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3300
      Width           =   3615
   End
   Begin VB.TextBox txtFing 
      Enabled         =   0   'False
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
      Left            =   2385
      MaxLength       =   10
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1590
      TabIndex        =   0
      Top             =   1380
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1575
      TabIndex        =   2
      Top             =   1860
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "R.U.C."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   2340
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   2820
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de ingreso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   3780
      Width           =   1680
   End
End
Attribute VB_Name = "frmCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBus_Click()
    With cli
        If .RecordCount > 0 Then
            cliente = InputBox("Ingrese El Cliente", "Búsqueda")
            If cliente <> "" Then
                .MoveFirst
                Do While Not .EOF
                    If UCase(!nomcli) = UCase(cliente) Then
                        mdProce.mostCli
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        Else
            MsgBox "No hay ningún Cliente almacenado", vbOKOnly + vbInformation, "Información"
        End If
    End With
End Sub

Private Sub cmbEli_Click()
    With cli
        If .RecordCount > 0 Then
            If MsgBox("¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                If .RecordCount = 1 Then
                    If MsgBox("Este es el último registro" + Chr(10) + Chr(13) + "¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                        cnn.Execute ("delete from TBcliente where CODCLI='" + txtCod.Text + "'")
                    End If
                Else
                    cnn.Execute ("delete from TBcliente where CODCLI='" + txtCod.Text + "'")
                    .Requery
                    mdProce.mostCli
                End If
            End If
        Else
            MsgBox "NO hay registros", vbOKOnly + vbInformation, "No se puede borrar"
        End If
    End With
End Sub

Private Sub cmbGra_Click()
    If txtNom.Text = "" Then
        MsgBox "Ingrese Cliente", vbOKOnly + vbInformation, "Cuidado"
        txtNom.SetFocus
    ElseIf txtRuc.Text = "" Then
        MsgBox "Ingrese Ruc", vbOKOnly + vbInformation, "Cuidado"
        txtRuc.SetFocus
    ElseIf Not IsNumeric(txtTel.Text) Then
        MsgBox "Número no válido", vbOKOnly + vbInformation, "Cuidado"
        txtTel.SetFocus
    ElseIf txtDire.Text = "" Then
        MsgBox "Ingrese dirección", vbOKOnly + vbInformation, "Cuidado"
        txtTel.SetFocus
    ElseIf txtFing.Text = "" Then
        MsgBox "Ingrese fecha", vbOKOnly + vbInformation, "Cuidado"
        txtFing.SetFocus
    Else
        If nuevo = True Then
            cnn.Execute ("insert into TBcliente values('" + txtCod.Text + "','" + txtNom.Text + "','" + txtRuc.Text + "','" + txtTel.Text + "','" + txtDire.Text + "','" + txtFing.Text + "')")
            cli.Requery
            If MsgBox("¿ Desea ingresar otro cliente ?", vbYesNo + vbInformation, "Pregunta") = vbYes Then
                cmbNue_Click
            Else
                mdProce.bloqCli
            End If
        Else
            cnn.Execute ("update TBcliente set CODCLI='" + txtCod.Text + "',nomcli='" + txtNom.Text + "',ruccli='" + txtRuc.Text + "',telef='" + txtTel.Text + "',dircli='" + txtDire.Text + "',fecing='" + txtFing.Text + "' where CODCLI='" + txtCod.Text + "'")
            cli.Requery
            mdProce.bloqCli
        End If
    End If
End Sub

Private Sub cmbModi_Click()
    nuevo = False
    mdProce.desbloqCli
End Sub

Private Sub cmbNue_Click()
    nuevo = True
    With cli
        mdProce.limCli
        mdProce.desbloqCli
         If .RecordCount > 0 Then
            .MoveLast
            txtCod.Text = Trim("CLIE" + String(3 - Len(Trim(Str(Val(Right(!CODCLI, 3)) + 1))), "0") + Trim(Str(Val(Right(!CODCLI, 3)) + 1)))
        Else
            txtCod.Text = "CLIE001"
        End If
        txtNom.SetFocus
    End With
End Sub

Private Sub cmbSal_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set cli = New ADODB.Recordset
    With cli
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBcliente"
        .Open
        If .RecordCount > 0 Then
            mdProce.mostCli
        End If
    End With
End Sub

Private Sub txtnom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRuc.SetFocus
End Sub

Private Sub txtDire_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFing.SetFocus
End Sub

Private Sub txtFing_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbGra.SetFocus
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTel.SetFocus
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDire.SetFocus
End Sub
