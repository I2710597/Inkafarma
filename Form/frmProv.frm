VERSION 5.00
Begin VB.Form frmProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton cmbSal 
      Height          =   870
      Left            =   6960
      TabIndex        =   19
      ToolTipText     =   "Exit"
      Top             =   2280
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmProv.frx":0000
   End
   Begin SistemaInkaFarma.mButton cmbBus 
      Height          =   870
      Left            =   5880
      TabIndex        =   18
      ToolTipText     =   "Buscar"
      Top             =   2280
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmProv.frx":D292
   End
   Begin SistemaInkaFarma.mButton cmbModi 
      Height          =   870
      Left            =   6960
      TabIndex        =   17
      ToolTipText     =   "Editar"
      Top             =   1320
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmProv.frx":1A524
   End
   Begin SistemaInkaFarma.mButton cmbEli 
      Height          =   870
      Left            =   5880
      TabIndex        =   16
      ToolTipText     =   "Eliminar"
      Top             =   1320
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmProv.frx":277B6
   End
   Begin SistemaInkaFarma.mButton cmbGra 
      Height          =   870
      Left            =   6960
      TabIndex        =   15
      ToolTipText     =   "Guardar"
      Top             =   360
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmProv.frx":34A48
      Enabled         =   0   'False
   End
   Begin SistemaInkaFarma.mButton cmbNue 
      Height          =   870
      Left            =   5880
      TabIndex        =   14
      ToolTipText     =   "Nuevo"
      Top             =   360
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmProv.frx":41CDA
   End
   Begin VB.TextBox txtWeb 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   40
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   40
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3240
   End
   Begin VB.TextBox txtTel 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   7
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtDire 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtRuc 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   11
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtProv 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox txtCod 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1208
      MaxLength       =   7
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Web"
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
      Left            =   675
      TabIndex        =   12
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Email"
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
      Left            =   600
      TabIndex        =   10
      Top             =   2640
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
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
      Left            =   315
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dirección"
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
      Left            =   270
      TabIndex        =   6
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "R.U.C."
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
      Left            =   570
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
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
      Left            =   165
      TabIndex        =   2
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "frmProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBus_Click()
    With prov
        If .RecordCount > 0 Then
            proveedor = InputBox("Ingrese El Nombre del Proveedor", "Búsqueda")
            If proveedor <> "" Then
                .MoveFirst
                Do While Not .EOF
                    If UCase(!prove) = UCase(proveedor) Then
                        mdProce.mostProv
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        Else
            MsgBox "No hay ningún Proveedor almacenado", vbOKOnly + vbInformation, "Información"
        End If
    End With
End Sub

Private Sub cmbEli_Click()
    With prov
        If .RecordCount > 0 Then
            If MsgBox("¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                If .RecordCount = 1 Then
                    If MsgBox("Este es el último registro" + Chr(10) + Chr(13) + "¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                        cnn.Execute ("delete from TBProveedor where codprov='" + txtCod.Text + "'")
                    End If
                Else
                    cnn.Execute ("delete from TBProveedor where codprov='" + txtCod.Text + "'")
                    .Requery
                    mdProce.mostProv
                End If
            End If
        Else
            MsgBox "NO hay registros", vbOKOnly + vbInformation, "No se puede borrar"
        End If
    End With
End Sub

Private Sub cmbGra_Click()
    If txtProv.Text = "" Then
        MsgBox "Ingrese proveedor", vbOKOnly + vbInformation, "Cuidado"
        txtProv.SetFocus
    ElseIf txtRuc.Text = "" Then
        MsgBox "Ingrese RUC", vbOKOnly + vbInformation, "Cuidado"
        txtRuc.SetFocus
    ElseIf txtDire.Text = "" Then
        MsgBox "Ingrese dirección", vbOKOnly + vbInformation, "Cuidado"
        txtDire.SetFocus
    ElseIf txtTel.Text = "" Then
        MsgBox "Ingrese teléfono", vbOKOnly + vbInformation, "Cuidado"
        txtTel.SetFocus
    ElseIf txtEmail.Text = "" Then
        MsgBox "Ingrese Email", vbOKOnly + vbInformation, "Cuidado"
        txtEmail.SetFocus
    ElseIf txtWeb.Text = "" Then
        MsgBox "Ingrese Página Web", vbOKOnly + vbInformation, "Cuidado"
        txtWeb.SetFocus
    Else
        If nuevo = True Then
            cnn.Execute ("insert into TBProveedor values('" + txtCod.Text + "','" + UCase(txtProv.Text) + "','" + txtRuc.Text + "','" + txtDire.Text + "','" + txtTel.Text + "','" + txtEmail.Text + "','" + txtWeb.Text + "')")
            prov.Requery
            If MsgBox("¿ Desea ingresar otro Proveedor ?", vbYesNo + vbInformation, "Pregunta") = vbYes Then
                cmbNue_Click
            Else
                mdProce.bloqProv
            End If
        Else
            cnn.Execute ("update TBProveedor set codprov='" + txtCod.Text + "',prove='" + UCase(txtProv.Text) + "',ruc='" + txtRuc.Text + "',DIREc='" + txtDire.Text + "',TELF='" + txtTel.Text + "',email='" + txtEmail.Text + "',pweb='" + txtWeb.Text + "' where codprov='" + txtCod.Text + "'")
            prov.Requery
            mdProce.bloqProv
        End If
    End If
End Sub

Private Sub cmbModi_Click()
    nuevo = False
    mdProce.desbloqProv
End Sub

Private Sub cmbNue_Click()
    nuevo = True
    With prov
        mdProce.limProv
        mdProce.desbloqProv
         If .RecordCount > 0 Then
            .MoveLast
            txtCod.Text = Trim("PROV" + String(3 - Len(Trim(Str(Val(Right(!codprov, 3)) + 1))), "0") + Trim(Str(Val(Right(!codprov, 3)) + 1)))
        Else
            txtCod.Text = "PROV001"
        End If
        txtProv.SetFocus
    End With
End Sub

Private Sub cmbSal_Click()
    If pasar = True Then
        pasar = False
        Unload Me
        frmArti.Show
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set prov = New ADODB.Recordset
    With prov
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBProveedor"
        .Open
        If .RecordCount > 0 Then
            mdProce.mostProv
        End If
    End With
End Sub

Private Sub txtDire_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTel.SetFocus
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtWeb.SetFocus
End Sub

Private Sub txtProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRuc.SetFocus
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDire.SetFocus
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtEmail.SetFocus
End Sub

Private Sub txtWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbGra.SetFocus
End Sub
