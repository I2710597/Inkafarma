VERSION 5.00
Begin VB.Form frmArti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicamentos"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
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
   ScaleHeight     =   5370
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton cmbSal 
      Height          =   870
      Left            =   6600
      TabIndex        =   24
      ToolTipText     =   "Exit"
      Top             =   2160
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmArti.frx":0000
   End
   Begin SistemaInkaFarma.mButton cmbBus 
      Height          =   870
      Left            =   5520
      TabIndex        =   23
      ToolTipText     =   "Buscar"
      Top             =   2160
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmArti.frx":D292
   End
   Begin SistemaInkaFarma.mButton cmbModi 
      Height          =   870
      Left            =   6600
      TabIndex        =   22
      ToolTipText     =   "Editar"
      Top             =   1200
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmArti.frx":1A524
   End
   Begin SistemaInkaFarma.mButton cmbEli 
      Height          =   870
      Left            =   5520
      TabIndex        =   21
      ToolTipText     =   "Elimininar"
      Top             =   1200
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmArti.frx":277B6
   End
   Begin SistemaInkaFarma.mButton cmbGra 
      Height          =   870
      Left            =   6600
      TabIndex        =   20
      ToolTipText     =   "Guardar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmArti.frx":34A48
      Enabled         =   0   'False
   End
   Begin SistemaInkaFarma.mButton cmbNue 
      Height          =   870
      Left            =   5520
      TabIndex        =   19
      ToolTipText     =   "Nuevo"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmArti.frx":41CDA
   End
   Begin VB.CommandButton cmbProv 
      Caption         =   "&Proveedores"
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
      Height          =   375
      Left            =   5085
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtCodp 
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
      Left            =   2198
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox cbProv 
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
      Height          =   315
      Left            =   2198
      Sorted          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox txtFexp 
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
      Left            =   2325
      MaxLength       =   10
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1215
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
      Left            =   2325
      MaxLength       =   10
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtPres 
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
      Left            =   2325
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtPrecd 
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
      Left            =   2325
      MaxLength       =   7
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtPrecs 
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
      Left            =   2325
      MaxLength       =   7
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtMedi 
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
      Left            =   2340
      MaxLength       =   100
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2775
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
      Left            =   2325
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   1440
      TabIndex        =   16
      Top             =   4440
      Width           =   660
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   3960
      Width           =   1020
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Expiración"
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
      Left            =   255
      TabIndex        =   12
      Top             =   3120
      Width           =   1980
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Ingreso"
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
      Left            =   465
      TabIndex        =   10
      Top             =   2640
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Presentación"
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
      Left            =   810
      TabIndex        =   8
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Precio dólares"
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
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Precio soles"
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
      Left            =   915
      TabIndex        =   4
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Medicamento"
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
      Left            =   765
      TabIndex        =   2
      Top             =   720
      Width           =   1290
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
      Left            =   1350
      TabIndex        =   0
      Top             =   240
      Width           =   660
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      Height          =   1335
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   6015
   End
End
Attribute VB_Name = "frmArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clave As Integer

Private Sub cbProv_Click()
    With prov
        If cbProv.ListIndex > -1 Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If UCase(!prove) = UCase(cbProv.Text) Then
                        txtCodp.Text = !codprov
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        End If
    End With
End Sub

Private Sub cbProv_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbBus_Click()
    With medi
        If .RecordCount > 0 Then
            medicamento = InputBox("Ingrese Medicamento", "Búsqueda")
            If medicamento <> "" Then
                .MoveFirst
                Do While Not .EOF
                    If UCase(!arti) = UCase(medicamento) Then
                        mdProce.mostArti
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        Else
            MsgBox "No hay ningún medicamento almacenado", vbOKOnly + vbInformation, "Información"
        End If
    End With
End Sub

Private Sub cmbEli_Click()
    With medi
        If .RecordCount > 0 Then
            If MsgBox("¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                If .RecordCount = 1 Then
                    If MsgBox("Este es el último registro" + Chr(10) + Chr(13) + "¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                        cnn.Execute ("delete from TBArticulo where codart='" + txtCod.Text + "'")
                    End If
                Else
                    cnn.Execute ("delete from TBArticulo where codart='" + txtCod.Text + "'")
                    .Requery
                    mdProce.mostArti
                End If
            End If
        Else
            MsgBox "NO hay registros", vbOKOnly + vbInformation, "No se puede borrar"
        End If
    End With
End Sub

Private Sub cmbGra_Click()
    If txtMedi.Text = "" Then
        MsgBox "Ingrese Medicamento", vbOKOnly + vbInformation, "Cuidado"
        txtMedi.SetFocus
    ElseIf txtPrecs.Text = "" Then
        MsgBox "Ingrese Precio en soles", vbOKOnly + vbInformation, "Cuidado"
        txtPrecs.SetFocus
    ElseIf txtPrecd.Text = "" Then
        MsgBox "Ingrese precio en dólares", vbOKOnly + vbInformation, "Cuidado"
        txtPrecd.SetFocus
    ElseIf txtPres.Text = "" Then
        MsgBox "Ingrese presentación", vbOKOnly + vbInformation, "Cuidado"
        txtPres.SetFocus
    ElseIf Not IsDate(txtFing.Text) Then
        MsgBox "Fecha no válida", vbOKOnly + vbInformation, "Cuidado"
        txtFing.SetFocus
    ElseIf Not IsDate(txtFexp.Text) Then
        MsgBox "Fecha no válida", vbOKOnly + vbInformation, "Cuidado"
        txtFexp.SetFocus
    ElseIf cbProv.Text = "" Then
        MsgBox "Selecciona un medicamento", vbOKOnly + vbInformation, "Cuidado"
        cbProv.SetFocus
    Else
        If nuevo = True Then
            cnn.Execute ("insert into TBArticulo values('" + txtCod.Text + "','" + txtMedi.Text + "','" + txtPrecs.Text + "','" + txtPrecd.Text + "','" + txtPres.Text + "','" + txtFing.Text + "','" + txtFexp.Text + "','" + txtCodp.Text + "','" + cbProv.Text + "')")
            medi.Requery
            With alm
                If .RecordCount > 0 Then
                    .MoveLast
                    clave = Val(Trim(!codalm)) + 1
                Else
                    clave = 1
                End If
                cnn.Execute ("insert into TBALMACEN values('" + String(10 - Len(Trim(Str(clave))), "0") + Trim(CStr(Val(Trim(clave)))) + "','" + txtCod.Text + "','" + txtMedi.Text + "','" + cantidad + "')")
                .Requery
            End With
            If MsgBox("¿ Desea ingresar otro Medicamento ?", vbYesNo + vbInformation, "Pregunta") = vbYes Then
                cmbNue_Click
            Else
                mdProce.bloqArti
            End If
        Else
            cnn.Execute ("update TBArticulo set codart='" + txtCod.Text + "',arti='" + UCase(txtMedi.Text) + "',precsol='" + txtPrecs.Text + "',precdol='" + txtPrecd.Text + "',present='" + txtPres.Text + "',fing='" + txtFing.Text + "',fexp='" + txtFexp.Text + "' where codart='" + txtCod.Text + "'")
            medi.Requery
            mdProce.bloqArti
        End If
    End If
End Sub

Private Sub cmbModi_Click()
    nuevo = False
    mdProce.desbloqArti
End Sub

Private Sub cmbNue_Click()
    nuevo = True
    With medi
        mdProce.limArti
        mdProce.desbloqArti
         If .RecordCount > 0 Then
            .MoveLast
            txtCod.Text = Trim("MEDI" + String(3 - Len(Trim(Str(Val(Right(!codart, 3)) + 1))), "0") + Trim(Str(Val(Right(!codart, 3)) + 1)))
        Else
            txtCod.Text = "MEDI001"
        End If
        txtMedi.SetFocus
    End With
End Sub

Private Sub cmbProv_Click()
    frmArti.Hide
    pasar = True
    frmProv.Show
End Sub

Private Sub cmbSal_Click()
    If pasar = True Then
        pasar = False
        frmGuia.Show
        Unload Me
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    With prov
        If .RecordCount > 0 Then
            .MoveFirst
            cbProv.Clear
            cbProv.Text = !prove
            Do While Not .EOF
                cbProv.AddItem !prove
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub Form_Load()
    Set prov = New ADODB.Recordset
    With prov
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBProveedor"
        .Open
        If .RecordCount > 0 Then
            .MoveFirst
            cbProv.Text = !prove
            Do While Not .EOF
                cbProv.AddItem !prove
                .MoveNext
            Loop
        End If
    End With
    Set medi = New ADODB.Recordset
    With medi
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBArticulo"
        .Open
        If .RecordCount > 0 Then
            mdProce.mostArti
        End If
    End With
End Sub
Private Sub txtFexp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbProv.SetFocus
End Sub

Private Sub txtFing_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFexp.SetFocus
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPrecs.SetFocus
End Sub

Private Sub txtPrecd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPres.SetFocus
End Sub

Private Sub txtPrecs_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Then
        KeyAscii = 0
    ElseIf KeyAscii = 46 Or KeyAscii = 8 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        txtPrecd.Text = Trim(CStr(Val(txtPrecs.Text) / cambio))
        txtPres.SetFocus
    End If
End Sub

Private Sub txtPres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFing.SetFocus
End Sub
