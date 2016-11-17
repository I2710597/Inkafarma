VERSION 5.00
Begin VB.Form frmEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleados"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
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
   ScaleHeight     =   4350
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton cmbSal 
      Height          =   870
      Left            =   5640
      TabIndex        =   17
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmEmp.frx":0000
   End
   Begin SistemaInkaFarma.mButton cmbBus 
      Height          =   870
      Left            =   4560
      TabIndex        =   16
      ToolTipText     =   "Buscar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmEmp.frx":D292
   End
   Begin SistemaInkaFarma.mButton cmbEli 
      Height          =   870
      Left            =   3480
      TabIndex        =   15
      ToolTipText     =   "Eliminar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmEmp.frx":1A524
   End
   Begin SistemaInkaFarma.mButton cmbModi 
      Height          =   870
      Left            =   2400
      TabIndex        =   14
      ToolTipText     =   "Editar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmEmp.frx":277B6
   End
   Begin SistemaInkaFarma.mButton cmbGra 
      Height          =   870
      Left            =   1320
      TabIndex        =   13
      ToolTipText     =   "Guardar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmEmp.frx":34A48
      Enabled         =   0   'False
   End
   Begin SistemaInkaFarma.mButton cmbNue 
      Height          =   870
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Nuevo"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmEmp.frx":41CDA
   End
   Begin VB.TextBox txtFing 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtDire 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtTel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   7
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtNom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtApel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtCod 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de ingreso"
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
      Left            =   870
      TabIndex        =   10
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Label Label5 
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
      Left            =   1515
      TabIndex        =   8
      Top             =   3360
      Width           =   780
   End
   Begin VB.Label Label4 
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
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nombres"
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
      Left            =   1545
      TabIndex        =   4
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos"
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
      Left            =   1530
      TabIndex        =   2
      Top             =   1920
      Width           =   765
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
      Left            =   1725
      TabIndex        =   0
      Top             =   1440
      Width           =   570
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBus_Click()
    With emp
        If .RecordCount > 0 Then
            empleado = InputBox("Ingrese El Apellido del Empleado", "Búsqueda")
            If empleado <> "" Then
                .MoveFirst
                Do While Not .EOF
                    If UCase(!apell) = UCase(empleado) Then
                        mdProce.mostEmp
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        Else
            MsgBox "No hay ningún Empleado almacenado", vbOKOnly + vbInformation, "Información"
        End If
    End With
End Sub

Private Sub cmbEli_Click()
    With emp
        If .RecordCount > 0 Then
            If MsgBox("¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                If .RecordCount = 1 Then
                    If MsgBox("Este es el último registro" + Chr(10) + Chr(13) + "¿ Seguro de borrarlo ?", vbYesNo + vbInformation, "Confirme") = vbYes Then
                        cnn.Execute ("delete from TBEmpleado where codemp='" + txtCod.Text + "'")
                    End If
                Else
                    cnn.Execute ("delete from TBEmpleado where codemp='" + txtCod.Text + "'")
                    .Requery
                    mdProce.mostEmp
                End If
            End If
        Else
            MsgBox "NO hay registros", vbOKOnly + vbInformation, "No se puede borrar"
        End If
    End With
End Sub

Private Sub cmbGra_Click()
    If txtApel.Text = "" Then
        MsgBox "Ingrese Apellido", vbOKOnly + vbInformation, "Cuidado"
        txtApel.SetFocus
    ElseIf txtNom.Text = "" Then
        MsgBox "Ingrese Nombre", vbOKOnly + vbInformation, "Cuidado"
        txtNom.SetFocus
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
            cnn.Execute ("insert into TBEmpleado values('" + txtCod.Text + "','" + txtApel.Text + "','" + txtNom.Text + "','" + txtTel.Text + "','" + txtDire.Text + "','" + txtFing.Text + "')")
            emp.Requery
            If MsgBox("¿ Desea ingresar otro Empleado ?", vbYesNo + vbInformation, "Pregunta") = vbYes Then
                cmbNue_Click
            Else
                mdProce.bloqEmp
            End If
        Else
            cnn.Execute ("update TBEmpleado set codemp='" + txtCod.Text + "',apell='" + txtApel.Text + "',nomb='" + txtNom.Text + "',telef='" + txtTel.Text + "',direc='" + txtDire.Text + "',fecing='" + txtFing.Text + "' where codemp='" + txtCod.Text + "'")
            emp.Requery
            mdProce.bloqEmp
        End If
    End If
End Sub

Private Sub cmbModi_Click()
    nuevo = False
    mdProce.desbloqEmp
End Sub

Private Sub cmbNue_Click()
    nuevo = True
    With emp
        mdProce.limEmp
        mdProce.desbloqEmp
         If .RecordCount > 0 Then
            .MoveLast
            txtCod.Text = Trim("EMPL" + String(3 - Len(Trim(Str(Val(Right(!codemp, 3)) + 1))), "0") + Trim(Str(Val(Right(!codemp, 3)) + 1)))
        Else
            txtCod.Text = "EMPL001"
        End If
        txtApel.SetFocus
    End With
End Sub

Private Sub cmbSal_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set emp = New ADODB.Recordset
    With emp
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBEmpleado"
        .Open
        If .RecordCount > 0 Then
            mdProce.mostEmp
        End If
    End With
End Sub

Private Sub txtApel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNom.SetFocus
End Sub

Private Sub txtDire_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFing.SetFocus
End Sub

Private Sub txtFing_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbGra.SetFocus
End Sub

Private Sub txtnom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTel.SetFocus
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDire.SetFocus
End Sub
