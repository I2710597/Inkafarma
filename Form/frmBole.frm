VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boleta de Venta"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
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
   ScaleHeight     =   6915
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton cmbSal 
      Height          =   870
      Left            =   3240
      TabIndex        =   24
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmBole.frx":0000
   End
   Begin SistemaInkaFarma.mButton cmbBus 
      Height          =   870
      Left            =   2280
      TabIndex        =   23
      ToolTipText     =   "Buscar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmBole.frx":D292
   End
   Begin SistemaInkaFarma.mButton cmbGra 
      Height          =   870
      Left            =   1320
      TabIndex        =   22
      ToolTipText     =   "Guardar"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmBole.frx":1A524
      Enabled         =   0   'False
   End
   Begin SistemaInkaFarma.mButton cmbNue 
      Height          =   870
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Nuevo"
      Top             =   240
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      Picture         =   "frmBole.frx":277B6
   End
   Begin VB.TextBox txtPrec 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   20
      Top             =   6480
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2520
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4445
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      BackColor       =   12648447
      Enabled         =   -1  'True
      AllowUserResizing=   3
      BorderStyle     =   0
      FormatString    =   $"frmBole.frx":34A48
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtDcto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9720
      MaxLength       =   2
      TabIndex        =   17
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox cbDescrip 
      Enabled         =   0   'False
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox txtCant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.ComboBox cbEmp 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtCli 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      MaxLength       =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtFec 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtNum 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Precio"
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
      Left            =   6720
      TabIndex        =   14
      Top             =   3120
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Total a Pagar"
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
      Left            =   8400
      TabIndex        =   19
      Top             =   6480
      Width           =   1125
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Descuento"
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
      Left            =   8760
      TabIndex        =   16
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
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
      Left            =   4800
      TabIndex        =   12
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
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
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Left            =   330
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      TabIndex        =   2
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "BOLETA DE VENTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "R.U.C. Nº 3310667010"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1575
      Left            =   7680
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   10575
   End
End
Attribute VB_Name = "frmBole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vendedor As String
Dim producto As String
Dim dd As Double
Dim acum As Double
Dim Nimp As Double
Dim clave As Integer
Dim aaa As String

Private Sub cbDescrip_Click()
    If grid.Rows = 1 Then
        If MsgBox(" ¿ Pagará en soles ? ", vbYesNo + vbInformation, "Confirme") = vbYes Then
            dolar = False
            With medi
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If cbDescrip.Text = !arti Then
                            txtPrec.Text = !precsol
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
        Else
            dolar = True
            With medi
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If cbDescrip.Text = !arti Then
                            txtPrec.Text = !precdol
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
        End If
    Else
        If dolar = True Then
            With medi
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If cbDescrip.Text = !arti Then
                            txtPrec.Text = !precdol
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
        Else
            With medi
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If cbDescrip.Text = !arti Then
                            txtPrec.Text = !precsol
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
        End If
    End If
    txtCant.SetFocus
End Sub

Private Sub cbDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbEmp_Click()
    With emp
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If Left(!apell, InStr(1, cbEmp.Text, " ") - 1) = Left(cbEmp.Text, InStr(1, cbEmp.Text, " ") - 1) Then
                    vendedor = !codemp
                    Exit Do
                Else
                    .MoveNext
                End If
            Loop
        End If
    End With
    cbDescrip.SetFocus
End Sub

Private Sub cbEmp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbBus_Click()
    With bole
        If .RecordCount > 0 Then
            boleta = InputBox("Ingrese Nº de Boleta", "Búsqueda")
            If boleta <> "" Then
                .MoveFirst
                Do While Not .EOF
                    If !numbol = boleta Then
                        mdProce.mostBole
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        Else
            MsgBox "No hay Boletas almacenadas", vbOKOnly + vbInformation, "Información"
        End If
    End With
End Sub

Private Sub cmbEli_Click()
    With detbole
        If .RecordCount > 0 Then
            If MsgBox(" ¿ Seguro de eliminarlo ? ", vbYesNo + vbInformation, "Confirme") = vbYes Then
                If .RecordCount = 1 Then
                    If MsgBox("Es el último registro" + Chr(10) + Chr(13) + "¿ Desea eliminarlo ?", vbYesNo + vbInformation, "Confirmación") = vbYes Then
                        cnn.Execute ("delete from TBDetboleta where numbol='" + txtNum.Text + "'")
                    End If
                Else
                    cnn.Execute ("delete from TBDetboleta where numbol='" + txtNum.Text + "'")
                    .Requery
                    mdProce.mostBole
                End If
                With bole
                    If .RecordCount = 1 Then
                        cnn.Execute ("delete from TBBoleta where numbol='" + txtNum.Text + "'")
                    Else
                        cnn.Execute ("delete from TBBoleta where numbol='" + txtNum.Text + "'")
                        .Requery
                    End If
                End With
            End If
            mdProce.limBole
            cmbNue.SetFocus
        End If
    End With
End Sub

Private Sub cmbGra_Click()
    If txtCli.Text = "" Then
        MsgBox "Ingrese cliente", vbOKOnly + vbInformation, "Cuidado"
        txtCli.SetFocus
    ElseIf cbEmp.Text = "" Then
        MsgBox "Ingrese Empleado", vbOKOnly + vbInformation, "Cuidado"
        cbEmp.SetFocus
    ElseIf grid.Rows = 1 Then
        MsgBox "Debe ingresar productos al Grid", vbOKOnly + vbInformation, "Cuidado"
        txtCant.SetFocus
    Else
        cnn.Execute ("insert into TBBoleta values('" + txtNum.Text + "','" + Format(CDate(txtFec.Text), "dd/mm/yyyy") + "','" + txtCli.Text + "','" + vendedor + "'," & Val(txtTot.Text) & ")")
        bole.Requery
        With detbole1
            If .RecordCount > 0 Then
                .MoveLast
                clave = Val(Trim(!detbol)) + 1
            Else
                clave = 1
            End If
            For X = 1 To grid.Rows - 1
                With medi
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            If grid.TextMatrix(X, 2) = !arti Then
                                producto = !codart
                                Exit Do
                            Else
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                cnn.Execute ("update TBALMACEN set stock=STOCK - " & Val(grid.TextMatrix(X, 1)) & " where codart='" + producto + "'")
                alm.Requery
                With kar
                    If .RecordCount > 0 Then
                        .MoveLast
                        clave = Val(Trim(!codmov)) + 1
                    Else
                        clave = 1
                    End If
                    tipo = "S"
                    cnn.Execute ("insert into TBKardex values('" + String(10 - Len(Trim(Str(clave))), "0") + Trim(CStr(Val(Trim(clave)))) + "','" + "BOLETA" + "','" + txtNum.Text + "','" + txtFec.Text + "','" + producto + "','" + grid.TextMatrix(X, 2) + "','" + tipo + "','" + "0" + "','" + grid.TextMatrix(X, 1) + "')")
                    .Requery
                End With
                cnn.Execute ("insert into TBDetboleta values('" + txtNum.Text + "','" + String(6 - Len(Trim(Str(clave))), "0") + Trim(CStr(Val(Trim(clave)))) + "','" + grid.TextMatrix(X, 1) + "','" + producto + "','" + grid.TextMatrix(X, 2) + "'," & grid.TextMatrix(X, 3) & "," & Val(grid.TextMatrix(X, 4)) & "," & Val(grid.TextMatrix(X, 5)) & ")")
                .Requery
                clave = clave + 1
            Next
        End With
        If MsgBox("¿ Desea generar otra Boleta ?", vbYesNo + vbInformation, "Pregunta") = vbYes Then
            cmbNue_Click
        Else
            mdProce.bloqBole
        End If
        acum = 0
    End If
End Sub

Private Sub cmbNue_Click()
    With bole
        mdProce.limBole
        mdProce.desbloqBole
         If .RecordCount > 0 Then
            .MoveLast
            txtNum.Text = String(7 - Len(Trim(Str(Val(!numbol) + 1))), "0") + Trim(Str(Val(!numbol) + 1))
        Else
            txtNum.Text = "0000001"
        End If
        txtFec.Text = Trim(CStr(Date)) + " " + Trim(CStr(Time))
        txtCli.SetFocus
    End With
End Sub

Private Sub cmbSal_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set alm = New ADODB.Recordset
    With alm
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBALMACEN"
        .Open
    End With
    Set kar = New ADODB.Recordset
    With kar
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBKardex"
        .Open
    End With
    Set detbole = New ADODB.Recordset
    With detbole
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from tbdetboleta where numbol='" + frmBole.txtNum.Text + "'"
        .Open
    End With
    Set medi = New ADODB.Recordset
    With medi
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBArticulo"
        .Open
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cbDescrip.AddItem !arti
                .MoveNext
            Loop
        End If
    End With
    Set emp = New ADODB.Recordset
    With emp
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBEmpleado"
        .Open
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cbEmp.AddItem !apell + " " + !nomb
                .MoveNext
            Loop
        End If
    End With
    Set detbole1 = New ADODB.Recordset
    With detbole1
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBDetboleta"
        .Open
    End With
    Set bole = New ADODB.Recordset
    With bole
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBBoleta"
        .Open
        If .RecordCount > 0 Then
            mdProce.mostBole
        End If
    End With
    dolar = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
acum = 0
bole.Requery
End Sub
Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Then
        KeyAscii = 0
    ElseIf KeyAscii = 46 Or KeyAscii = 8 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        With medi
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If cbDescrip.Text = !arti Then
                        aaa = !codart
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        End With
        With alm
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If aaa = !codart Then
                        If !stock = 0 Or Val(!stock) < 0 Then
                            MsgBox "No hay stock para este producto", vbInformation + vbOKOnly, "Cuidado"
                            txtCant.SetFocus
                            Exit Sub
                        Else
                            If Val(txtCant.Text) > Val(!stock) Then
                                MsgBox "La cantidad que pide" + Chr(10) + Chr(13) + "es mayor a la del Stock", vbInformation + vbOKOnly, "Cuidado"
                                MsgBox "te puedo dar" + " " + !stock + " nada mas", vbInformation + vbOKOnly, "Cuidado"
                                txtCant.SetFocus
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                    Else
                        .MoveNext
                        'txtDcto.SetFocus
                        'Exit Sub
                    End If
                Loop
            End If
        End With
    End If
End Sub

Private Sub txtCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbEmp.SetFocus
End Sub

Private Sub txtDcto_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Then
        KeyAscii = 0
    ElseIf KeyAscii = 46 Or KeyAscii = 8 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        If txtDcto <> "" Then
            dd = Val(txtDcto) / 100
            grid.Rows = grid.Rows + 1
            grid.TextMatrix(grid.Rows - 1, 0) = grid.Rows - 1
            grid.TextMatrix(grid.Rows - 1, 1) = txtCant.Text
            grid.TextMatrix(grid.Rows - 1, 2) = cbDescrip.Text
            grid.TextMatrix(grid.Rows - 1, 3) = txtPrec.Text
            grid.TextMatrix(grid.Rows - 1, 4) = Trim(CStr(Val(txtCant.Text) * Val(txtPrec) * dd))
            grid.TextMatrix(grid.Rows - 1, 5) = Val(txtCant.Text) * Val(txtPrec) - Val(txtCant.Text) * Val(txtPrec) * dd
        Else
            grid.Rows = grid.Rows + 1
            grid.TextMatrix(grid.Rows - 1, 0) = grid.Rows - 1
            grid.TextMatrix(grid.Rows - 1, 1) = txtCant.Text
            grid.TextMatrix(grid.Rows - 1, 2) = cbDescrip.Text
            grid.TextMatrix(grid.Rows - 1, 3) = txtPrec.Text
            grid.TextMatrix(grid.Rows - 1, 4) = ""
            grid.TextMatrix(grid.Rows - 1, 5) = Trim(CStr(Val(txtCant.Text) * Val(txtPrec)))
        End If
        Nimp = Val(grid.TextMatrix(grid.Rows - 1, 5))
        acum = acum + Nimp
        txtTot.Text = acum
        If MsgBox(" ¿ Desea seguir ingresando ? ", vbYesNo + vbInformation, "Confirme") = vbYes Then
            txtCant.Text = ""
            cbDescrip.Text = ""
            txtPrec.Text = ""
            txtDcto.Text = ""
            txtCant.SetFocus
        Else
            cmbGra.SetFocus
        End If
    End If
End Sub
