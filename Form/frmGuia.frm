VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGuia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guía de Remisión"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin SistemaInkaFarma.mButton cmbSal 
      Height          =   870
      Left            =   3240
      TabIndex        =   23
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   870
      _extentx        =   1535
      _extenty        =   1535
      picture         =   "frmGuia.frx":0000
   End
   Begin SistemaInkaFarma.mButton cmbBus 
      Height          =   870
      Left            =   2280
      TabIndex        =   22
      ToolTipText     =   "Buscar"
      Top             =   240
      Width           =   870
      _extentx        =   1535
      _extenty        =   1535
      picture         =   "frmGuia.frx":D292
   End
   Begin SistemaInkaFarma.mButton cmbGra 
      Height          =   870
      Left            =   1320
      TabIndex        =   21
      ToolTipText     =   "Guardar"
      Top             =   240
      Width           =   870
      _extentx        =   1535
      _extenty        =   1535
      picture         =   "frmGuia.frx":1A524
      enabled         =   0   'False
   End
   Begin SistemaInkaFarma.mButton cmbNue 
      Height          =   870
      Left            =   360
      TabIndex        =   20
      ToolTipText     =   "Nuevo"
      Top             =   240
      Width           =   870
      _extentx        =   1535
      _extenty        =   1535
      picture         =   "frmGuia.frx":277B6
   End
   Begin VB.CommandButton cmbArti 
      Caption         =   "&Articulos"
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
      Height          =   375
      Left            =   7118
      TabIndex        =   19
      Top             =   5618
      Width           =   1095
   End
   Begin VB.CommandButton cmbBaj 
      Caption         =   "&Bajar"
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
      Height          =   375
      Left            =   7118
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3938
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Ingreso de Artículos  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   278
      TabIndex        =   12
      Top             =   3458
      Width           =   8055
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   1575
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         BackColor       =   12648447
         AllowUserResizing=   3
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "^Item        |^Cantidad     |<Articulo                                                                             "
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
      Begin VB.TextBox txtCant 
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
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cbArti 
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
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label7 
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
         Left            =   5040
         TabIndex        =   15
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Artículos"
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
         TabIndex        =   13
         Top             =   480
         Width           =   750
      End
   End
   Begin VB.ComboBox cbEmp 
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
      Height          =   315
      Left            =   5318
      Sorted          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2498
      Width           =   2775
   End
   Begin VB.ComboBox cbProv 
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
      Height          =   315
      Left            =   1478
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1658
      Width           =   2535
   End
   Begin VB.TextBox txtFec 
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
      Left            =   6038
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtruc 
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
      Left            =   1478
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2618
      Width           =   1935
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
      Left            =   1478
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2138
      Width           =   3135
   End
   Begin VB.TextBox txtNum 
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
      Left            =   7118
      MaxLength       =   7
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   338
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   885
      TabIndex        =   6
      Top             =   2625
      Width           =   480
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
      Left            =   585
      TabIndex        =   4
      Top             =   2145
      Width           =   780
   End
   Begin VB.Label Label4 
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
      Left            =   480
      TabIndex        =   2
      Top             =   1665
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Empleado"
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
      Left            =   5318
      TabIndex        =   10
      Top             =   2265
      Width           =   825
   End
   Begin VB.Label Label2 
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
      Left            =   5318
      TabIndex        =   8
      Top             =   1665
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6525
      TabIndex        =   0
      Top             =   225
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   278
      Top             =   1418
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   5078
      Top             =   1418
      Width           =   3255
   End
End
Attribute VB_Name = "frmGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M As Integer
Dim aa(1 To 20)  As String
Dim clave As Integer
Dim tipo As String
Dim ss As Integer
Dim medica As String
Dim proveedor As String
Dim empleado As String

Private Sub cbArti_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbEmp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbProv_Change()
    cbProv_Click
End Sub

Private Sub cbProv_Click()
    With prov
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If cbProv.Text = !prove Then
                    txtDire.Text = !direc
                    txtRuc.Text = !ruc
                    Exit Do
                Else
                    .MoveNext
                End If
            Loop
        End If
    End With

End Sub

Private Sub cbProv_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbArti_Click()
    If txtCant.Text = "" Then
        MsgBox "Ingrese cantidad", vbOKOnly + vbInformation, "Cuidado"
        txtCant.SetFocus
    Else
        cantidad = txtCant.Text
        txtCant.Text = ""
        pasar = True
        frmGuia.Hide
        frmArti.Show
    End If
End Sub

Private Sub cmbBaj_Click()
    If cbArti.Text = "" Then
        MsgBox "Seleccione un medicamento", vbOKOnly + vbInformation, "Cuidado"
        cbArti.SetFocus
    ElseIf txtCant.Text = "" Then
        MsgBox "Ingrese cantidad", vbOKOnly + vbInformation, "Cuidado"
        txtCant.SetFocus
    Else
        Grid.Rows = Grid.Rows + 1
        Grid.TextMatrix(Grid.Rows - 1, 0) = Grid.Rows - 1
        Grid.TextMatrix(Grid.Rows - 1, 1) = txtCant.Text
        Grid.TextMatrix(Grid.Rows - 1, 2) = cbArti.Text
        With medi
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If cbArti.Text = !arti Then
                        aa(M) = !codart
                        M = M + 1
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        End With
        If MsgBox("¿ Seguirá ingresando ?", vbInformation + vbYesNo, "Pregunta") = vbYes Then
            cbArti.Text = ""
            txtCant.Text = ""
            cbArti.SetFocus
        Else
            cmbGra.SetFocus
        End If
    End If
End Sub

Private Sub cmbBus_Click()
    With guia
        If .RecordCount > 0 Then
            guiar = InputBox("Ingrese Nº de Guía", "Búsqueda")
            If guiar <> "" Then
                .MoveFirst
                Do While Not .EOF
                    If !numguia = guiar Then
                        mdProce.mostGuia
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        Else
            MsgBox "No hay Guías almacenadas", vbOKOnly + vbInformation, "Información"
        End If
    End With
End Sub

Private Sub cmbGra_Click()
    If cbProv.Text = "" Then
        MsgBox "Seleccione Proveedor", vbOKOnly + vbInformation, "Cuidado"
        cbProv.SetFocus
    ElseIf cbEmp.Text = "" Then
        MsgBox "Seleccione Empleado", vbOKOnly + vbInformation, "Cuidado"
        cbEmp.SetFocus
    ElseIf Grid.Rows = 1 Then
        MsgBox "Debe ingresar Medicamentos al Grid", vbOKOnly + vbInformation, "Cuidado"
        cbArti.SetFocus
    Else
        With prov
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If cbProv.Text = !prove Then
                        proveedor = !codprov
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        End With
        With emp
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If Left(cbEmp.Text, InStr(1, cbEmp.Text, " ") - 1) = Left(!apell, InStr(1, !apell, " ") - 1) Then
                        empleado = !codemp
                        Exit Do
                    Else
                        .MoveNext
                    End If
                Loop
            End If
        End With
        cnn.Execute ("insert into TBGuia values('" + txtNum.Text + "','" + Format(CDate(txtFec.Text), "dd/mm/yyyy") + "','" + proveedor + "','" + empleado + "')")
        guia.Requery
        With detguia
            If .RecordCount > 0 Then
                .MoveLast
                clave = Val(Trim(!detguia)) + 1
            Else
                clave = 1
            End If
            For X = 1 To Grid.Rows - 1
                With medi
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            If Grid.TextMatrix(X, 2) = !arti Then
                                medica = !codart
                                Exit Do
                            Else
                                .MoveNext
                            End If
                        Loop
                    End If
                End With
                cnn.Execute ("insert into TBDetguia values('" + txtNum.Text + "','" + String(6 - Len(Trim(Str(clave))), "0") + Trim(CStr(Val(Trim(clave)))) + "','" + Grid.TextMatrix(X, 1) + "','" + medica + "','" + Grid.TextMatrix(X, 2) + "')")
                .Requery
                clave = clave + 1
            Next
        End With
        For X = 1 To Grid.Rows - 1
            With kar
                If .RecordCount > 0 Then
                    .MoveLast
                    clave = Val(Trim(!codmov)) + 1
                Else
                    clave = 1
                End If
                tipo = "E"
                
                cnn.Execute ("insert into TBKardex values('" + String(10 - Len(Trim(Str(clave))), "0") + Trim(CStr(Val(Trim(clave)))) + "','" + "GUIA REMISION" + "','" + txtNum.Text + "','" + txtFec.Text + "','" + aa(X) + "','" + Grid.TextMatrix(X, 2) + "','" + tipo + "','" + Grid.TextMatrix(X, 1) + "','" + "0" + "')")
                .Requery
            End With
            With alm
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If !codart = aa(X) Then
                            ss = !stock
                            cnn.Execute ("update TBALMACEN set stock=STOCK + " & Val(Grid.TextMatrix(X, 1)) & " where codart='" + aa(X) + "'")
                            .Requery
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
        Next X
        If MsgBox("¿ Desea generar otra Guía ?", vbYesNo + vbInformation, "Pregunta") = vbYes Then
            cmbNue_Click
        Else
            mdProce.bloqGuia
        End If
    End If
End Sub

Private Sub cmbNue_Click()
    With guia
        mdProce.limpguia
        mdProce.desblogGuia
         If .RecordCount > 0 Then
            .MoveLast
            txtNum.Text = String(7 - Len(Trim(Str(Val(!numguia) + 1))), "0") + Trim(Str(Val(!numguia) + 1))
        Else
            txtNum.Text = "0000001"
        End If
        txtFec.Text = Trim(CStr(Date)) + " " + Trim(CStr(Time))
        cbProv.SetFocus
        M = 1
    End With
End Sub

Private Sub cmbSal_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Set medi = New ADODB.Recordset
    With medi
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBArticulo"
        .Open
        If .RecordCount > 0 Then
            .MoveFirst
            cbArti.Clear
            Do While Not .EOF
                cbArti.AddItem !arti
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub Form_Load()
    Set detguia = New ADODB.Recordset
    With detguia
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from tbdetguia where numguia='" + txtNum.Text + "'"
        .Open
    End With
    Set prov = New ADODB.Recordset
    With prov
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBProveedor"
        .Open
        If .RecordCount > 0 Then
            .MoveFirst
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
            .MoveFirst
            Do While Not .EOF
                cbArti.AddItem !arti
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
    Set guia = New ADODB.Recordset
    With guia
        .ActiveConnection = cnn
        .CursorType = adOpenKeyset
        .Source = "select * from TBGuia"
        .Open
        If .RecordCount > 0 Then
            mdProce.mostGuia
        End If
    End With
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
    pasar = False
End Sub

