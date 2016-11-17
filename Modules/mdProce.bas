Attribute VB_Name = "mdProce"
Public Sub mostProv()
    With prov
        If .RecordCount > 0 Then
            frmProv.txtCod.Text = !codprov
            frmProv.txtProv.Text = !prove
            frmProv.txtRuc.Text = !ruc
            frmProv.txtDire.Text = !direc
            frmProv.txtTel.Text = !telf
            frmProv.txtEmail.Text = !email
            frmProv.txtWeb.Text = !pweb
        End If
    End With
End Sub

Public Sub bloqProv()
    frmProv.txtProv.Enabled = False
    frmProv.txtRuc.Enabled = False
    frmProv.txtDire.Enabled = False
    frmProv.txtTel.Enabled = False
    frmProv.txtEmail.Enabled = False
    frmProv.txtWeb.Enabled = False
    frmProv.cmbNue.Enabled = True
    frmProv.cmbGra.Enabled = False
    frmProv.cmbEli.Enabled = True
    frmProv.cmbModi.Enabled = True
    frmProv.cmbBus.Enabled = True
End Sub

Public Sub desbloqProv()
    frmProv.txtProv.Enabled = True
    frmProv.txtRuc.Enabled = True
    frmProv.txtDire.Enabled = True
    frmProv.txtTel.Enabled = True
    frmProv.txtEmail.Enabled = True
    frmProv.txtWeb.Enabled = True
    frmProv.cmbNue.Enabled = False
    frmProv.cmbGra.Enabled = True
    frmProv.cmbEli.Enabled = False
    frmProv.cmbModi.Enabled = False
    frmProv.cmbBus.Enabled = False
End Sub

Public Sub limProv()
    frmProv.txtCod.Text = ""
    frmProv.txtProv.Text = ""
    frmProv.txtRuc.Text = ""
    frmProv.txtDire.Text = ""
    frmProv.txtTel.Text = ""
    frmProv.txtEmail.Text = ""
    frmProv.txtWeb.Text = ""
    
End Sub

Public Sub mostArti()
    With medi
        If .RecordCount > 0 Then
            frmArti.txtCod.Text = !codart
            frmArti.txtMedi.Text = !arti
            frmArti.txtPrecs.Text = !precsol
            frmArti.txtPrecd.Text = !precdol
            frmArti.txtPres.Text = !present
            frmArti.txtFing.Text = !fing
            frmArti.txtFexp.Text = !fexp
            frmArti.txtCodp.Text = !codprov
            frmArti.cbProv.Text = !prov
        End If
    End With
End Sub

Public Sub limArti()
    frmArti.txtCod.Text = ""
    frmArti.txtMedi.Text = ""
    frmArti.txtPrecs.Text = ""
    frmArti.txtPrecd.Text = ""
    frmArti.txtPres.Text = ""
    frmArti.txtFing.Text = ""
    frmArti.txtFexp.Text = ""
    frmArti.txtCodp.Text = ""
    frmArti.cbProv.Text = ""
End Sub

Public Sub bloqArti()
    frmArti.cmbNue.Enabled = True
    frmArti.cmbGra.Enabled = False
    frmArti.cmbEli.Enabled = True
    frmArti.cmbModi.Enabled = True
    frmArti.cmbBus.Enabled = True
    frmArti.txtMedi.Enabled = False
    frmArti.txtPrecs.Enabled = False
    frmArti.txtPrecd.Enabled = False
    frmArti.txtPres.Enabled = False
    frmArti.txtFing.Enabled = False
    frmArti.txtFexp.Enabled = False
    frmArti.txtCodp.Enabled = False
    frmArti.cbProv.Enabled = False
    frmArti.cmbProv.Enabled = False
End Sub


Public Sub desbloqArti()
    frmArti.cmbNue.Enabled = False
    frmArti.cmbGra.Enabled = True
    frmArti.cmbEli.Enabled = False
    frmArti.cmbModi.Enabled = False
    frmArti.cmbBus.Enabled = False
    frmArti.txtMedi.Enabled = True
    frmArti.txtPrecs.Enabled = True
    frmArti.txtPrecd.Enabled = True
    frmArti.txtPres.Enabled = True
    frmArti.txtFing.Enabled = True
    frmArti.txtFexp.Enabled = True
    frmArti.txtCodp.Enabled = True
    frmArti.cbProv.Enabled = True
    frmArti.cmbProv.Enabled = True
End Sub

Public Sub desbloqEmp()
    frmEmp.cmbNue.Enabled = False
    frmEmp.cmbGra.Enabled = True
    frmEmp.cmbEli.Enabled = False
    frmEmp.cmbModi.Enabled = False
    frmEmp.cmbBus.Enabled = False
    frmEmp.txtApel.Enabled = True
    frmEmp.txtNom.Enabled = True
    frmEmp.txtTel.Enabled = True
    frmEmp.txtDire.Enabled = True
    frmEmp.txtFing.Enabled = True
End Sub

Public Sub desbloqCli()
    frmCli.cmbNue.Enabled = False
    frmCli.cmbGra.Enabled = True
    frmCli.cmbEli.Enabled = False
    frmCli.cmbModi.Enabled = False
    frmCli.cmbBus.Enabled = False
    frmCli.txtNom.Enabled = True
    frmCli.txtRuc.Enabled = True
    frmCli.txtTel.Enabled = True
    frmCli.txtDire.Enabled = True
    frmCli.txtFing.Enabled = True
End Sub

Public Sub bloqEmp()
    frmEmp.cmbNue.Enabled = True
    frmEmp.cmbGra.Enabled = False
    frmEmp.cmbEli.Enabled = True
    frmEmp.cmbModi.Enabled = True
    frmEmp.cmbBus.Enabled = True
    frmEmp.txtApel.Enabled = False
    frmEmp.txtNom.Enabled = False
    frmEmp.txtTel.Enabled = False
    frmEmp.txtDire.Enabled = False
    frmEmp.txtFing.Enabled = False
End Sub

Public Sub bloqCli()
    frmCli.cmbNue.Enabled = True
    frmCli.cmbGra.Enabled = False
    frmCli.cmbEli.Enabled = True
    frmCli.cmbModi.Enabled = True
    frmCli.cmbBus.Enabled = True
    frmCli.txtNom.Enabled = False
    frmCli.txtRuc.Enabled = False
    frmCli.txtTel.Enabled = False
    frmCli.txtDire.Enabled = False
    frmCli.txtFing.Enabled = False
End Sub

Public Sub limEmp()
    frmEmp.txtCod.Text = ""
    frmEmp.txtApel.Text = ""
    frmEmp.txtNom.Text = ""
    frmEmp.txtTel.Text = ""
    frmEmp.txtDire.Text = ""
    frmEmp.txtFing.Text = ""
End Sub

Public Sub limCli()
    frmCli.txtCod.Text = ""
    frmCli.txtNom.Text = ""
    frmCli.txtRuc.Text = ""
    frmCli.txtTel.Text = ""
    frmCli.txtDire.Text = ""
    frmCli.txtFing.Text = ""
End Sub

Public Sub mostEmp()
    With emp
        If .RecordCount > 0 Then
            frmEmp.txtCod.Text = !codemp
            frmEmp.txtApel.Text = !apell
            frmEmp.txtNom.Text = !nomb
            frmEmp.txtTel.Text = !telef
            frmEmp.txtDire.Text = !direc
            frmEmp.txtFing.Text = !fecing
        End If
    End With
End Sub

Public Sub mostCli()
    With cli
        If .RecordCount > 0 Then
            frmCli.txtCod.Text = !CODCLI
            frmCli.txtNom.Text = !nomcli
            frmCli.txtRuc.Text = !ruccli
            frmCli.txtTel.Text = !telef
            frmCli.txtDire.Text = !dircli
            frmCli.txtFing.Text = !fecing
        End If
    End With
End Sub

Public Sub bloqBole()
    frmBole.cmbNue.Enabled = True
    frmBole.cmbGra.Enabled = False
    frmBole.cmbBus.Enabled = True
    frmBole.txtFec.Enabled = False
    frmBole.txtCli.Enabled = False
    frmBole.cbEmp.Enabled = False
    frmBole.txtCant.Enabled = False
    frmBole.cbDescrip.Enabled = False
    frmBole.txtDcto.Enabled = False
End Sub

Public Sub bloqFact()
    frmFact.cmbNue.Enabled = True
    frmFact.cmbGra.Enabled = False
    frmFact.cmbBus.Enabled = True
    frmFact.cmbCli.Enabled = False
    frmFact.txtFec.Enabled = False
    frmFact.cbCli.Enabled = False
    frmFact.cbEmp.Enabled = False
    frmFact.txtCant.Enabled = False
    frmFact.cbDescrip.Enabled = False
    frmFact.txtDcto.Enabled = False
End Sub

Public Sub desbloqBole()
    frmBole.cmbNue.Enabled = False
    frmBole.cmbGra.Enabled = True
    frmBole.cmbBus.Enabled = False
    frmBole.txtCli.Enabled = True
    frmBole.cbEmp.Enabled = True
    frmBole.txtCant.Enabled = True
    frmBole.cbDescrip.Enabled = True
    frmBole.txtDcto.Enabled = True
End Sub

Public Sub desbloqFact()
    frmFact.cmbNue.Enabled = False
    frmFact.cmbGra.Enabled = True
    frmFact.cmbBus.Enabled = False
    frmFact.cmbCli.Enabled = True
    frmFact.cbCli.Enabled = True
    frmFact.cbEmp.Enabled = True
    frmFact.txtCant.Enabled = True
    frmFact.cbDescrip.Enabled = True
    frmFact.txtDcto.Enabled = True
End Sub

Public Sub limBole()
    frmBole.txtNum.Text = ""
    frmBole.txtCli.Text = ""
    frmBole.cbEmp.Text = ""
    frmBole.txtCant.Text = ""
    frmBole.cbDescrip.Text = ""
    frmBole.txtPrec.Text = ""
    frmBole.txtDcto.Text = ""
    frmBole.grid.Rows = 1
    frmBole.txtTot.Text = ""
End Sub

Public Sub limFact()
    frmFact.txtNum.Text = ""
    frmFact.cbCli.Text = ""
    frmFact.cbEmp.Text = ""
    frmFact.txtCant.Text = ""
    frmFact.cbDescrip.Text = ""
    frmFact.txtPrec.Text = ""
    frmFact.txtDcto.Text = ""
    frmFact.grid.Rows = 1
    frmFact.txtTot.Text = ""
    frmFact.txtSub.Text = ""
    frmFact.txtigv.Text = ""
End Sub

Public Sub mostBole()
    With bole
        If .RecordCount > 0 Then
            frmBole.txtNum.Text = !numbol
            frmBole.txtFec.Text = !fec
            frmBole.txtCli.Text = !cli
            empleado = !codemp
            With emp
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If empleado = !codemp Then
                            frmBole.cbEmp.Text = !apell + " " + !nomb
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
            Set detbole = New ADODB.Recordset
            With detbole
                .ActiveConnection = cnn
                .CursorType = adOpenKeyset
                .Source = "select * from tbdetboleta where numbol='" + frmBole.txtNum.Text + "'"
                .Open
                If .RecordCount > 0 Then
                    frmBole.grid.Rows = 1
                    For f = 1 To .RecordCount
                        frmBole.grid.Rows = frmBole.grid.Rows + 1
                        frmBole.grid.TextMatrix(frmBole.grid.Rows - 1, 0) = Trim(CStr(frmBole.grid.Rows - 1))
                        frmBole.grid.TextMatrix(frmBole.grid.Rows - 1, 1) = !cant
                        frmBole.grid.TextMatrix(frmBole.grid.Rows - 1, 2) = !descrip
                        frmBole.grid.TextMatrix(frmBole.grid.Rows - 1, 3) = !punit
                        frmBole.grid.TextMatrix(frmBole.grid.Rows - 1, 4) = !dcto
                        frmBole.grid.TextMatrix(frmBole.grid.Rows - 1, 5) = !imp
                        .MoveNext
                    Next
                End If
            End With
            frmBole.txtTot.Text = !tot
        End If
    End With
End Sub

Public Sub mostfact()
    With fact
        If .RecordCount > 0 Then
            frmFact.txtNum.Text = !numfact
            frmFact.txtFec.Text = !fec
            cliente = !CODCLI
            empleado = !codemp
            With cli
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If cliente = !CODCLI Then
                            frmFact.cbCli.Text = !nomcli
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
                        If empleado = !codemp Then
                            frmFact.cbEmp.Text = !apell + " " + !nomb
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
            Set detfact = New ADODB.Recordset
            With detfact
                .ActiveConnection = cnn
                .CursorType = adOpenKeyset
                .Source = "select * from tbdetfactura where numfac='" + frmFact.txtNum.Text + "'"
                .Open
                If .RecordCount > 0 Then
                    frmFact.grid.Rows = 1
                    For f = 1 To .RecordCount
                        frmFact.grid.Rows = frmFact.grid.Rows + 1
                        frmFact.grid.TextMatrix(frmFact.grid.Rows - 1, 0) = Trim(CStr(frmFact.grid.Rows - 1))
                        frmFact.grid.TextMatrix(frmFact.grid.Rows - 1, 1) = !cant
                        frmFact.grid.TextMatrix(frmFact.grid.Rows - 1, 2) = !descrip
                        frmFact.grid.TextMatrix(frmFact.grid.Rows - 1, 3) = !punit
                        frmFact.grid.TextMatrix(frmFact.grid.Rows - 1, 4) = !dcto
                        frmFact.grid.TextMatrix(frmFact.grid.Rows - 1, 5) = !imp
                        .MoveNext
                    Next
                End If
            End With
            frmFact.txtSub.Text = !Sub
            frmFact.txtigv.Text = !igv
            frmFact.txtTot.Text = !tot
        End If
    End With
End Sub

Public Sub mostGuia()
    With guia
        If .RecordCount > 0 Then
            frmGuia.txtNum.Text = !numguia
            frmGuia.txtFec.Text = !fec
            proveedor = !codprov
            With prov
                If .RecordCount > o Then
                    .MoveFirst
                    Do While Not .EOF
                        If !codprov = proveedor Then
                            frmGuia.cbProv.Text = !prove
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
            empleado = !codemp
            With emp
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If empleado = !codemp Then
                            frmGuia.cbEmp.Text = !apell + " " + !nomb
                            Exit Do
                        Else
                            .MoveNext
                        End If
                    Loop
                End If
            End With
            Set detguia = New ADODB.Recordset
            With detguia
                .ActiveConnection = cnn
                .CursorType = adOpenKeyset
                .Source = "select * from tbdetguia where numguia='" + frmGuia.txtNum.Text + "'"
                .Open
                If .RecordCount > 0 Then
                    frmGuia.grid.Rows = 1
                    For f = 1 To .RecordCount
                        frmGuia.grid.Rows = frmGuia.grid.Rows + 1
                        frmGuia.grid.TextMatrix(frmGuia.grid.Rows - 1, 0) = Trim(CStr(frmGuia.grid.Rows - 1))
                        frmGuia.grid.TextMatrix(frmGuia.grid.Rows - 1, 1) = !cant
                        frmGuia.grid.TextMatrix(frmGuia.grid.Rows - 1, 2) = !arti
                        .MoveNext
                    Next
                End If
            End With
        End If
    End With
End Sub

Public Sub limpguia()
    frmGuia.txtNum.Text = ""
    frmGuia.cbProv.Text = ""
    frmGuia.txtDire.Text = ""
    frmGuia.txtRuc.Text = ""
    frmGuia.txtFec.Text = ""
    frmGuia.cbEmp.Text = ""
    frmGuia.txtCant.Text = ""
    frmGuia.cbArti.Text = ""
    frmGuia.grid.Rows = 1
End Sub

Public Sub desblogGuia()
    frmGuia.cmbNue.Enabled = False
    frmGuia.cmbGra.Enabled = True
    frmGuia.cmbBus.Enabled = False
    frmGuia.cbProv.Enabled = True
    frmGuia.cbEmp.Enabled = True
    frmGuia.txtCant.Enabled = True
    frmGuia.cbArti.Enabled = True
    frmGuia.cmbBaj.Enabled = True
    frmGuia.cmbArti.Enabled = True
End Sub

Public Sub bloqGuia()
    frmGuia.cmbNue.Enabled = True
    frmGuia.cmbGra.Enabled = False
    frmGuia.cmbBus.Enabled = True
    frmGuia.cbProv.Enabled = False
    frmGuia.cbEmp.Enabled = False
    frmGuia.txtCant.Enabled = False
    frmGuia.cbArti.Enabled = False
    frmGuia.cmbBaj.Enabled = False
    frmGuia.cmbArti.Enabled = False
End Sub
