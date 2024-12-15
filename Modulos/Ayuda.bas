Attribute VB_Name = "ModuloAyuda"
Global Campo(30) As String 'Nombre real en la base de datos de los campos que quiero ver en la ayuda
Global AyuTitulo As String 'Titulo de la Ayuda
Global TablasSQL As String 'Tablas a Consultar para mostrar la ayuda
Global LinkSQL As String ' Links entre las tablas si lo hay
Global ColDestino As Integer 'Columna donde se encuentra la clave dos
Global Destino2 As Boolean 'Marca cuando se utiliza un segundo control ayudas con dos claves
Public Sub PROVINCIA_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.prv_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT PRV_DESC FROM PROVINCIA  WHERE PRV_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.prv_desc.Text = snp!prv_desc & ""
                    If Prop Then BTActualizar frm
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.prv_id.Text = ""
                        frm.prv_desc.Text = ""
                    End If
                    If Prop Then BTAgregar frm
                    frm.prv_id.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.prv_id = ""
            frm.prv_desc = ""
        End If
    End If
End Sub


Public Sub LOCALIDAD_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.loc_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT LOC_DESC, L.PRV_ID, PRV_DESC FROM LOCALIDAD AS L, PROVINCIA AS P WHERE P.PRV_ID=L.PRV_ID AND L.LOC_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.loc_desc.Text = snp!loc_desc
                    frm.prv_id.Text = snp!prv_id
                    frm.prv_desc.Text = snp!prv_desc
                    If Prop Then BTActualizar frm
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.loc_id.Text = ""
                        frm.loc_desc.Text = ""
                        frm.prv_id.Text = ""
                        frm.prv_desc.Text = ""
                    End If
                    If Prop Then BTAgregar frm
                    frm.loc_id.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.loc_id = ""
            frm.loc_desc = ""
            frm.prv_id.Text = ""
            frm.prv_desc.Text = ""
        End If
    End If
End Sub
Public Sub RENDICIONC_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.txtRend_Nro.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT LOC_DESC, L.PRV_ID, PRV_DESC FROM LOCALIDAD AS L, PROVINCIA AS P WHERE P.PRV_ID=L.PRV_ID AND L.LOC_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.loc_desc.Text = snp!loc_desc
                    frm.prv_id.Text = snp!prv_id
                    frm.prv_desc.Text = snp!prv_desc
                    If Prop Then BTActualizar frm
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                Case 0
                    MsgBox "Rendición Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.loc_id.Text = ""
                        frm.loc_desc.Text = ""
                        frm.prv_id.Text = ""
                        frm.prv_desc.Text = ""
                    End If
                    If Prop Then BTAgregar frm
                    frm.txtRend_Nro.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            sSQL$ = "SELECT max(rend_nro) as numero FROM REND_COB"
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.txtRend_Nro.Text = Val(ChkNull(snp!Numero)) + 1
                Case 0
                    frm.txtRend_Nro.Text = 0
            End Select
'            frm.loc_id = ""
'            frm.loc_desc = ""
'            frm.prv_id.Text = ""
'            frm.prv_desc.Text = ""
        End If
    End If
End Sub




Public Sub ZONA_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.zon_id.Text
    If Trim(Codigo) <> "" Then
            sSQL$ = "SELECT ZON_DESC, ZON_OBS FROM ZONA WHERE ZON_ID=" & XS(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.zon_desc.Text = snp!zon_desc
                    If Prop Then
                        frm.zon_obs.Text = snp!zon_obs & ""
                        BTActualizar frm
                    End If
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.zon_id.Text = ""
                        frm.zon_desc.Text = ""
                    End If
                    If Prop Then BTAgregar frm
                    frm.zon_id.SetFocus
            End Select
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.zon_id.Text = ""
            frm.zon_desc.Text = ""
        End If
    End If
End Sub





Public Sub TIPO_PAGO_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.tpp_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT TPP_DESC, TPP_OBS, TPP_ABREV, TPP_RECMANUAL, TPP_RECPREIMP, TPP_DEBAUTO FROM TIPO_PAGO WHERE TPP_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.tpp_desc.Text = snp!tpp_desc & ""
                    If Prop Then
                        frm.tpp_obs = snp!tpp_obs & ""
                        frm.tpp_abrev = snp!tpp_abrev & ""
                        frm.tpp_recmanual = snp!tpp_recmanual & ""
                        frm.tpp_recpreimp = snp!tpp_recpreimp & ""
                        frm.tpp_debauto = snp!tpp_debauto & ""
                        BTActualizar frm
                    End If
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.tpp_id.Text = ""
                        frm.tpp_desc.Text = ""
                    End If
                    If Prop Then BTAgregar frm
                    frm.tpp_id.SetFocus
            End Select
        Else
            MP 0
            MsgBox "El código debe ser numérico positivo", 48, AppName
            If Prop Then
                frm.BlancoCampos
            Else
                frm.tpp_id.Text = ""
                frm.tpp_desc.Text = ""
            End If
            frm.tpp_id.SetFocus
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.tpp_id.Text = ""
            frm.tpp_desc.Text = ""
        End If
    End If
End Sub






Public Sub BARRIO_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.bar_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT BAR_ID, BAR_DESC, LOC_DESC, L.PRV_ID, PRV_DESC, l.loc_id, bar_codpos FROM LOCALIDAD AS L, PROVINCIA AS P, BARRIO AS B WHERE P.PRV_ID=L.PRV_ID AND B.LOC_ID=L.LOC_ID AND B.BAR_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.bar_Desc.Text = snp!bar_Desc
                    frm.loc_id.Text = snp!loc_id
                    frm.loc_desc.Text = snp!loc_desc
                    frm.prv_id.Text = snp!prv_id
                    frm.prv_desc.Text = snp!prv_desc
                    If Prop Then
                        BTActualizar frm
                        frm.bar_codpos.Text = snp!bar_codpos
                    End If
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.bar_id.Text = ""
                        frm.bar_Desc.Text = ""
                        frm.loc_id.Text = ""
                        frm.loc_desc.Text = ""
                        frm.prv_id.Text = ""
                        frm.prv_desc.Text = ""
                    End If
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.bar_id = ""
                        frm.bar_Desc = ""
                    End If
                    frm.bar_id.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.bar_id = ""
            frm.bar_Desc = ""
        End If
    End If
End Sub





Public Sub CLIENTE_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.cli_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT cli_nombre, cli_fecnac, cli_nropami, tdo_id, cli_nrodoc, cli_te, cli_celular, cli_calle, cli_nro, cli_piso, cli_dpto, cli_seccional, bar_id, cli_obs FROM CLIENTE WHERE CLI_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.cli_nombre = snp.Fields(0) & ""
                    frm.cli_fecnac = snp.Fields(1) & ""
                    frm.cli_nropami = snp.Fields(2) & ""
                    frm.tdo_id.ListIndex = snp.Fields(3) & ""
                    frm.cli_nrodoc = snp.Fields(4) & ""
                    frm.cli_te = snp.Fields(5) & ""
                    frm.cli_celular = snp.Fields(6) & ""
                    frm.cli_calle = snp.Fields(7) & ""
                    frm.cli_nro = snp.Fields(8) & ""
                    frm.cli_piso = snp.Fields(9) & ""
                    frm.cli_dpto = snp.Fields(10) & ""
                    frm.cli_seccional = snp.Fields(11) & ""
                    frm.bar_id = snp.Fields(12) & ""
                    frm.cli_obs = snp.Fields(13) & ""
                    If Prop Then BTActualizar frm
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.cli_id.Text = ""
                        frm.cli_nombre.Text = ""
                    End If
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.BlancoCampos
                    If Prop Then BTAgregar frm
                    frm.cli_id.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.cli_id = ""
            frm.cli_nombre = ""
        End If
    End If
End Sub






Public Sub POLIZA_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Codigo = frm.pol_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT pol_fecvig, pol_fecdesde, pol_fecalta, pol_monto_cuota, pol_monto_aseg, pol_obs, cli_id, cli_id_sol, usu_baj_fec, ven_id , cob_id, tpp_id, pol_bloqueo, pol_fecha_fallecimiento, pol_nro_certificado, cab_id FROM POLIZA WHERE POL_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If Prop Then
                        frm.Image1(0).Visible = True
                        frm.Image1(1).Visible = False
                        frm.pol_fecvig = Format(snp.Fields(0) & "", "dd/mm/yyyy")
                        frm.pol_fecdesde = Format(snp.Fields(1) & "", "dd/mm/yyyy")
                        frm.pol_fecalta = Format(snp.Fields(2) & "", "dd/mm/yyyy")
                        frm.pol_monto_cuota = Format(snp.Fields(3) & "", "0.00")
                        frm.pol_monto_aseg = Format(snp.Fields(4) & "", "0.00")
                        frm.pol_obs = snp.Fields(5) & ""
                        frm.cli_id(0) = snp.Fields(6) & ""
                        frm.cli_id(1) = snp.Fields(7) & ""
                        frm.cli_id_LostFocus 0
                        frm.cli_id_LostFocus 1
                        frm.ven_id = snp.Fields(9) & ""
                        frm.cob_id = snp.Fields(10) & ""
                        frm.ven_id_LostFocus
                        frm.cob_id_LostFocus
                        frm.tpp_id = snp.Fields(11) & ""
                        frm.tpp_id_LostFocus
                        frm.pol_bloqueo.Value = IIf(snp.Fields(12) & "" = "SI", 1, 0)
                        frm.pol_nro_certificado = snp!pol_nro_certificado & ""
                        If IsNull(snp!cab_id) Then
                            BTActualizar frm
                            frm.cab_descripcion.ListIndex = -1
                            frm.pol_fecha_fallecimiento = snp!pol_fecha_fallecimiento & ""
                            frm.Image1(1).Visible = False
                            frm.Image1(0).Visible = True
                            frm.lblBaja.Visible = False
                            frm.cAltaNuevamente.Visible = False
                        Else
                            For i = 0 To frm.cab_descripcion.ListCount - 1
                                If frm.cab_descripcion.ItemData(i) = Val(snp!cab_id & "") Then
                                    frm.cab_descripcion.ListIndex = i
                                End If
                            Next
                            frm.pol_fecha_fallecimiento = snp!pol_fecha_fallecimiento & ""
                            BTConsultar frm
                            frm.Image1(1).Visible = True
                            frm.Image1(0).Visible = False
                            frm.lblBaja.Visible = True
                            frm.cAltaNuevamente.Visible = True
                        End If
                        frm.BloquearControles True
                    End If
                Case -1, -2
                    If Prop Then
                        frm.BlancoCampos
                        frm.BloquearControles False
                    End If
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        BTAgregar frm
                        frm.BlancoCampos
                        frm.BloquearControles False
                    End If
                    frm.pol_id.SetFocus
            End Select
        Else
            MsgBox "La póliza debe ser numérica positiva", 48, AppName
            frm.pol_id = ""
            frm.BloquearControles False
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
            frm.BloquearControles False
        End If
    End If
End Sub







Public Sub PERSONAL_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.per_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT per_nombre, per_fecnac, per_fecbaja, tdo_id, per_nrodoc, per_te, per_celular, per_calle, per_nro, per_piso, per_dpto, per_seccional, bar_id, per_obs, per_per_id, niv_id, fun_id , zon_id FROM PERSONAL WHERE per_id=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If Prop Then
                        frm.per_nombre = snp.Fields(0) & ""
                        frm.per_fecnac = snp.Fields(1) & ""
                        frm.per_fecbaja = snp.Fields(2) & ""
                        frm.tdo_id.ListIndex = snp.Fields(3) & ""
                        frm.per_nrodoc = snp.Fields(4) & ""
                        frm.per_te = snp.Fields(5) & ""
                        frm.per_celular = snp.Fields(6) & ""
                        frm.per_calle = snp.Fields(7) & ""
                        frm.per_nro = snp.Fields(8) & ""
                        frm.per_piso = snp.Fields(9) & ""
                        frm.per_dpto = snp.Fields(10) & ""
                        frm.per_seccional = snp.Fields(11) & ""
                        frm.bar_id = snp.Fields(12) & ""
                        frm.per_obs = snp.Fields(13) & ""
                        frm.per_per_id = snp.Fields(14) & ""
                        frm.per_funcion.ListIndex = Val(snp.Fields(16) & "")
                        frm.per_nivel.ListIndex = Val(snp.Fields(15) & "")
                        frm.zon_id = snp.Fields(17) & ""
                        BTActualizar frm
                    Else
                        Select Case Val(snp.Fields(15) & "")
                            Case 1
                                Nivel = "Supervisor"
                            Case 2 'Jefe de Equipo
                                Nivel = "Jefe de Equipo"
                            Case 3 'Para el Vendedor
                                Nivel = "Vendedor"
                        End Select
                        frm.per_nombre = snp.Fields(0) & "-" & Nivel
                    End If
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.per_id.Text = ""
                        frm.per_nombre.Text = ""
                    End If
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.BlancoCampos
                    frm.per_id = ""
                    If Prop Then BTAgregar frm
                    frm.per_id.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.per_id = ""
            frm.per_nombre = ""
        End If
    End If
End Sub







Public Sub COBRADOR_LF(ByRef frm As Form)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.cob_id
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "select per_nombre, per_fecbaja, zona.zon_id, zon_desc from PERSONAL, ZONA WHERE zona.zon_id=PERSONAL.zon_id and per_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case Is > 0
                    frm.cob_nombre = snp!per_nombre & ""
                    frm.zon_id = snp!zon_id & ""
                    frm.zon_desc.Text = snp!zon_desc & ""
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    frm.cob_id = ""
                    frm.cob_nombre = ""
                    frm.zon_id = ""
                    frm.zon_desc = ""
                Case 0
                    MsgBox "Código de cobrador Inexistente", 64, AppName
                    frm.cob_id = ""
                    frm.cob_id.SetFocus
                    frm.cob_nombre = ""
                    frm.zon_id = ""
                    frm.zon_desc = ""
            End Select
        End If
    Else
        frm.cob_id = ""
        frm.cob_nombre = ""
        frm.zon_id = ""
        frm.zon_desc = ""
    End If
End Sub








Public Sub COBRADOR2_LF(ByRef frm As Form)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.per_id
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) And Val(Codigo) > 0 Then
            sSQL$ = "select per_nombre, per_fecbaja, personal.fun_id, fun_cobrador, fun_vendedor, zona.zon_id, zon_desc FROM PERSONAL, FUNCION, ZONA WHERE funcion.fun_id=personal.fun_id and zona.zon_id=personal.zon_id and per_id=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If Val(snp!fun_cobrador & "") = 0 Then
                        MsgBox "El código ingresado no pertenece a personal con atributos de cobrador.", 48, AppName
                        frm.ven_id.SetFocus
                        frm.ven_id = ""
                        frm.per_nombre = ""
                        frm.zon_id = ""
                        frm.zon_desc = ""
                        Exit Sub
                    End If
                    frm.per_nombre = snp!per_nombre & ""
                    frm.zon_id = snp!zon_id & ""
                    frm.zon_desc.Text = snp!zon_desc & ""
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    frm.per_id = ""
                    frm.per_nombre = ""
                    frm.zon_id = ""
                    frm.zon_desc = ""
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.per_id = ""
                    frm.per_id.SetFocus
                    frm.per_nombre = ""
                    frm.zon_id = ""
                    frm.zon_desc = ""
            End Select
        End If
    Else
        frm.per_id = ""
        frm.per_nombre = ""
        frm.zon_id = ""
        frm.zon_desc = ""
    End If
End Sub


Public Sub VENDEDOR_LF(ByRef frm As Form)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.ven_id
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) And Val(Codigo) > 0 Then
            sSQL$ = "select ven_nombre, ven_fecbaja, zona.zon_id, zon_desc, ven_ven_id FROM Vendedor, ZONA WHERE zona.zon_id=Vendedor.zon_id and ven_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If IsNull(snp!ven_ven_id) Then
                        MsgBox "El código ingresado no pertenece a un vendedor sino a un Jefe de Equipo.", 48, AppName
                        frm.ven_id = ""
                        frm.ven_nombre = ""
                        Exit Sub
                    End If
                    frm.ven_nombre = snp!ven_nombre & ""
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    frm.ven_id = ""
                    frm.ven_nombre = ""
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.ven_id = ""
                    frm.ven_id.SetFocus
                    frm.ven_nombre = ""
            End Select
        End If
    Else
        frm.ven_id = ""
        frm.ven_nombre = ""
    End If
End Sub
Public Sub JEFE_LF(ByRef frm As Form)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.ven_id
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) And Val(Codigo) > 0 Then
            sSQL$ = "select ven_nombre, ven_fecbaja, zona.zon_id, zon_desc, ven_ven_id FROM Vendedor, ZONA WHERE zona.zon_id=Vendedor.zon_id and ven_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If Not IsNull(snp!ven_ven_id) Then
                        MsgBox "El código ingresado no pertenece a un jefe de equipo.", 48, AppName
                        frm.ven_id = ""
                        frm.ven_nombre = ""
                        Exit Sub
                    End If
                    frm.ven_nombre = snp!ven_nombre & ""
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    frm.ven_id = ""
                    frm.ven_nombre = ""
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.ven_id = ""
                    frm.ven_id.SetFocus
                    frm.ven_nombre = ""
            End Select
        End If
    Else
        frm.ven_id = ""
        frm.ven_nombre = ""
    End If
End Sub
Public Sub JEFE2_LF(ByRef frm As Form)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.ven_ven_id
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) And Val(Codigo) > 0 Then
            sSQL$ = "select ven_nombre, ven_fecbaja, zona.zon_id, zon_desc, ven_ven_id FROM Vendedor, ZONA WHERE zona.zon_id=Vendedor.zon_id and ven_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If Not IsNull(snp!ven_ven_id) Then
                        MsgBox "El código ingresado no pertenece a un jefe de equipo.", 48, AppName
                        frm.ven_ven_id = ""
                        frm.ven_ven_nombre = ""
                        Exit Sub
                    End If
                    frm.ven_ven_nombre = snp!ven_nombre & ""
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    frm.ven_ven_id = ""
                    frm.ven_ven_nombre = ""
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.ven_ven_id = ""
                    frm.ven_ven_id.SetFocus
                    frm.ven_ven_nombre = ""
            End Select
        End If
    Else
        frm.ven_ven_id = ""
        frm.ven_ven_nombre = ""
    End If
End Sub

Public Sub VENDEDORES_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Dim Nivel As String
    Codigo = frm.ven_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT per_nombre, per_fecnac, per_fecbaja, tdo_id, per_nrodoc, per_te, per_celular, per_calle, per_nro, per_piso, per_dpto, per_seccional, bar_id, per_obs, per_per_id, niv_id, zon_id FROM PERSONAL WHERE per_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    If Prop Then
                        frm.per_nombre = snp.Fields(0) & ""
                        frm.per_fecnac = snp.Fields(1) & ""
                        frm.per_fecbaja = snp.Fields(2) & ""
                        frm.tdo_id.ListIndex = snp.Fields(3) & ""
                        frm.per_nrodoc = snp.Fields(4) & ""
                        frm.per_te = snp.Fields(5) & ""
                        frm.per_celular = snp.Fields(6) & ""
                        frm.per_calle = snp.Fields(7) & ""
                        frm.per_nro = snp.Fields(8) & ""
                        frm.per_piso = snp.Fields(9) & ""
                        frm.per_dpto = snp.Fields(10) & ""
                        frm.per_seccional = snp.Fields(11) & ""
                        frm.bar_id = snp.Fields(12) & ""
                        frm.per_obs = snp.Fields(13) & ""
                        frm.per_per_id = snp.Fields(14) & ""
                        frm.per_nivel.ListIndex = Val(snp.Fields(15) & "")
                        frm.zon_id = snp.Fields(16) & ""
                        BTActualizar frm
                    Else
                        Select Case Val(snp.Fields(15) & "")
                            Case 2 'Jefe de Equipo
                                Nivel = " - Jefe de Equipo"
                            Case 3 'Para el Vendedor
                                Nivel = " - Vendedor"
                        End Select
                        frm.ven_nombre = snp.Fields(0) & Nivel
                    End If
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.ven_id.Text = ""
                        frm.ven_nombre.Text = ""
                    End If
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    frm.ven_id = ""
                    If Prop Then
                        BTAgregar frm
                        frm.BlancoCampos
                    End If
                    frm.ven_id.SetFocus
            End Select
        Else
            frm.ven_id = ""
            frm.ven_nombre = ""
            MSJ frm, "Código Inválido. Debe ser numérico positivo"
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            If Prop Then
                frm.per_id = ""
                frm.per_nombre = ""
            End If
        End If
    End If
End Sub








Public Sub COBRADORES_LF(ByRef frm As Form, ByRef Prop As Boolean)
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    Codigo = frm.cob_id.Text
    If Trim(Codigo) <> "" Then
        If IsNumeric(Codigo) Then
            sSQL$ = "SELECT COB_nombre, COB_fecnac, COB_fecbaja, tdo_id, COB_nrodoc, COB_te, COB_celular, COB_calle, COB_nro, COB_piso, COB_dpto, COB_seccional, bar_id, COB_obs, zon_id FROM COBRADOR WHERE COB_ID=" & XN(Codigo)
            Select Case QRY(sSQL, snp)
                Case 1
                    frm.cob_nombre = snp.Fields(0) & ""
                    frm.zon_id = snp.Fields(14) & ""
                    If Prop Then
                        frm.cob_fecnac = snp.Fields(1) & ""
                        frm.cob_fecbaja = snp.Fields(2) & ""
                        frm.tdo_id.ListIndex = snp.Fields(3) & ""
                        frm.cob_nrodoc = snp.Fields(4) & ""
                        frm.cob_te = snp.Fields(5) & ""
                        frm.cob_celular = snp.Fields(6) & ""
                        frm.cob_calle = snp.Fields(7) & ""
                        frm.cob_nro = snp.Fields(8) & ""
                        frm.cob_piso = snp.Fields(9) & ""
                        frm.cob_dpto = snp.Fields(10) & ""
                        frm.cob_seccional = snp.Fields(11) & ""
                        frm.bar_id = snp.Fields(12) & ""
                        frm.cob_obs = snp.Fields(13) & ""
                        BTActualizar frm
                    End If
                Case -1, -2
                    MsgBox "Error en la Consulta", 48, AppName
                    If Prop Then
                        frm.BlancoCampos
                    Else
                        frm.cob_id.Text = ""
                        frm.cob_nombre.Text = ""
                        frm.zon_id = ""
                        frm.zon_desc = ""
                    End If
                Case 0
                    MsgBox "Código Inexistente", 64, AppName
                    If Prop Then
                        frm.BlancoCampos
                        BTAgregar frm
                    Else
                        frm.cob_id.Text = ""
                        frm.cob_nombre.Text = ""
                        frm.zon_id = ""
                        frm.zon_desc = ""
                    End If
                    frm.cob_id.SetFocus
            End Select
        End If
    Else
        If Prop Then
            frm.BlancoCampos
            BTAgregar frm
        Else
            frm.cob_id = ""
            frm.cob_nombre = ""
            frm.zon_id = ""
            frm.zon_desc = ""
        End If
    End If
End Sub










