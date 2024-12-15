Attribute VB_Name = "Procesos"
Public Function GeneraDigitoVerificador(Numero As String) As String
    Dim AuxMultiplicar As String
    Dim Suma As Integer, R As Integer
    Dim Resto As Currency
    Dim Digito As String
    
    AuxMultiplicar = "131313131313"
    Suma = 0
    For R = 1 To 12
        Suma = Suma + (Val(Mid(Numero, R, 1)) * Val(Mid(AuxMultiplicar, R, 1)))
    Next R
    Resto = Suma Mod 10
    If Resto <> 0 Then
        Digito = Trim(Str(10 - Resto))
        GeneraDigitoVerificador = Numero + Digito
    Else
        GeneraDigitoVerificador = Numero & "0"
    End If
End Function



Public Function CambiarFechaVigencia(Poliza As Long, NuevaFecha As Date) As Integer
On Error GoTo ERRTRAN
    If MsgBox("¿Está seguro que desea cambiar la fecha de vigencia de la póliza?", 36, AppName) = 6 Then
        sSQL(0) = "update poliza set pol_fecvig=" & XD(NuevaFecha) & " where pol_id=" & XN(Poliza)
        Base.Execute sSQL(0), Modo_Paso
    End If
CambiarFechaVigencia = 1
Exit Function
ERRTRAN:
    CambiarFechaVigencia = -1
    ErrManager Err
End Function

Public Function GeneraDebitos() As Integer
'Primero se corre el proceso de imputación de créditos a débitos
'Este proceso debe considerar a partir de un mes de vecimiento
'generar débitos por el monto correspondiente a una cuota de
'póliza. Una vez generados los débitos hay que verificar que
'no exista crédito para esa póliza; de existir debo restar al
'débito el monto del crédito. Si cancelo el débito debo marcar
'que no se imprima recibo sobre ese débito porque ya fue pagado
'probablemente por un recibo manual.
'Como idea pendiente queda el hacer una tabla debgral que maneje
'información sobre cada vez que se tira el proceso
Dim UltNro As Long
Dim snp As Recordset
Dim snp1 As Recordset
Dim Imprimir As String * 2
Dim CantDebitos As Integer
CantDebitos = 0

MP 11
' agrege control para que se generen los debitos
' siempre y cuando no existan para esa poliza
' ni esa fecha contable
' es para el caso en que desbloqueen polizas y
' tengan que generar los debitos de esas polizas
'solamente

'If YaRealizado Then
'    GeneraDebitos = 0
'    Exit Function
'Else
    Dim mDia, fechita As String
    Select Case Mid(Format(FecCont, "dd/mm/yyyy"), 4, 2)
        Case "01", "03", "05", "07", "08", "10", "12"
            mDia = "31"
        Case "02"
            mDia = "28"
        Case Else
            mDia = "30"
    End Select
    fechita = mDia & Mid(Format(FecCont, "dd/mm/yyyy"), 3, 8)
    sSQL(0) = "select pol_id, tpp_recpreimp from poliza, tipo_pago where poliza.tpp_id=tipo_pago.tpp_id and isnull(usu_baj_fec) and isnull(pol_fec_bloqueo) and pol_fecdesde<=" + XDQ(fechita)
    'andres
    'sSQL(0) = "select pol_id, tpp_recpreimp from poliza, tipo_pago where poliza.tpp_id=tipo_pago.tpp_id and isnull(usu_baj_fec) and isnull(pol_fec_bloqueo)"
    If QRY(sSQL(0), snp) > 0 Then
        WKS.BeginTrans
        Do While Not snp.EOF
            'VEO SI ESTA GENERADO EL DEBITO
            sSQL(1) = "select * from debito where pol_id = " & snp!pol_id & " and DEB_FECHA_CONTABLE=" + XDQ(FecCont)
            If QRY(sSQL(1), snp1) > 0 Then
                'debito ya generado
            Else
                CantDebitos = CantDebitos + 1
                UltNro = TraeUltNum("debito")
                If Val(snp!tpp_recpreimp & "") < 0 Then
                    Imprimir = "SI"
                Else
                    Imprimir = "NO"
                End If
                sSQL(0) = "insert into debito (deb_id, pol_id, deb_fecha, deb_importe, deb_saldo, deb_imprimir, deb_fecha_contable, usu_ins_fec, usu_ins_id)" & _
                        " values(" & XN(UltNro) & ", " & XN(snp!pol_id) & ", " & XD(Date) & ", " & XN(MontoCuota) & ", " & XN(MontoCuota) & ", " & XS(Imprimir) & ", " & XD(FecCont) & ", " & XD(Date) & ", " & XN(P_UID) & ")"
                Base.Execute sSQL(0), Modo_Paso
                If ActUltNum("debito") < 1 Then GoTo ERRTRAN
             End If
             snp.MoveNext
        Loop
        'MSJ frmProcesos, "Total de débitos generados: " & snp.RecordCount
        MSJ frmProcesos, "Total de débitos generados: " & CantDebitos
        WKS.CommitTrans
        GeneraDebitos = 1
    Else
        MP 0
        MsgBox "No existen pólizas cargadas sobre las cuales se puedan emitir débitos", 48, AppName
        GeneraDebitos = 0
    End If
'End If
MP 0
Exit Function
ERRTRAN:
    WKS.Rollback
    MP 0
    MsgBox "Ocurrió un error mientras se intentaban generar los débitos correspondientes al mes contable en curso.", 48, AppName
    GeneraDebitos = -1
End Function


Public Sub ImputacionTotal()
    MP 11
    Dim snp As Recordset
    Dim TotPol As Integer
    Set snp = TraerPolizas(4) 'Activas
    If snp.RecordCount > 0 Then
        snp.MoveLast
        TotPol = snp.RecordCount
        snp.MoveFirst
        WKS.BeginTrans
        Do While Not snp.EOF
            If ImputaCredito(snp!pol_id) >= 0 Then
'                MainMDI.flood(0).FloodPercent = (snp.AbsolutePosition / TotPol) * 100
                snp.MoveNext
            Else
                GoTo ERRTRAN
            End If
        Loop
        WKS.CommitTrans
    End If
    snp.Close
    Set snp = Nothing
    MP 0
    Exit Sub
ERRTRAN:
    WKS.Rollback
    MP 0
    MsgBox "Ha ocurrido un error mientras se imputaban créditos a los débitos en la poliza Número " & snp!pol_id & ".", 48, AppName
End Sub

Public Function ImputaCredito(pol_id As Long) As Byte
    Dim I As Byte
    Dim SaldoCredito, SaldoDebito, ImporteImputado As Single
    Dim TotDeb As Integer
    Dim snp0 As Recordset
    Dim snp1 As Recordset
    On Error GoTo ERRTRAN
    sSQL(0) = "select deb_id, pol_id, deb_fecha, deb_importe, deb_saldo from debito where pol_id=" & XN(pol_id) & " and deb_saldo>0 order by deb_fecha_contable, deb_fecha"
    If QRY(sSQL(0), snp0) > 0 Then
        snp0.MoveLast
        TotDeb = snp0.RecordCount
        snp0.MoveFirst
        Do While Not snp0.EOF
'            MainMDI.flood(1).FloodPercent = (snp0.AbsolutePosition / TotDeb) * 100
            SaldoDebito = snp0!DEB_SALDO
            'Verifico si tiene crédito
            Set snp1 = PolizaTieneCredito(pol_id)
            If snp1.RecordCount > 0 Then
                SaldoCredito = Val(snp1!cre_saldo & "")
                'Como tiene crédito imputo el crédito hasta agotarlo o hasta agotar el débito y actualizando saldos en cada imputación
                If SaldoCredito > SaldoDebito Then
                    ImporteImputado = SaldoDebito
                    SaldoCredito = SaldoCredito - ImporteImputado
                    SaldoDebito = 0
                Else
                    'Si es menor o igual todo el saldo es usado. el Importe Imputado = Credito y el saldodebito lo que quedo sin cubrir
                    ImporteImputado = SaldoCredito
                    SaldoDebito = SaldoDebito - ImporteImputado
                    SaldoCredito = 0
                End If
                sSQL(0) = "insert into imputacion (deb_id,cre_id,imp_importe,imp_fecha, imp_fecha_contable) " & _
                          "values(" & XN(snp0!deb_id) & ", " & XN(snp1!cre_id) & ", " & _
                           XN(ImporteImputado) & ", " & XD(Date) & ", " & XD(FecCont) & ")"
                sSQL(1) = "update debito set deb_saldo=" & XN(SaldoDebito) & " where deb_id=" & XN(snp0!deb_id)
                sSQL(2) = "update credito set cre_saldo=" & XN(SaldoCredito) & " where cre_id=" & XN(snp1!cre_id)
                For I = 0 To 2
                    Base.Execute sSQL(I), Modo_Paso
                Next
                Debug.Print pol_id
            Else
                ImputarCredito = 1
                MP 0
                Exit Function
            End If
            snp1.Close
            Set snp1 = Nothing
            If SaldoDebito = 0 Then
                snp0.MoveNext
            End If
        Loop
    End If
    ImputarCredito = 1
    MP 0
Exit Function
ERRTRAN:
    ImputarCredito = -1
End Function


Public Function ReciboManualExiste(rmn_id As Long) As Integer
    Dim snp As Recordset
    Dim sSQL As String
    sSQL = "select rmn_id from rec_manual where rmn_id=" & XN(rmn_id)
    ReciboManualExiste = QRY(sSQL, snp)
End Function

Public Function TraerPolizas(Tipo As Integer) As Recordset
    Dim snp As Recordset
    Dim RC As Variant
    Dim Estado As String
    'Activa = 1
    'Suspendida = 2
    'Baja = 3
    'Todas = 4
    Select Case Tipo
        Case 1
            Estado = "where isnull(usu_baj_fec) and pol_bloqueo='NO'"
        Case 2
            Estado = "where isnull(usu_baj_fec) and pol_bloqueo='NO'"
        Case 3
            Estado = "where not isnull(usu_baj_fec)"
        Case 4
            Estado = ""
    End Select
    sSQL(0) = "select pol_id from poliza " & Estado
    On Error GoTo nada
    RC = QRY(sSQL(0), snp)
    Set TraerPolizas = snp
    Exit Function
nada:
End Function

Public Function PolizaTieneCredito(pol_id As Long) As Recordset
Dim snp As Recordset
Dim RC As Variant
'Devuelve un Recordset con los datos de los créditos con saldos mayores a cero
sSQL(0) = "select cre_id, cre_saldo, cre_fecha, cre_fecha_contable, cre_rec_tipo, cre_rec_numero " & _
        "from credito " & _
        "where pol_id=" & XN(pol_id) & " and cre_saldo > 0 order by cre_fecha_contable, cre_saldo"
On Error GoTo nada
RC = QRY(sSQL(0), snp)
    Set PolizaTieneCredito = snp
    Exit Function
nada:
End Function


Private Function YaRealizado() As Boolean
'Verifico que ya no existan dentro de la tabla de débitos
'débitos emitidos para el actual mes contable
Dim snp As Recordset
sSQL(0) = "select max(deb_fecha_contable) as FecCont from debito"
Select Case QRY(sSQL(0), snp)
    Case 0
        YaRealizado = False
    Case 1
        If snp!FecCont & "" <> "" Then
            If CDate(snp!FecCont & "") < CDate(FecCont) Then
                YaRealizado = False
            Else
                MP 0
                YaRealizado = True
                MsgBox "Este proceso de generación de débitos ya ha sido realizado para el presente mes contable", 48, AppName
                Exit Function
            End If
        Else
            YaRealizado = False
        End If
    Case Is < 0
        YaRealizado = True
End Select
snp.Close
Set snp = Nothing
End Function


Public Sub GenerarDebitosEntre(Poliza As Long, Desde As Date, Hasta As Date, Pol_monto_cuota As Single)
    Dim UltNro As Long
    Do While Desde <= Hasta
        UltNro = TraeUltNum("debito")
        sSQL(0) = "insert into debito (deb_id, pol_id, deb_fecha, deb_importe, deb_saldo, deb_imprimir, deb_fecha_contable, usu_ins_fec, usu_ins_id)" & _
                " values(" & XN(UltNro) & ", " & XN(Poliza) & ", " & XD(Desde) & ", " & XN(Pol_monto_cuota) & ", " & XN(Pol_monto_cuota) & ", " & XS("NO") & ", " & XD(Desde) & ", " & XD(Date) & ", " & XN(P_UID) & ")"
        Base.Execute sSQL(0), Modo_Paso
        If ActUltNum("debito") < 1 Then MsgBox "ERROR"
        Desde = DateAdd("m", 1, Desde)
    Loop
End Sub
