Attribute VB_Name = "Funciones"
Option Explicit
Dim Letrero As String

Sub DeshabilitarMenu(Frm As Form)
 Dim i As Integer
    For i = 0 To Menu.Controls.Count - 1
        If TypeName(Menu.Controls(i)) = "Menu" And UCase(Left(Frm.Controls(i).Name, 7)) <> "MNURAYA" Then
           Menu.Controls(i).Enabled = False
        End If
    Next
End Sub

Public Sub EliminarFilasDeGrilla(Grilla As MSFlexGrid)
'Borra totalmente una grilla sin dejar Rows
    Grilla.Rows = 1
    Do While Grilla.Rows > 1
        Grilla.RemoveItem 1
    Loop
End Sub

Public Sub LimpiarFilasDeGrilla(Grilla As MSFlexGrid, Optional Fila As Long)
   Dim i As Integer
   Dim J As Integer
'Borra datos de una o todas las filas
    Dim CantidadDeFilas As Integer
    Dim ActFila As Integer
    ActFila = Grilla.row
    If Fila <> 0 Then
        CantidadDeFilas = Fila + 1
    Else
        Fila = 1
        CantidadDeFilas = Grilla.Rows
    End If
    For i = Fila To CantidadDeFilas - 1
        Grilla.row = i
        For J = 0 To Grilla.Cols - 1
            Grilla.Col = J
            Grilla.Text = ""
        Next
    Next
    Grilla.row = ActFila
End Sub

Public Function ValidarPorcentaje(Control As Control) As Boolean
    If CDbl(Control.Text) <= 100 Then
       Control.Text = Format(Control, "0.00")
       ValidarPorcentaje = True
    Else
       MsgBox "Error, Porcentaje mayor al 100%", 16, TIT_MSGBOX
       Control.SetFocus
       SelecTexto Control
       ValidarPorcentaje = False
    End If
End Function

Public Function Valido_Importe(mTEXTO As String) As String
    Valido_Importe = IIf(Trim(mTEXTO) = "", "0,00", Format(mTEXTO, "#,##0.00"))
End Function

Public Function SeteoImpresora(Papel As Integer, Orientacion As Integer, Modo_Escala As Integer, Calidad_Impresion As Integer, Fuente As String, Fuente_Tamano As Integer, Fuente_Negrita As Boolean, Ancho_Impresion As Integer, Largo_Impresion As Double, Optional Ancho_Recuadro As Double) As Boolean
    '*************************************************************************************
    'A) CHEQUES:
    '   1) Llamada a la función: SeteoImpresora(256, 1, 7, -1, "Roman 10cpi", 10, False, 12220, 7950)
    '   2) Configuración de la impresora:
    '       * Driver              : Generic IBM Graphics 9pin
    '       * Paper Size          : Custom
    '               Width         : 1560
    '               Length        : 710
    '               Unit          : 0.1 milimeters
    '
    'B) OBLEAS
    '   1) LLamada a la función: SeteoImpresora(256, 2, 6, -4, "Times New Roman", 8, False, 0, 0, 3)
    '   2) Configuración de la impresora:
    '       * Paper Size          : CUSTOM
    '             Width           : 560
    '             Lenght          : 1000
    '             Unit            : 0,1 MILIMETERS
    '       * Orientation         : LANDSCAPE
    '       * Paper Source        : TOF BACKUP ENABLED
    '       * Media Choice        : SPEED 1.5 TIPS
    '       * Graphics Resolutions: 200 DPI
    '       * Dithering           : NONE
    '       * Intensity           : 100
    '       * Print Quality       : DENSITY 8
    '
    'C) TALON DE CONTROL
    '   1)LLamado a la funcion:Call SeteoImpresora(256, 1, 6, -1, "Roman 10cpi", 10, False, 220, 75)
    '   2) Configuración de la impresora:
    '       * Driver              : Generic IBM Graphics 9pin
    '       * Paper Size          : Custom
    '               Width         :
    '               Length        :
    '               Unit          : 0.1 milimeters
    '
    '***************************************************************************************
    On Error GoTo ErrorPrint
    Printer.PaperSize = Papel
    'Constante        Valor  Descripción
    'vbPRPSLetter       1    Carta, 216 x 279 mm
    'vbPRPSLegal        5    Oficio, 216 x 356 mm
    'vbPRPSA4           9    A4, 210 x 297 mm
    'vbPRPSUser        256   Definido por el usuario

    Printer.Orientation = Orientacion
    'Constante                   Descripción
    'VtOrientationHorizontal     El texto se muestra horizontalmente.
    'VtOrientationVertical       Las letras del texto se dibujan una encima de otra de arriba a abajo.
    'VtOrientationUp             El texto se rota para que se lea de abajo a arriba.
    'VtOrientationDown           El texto se rota para que se lea de arriba a abajo.
    
    Printer.ScaleMode = Modo_Escala
    'Constante    Valor   Descripción
    'vbUser         0      Indica que una o más de las propiedades ScaleHeight, ScaleWidth, ScaleLeft y ScaleTop tienen valores personalizados.
    'VbTwips        1      (Predeterminado) Twip (1440 twips por pulgada lógica; 567 twips por centímetro lógico).
    'VbPoints       2      Punto (72 puntos por pulgada lógica).
    'VbPixels       3      Píxel (la unidad mínima de la resolución del monitor o la impresora).
    'vbCharacters   4      Carácter (horizontal = 120 twips por unidad; vertical = 240 twips por unidad).
    'VbInches       5      Pulgada.
    'VbMillimeters  6      Milímetro.
    'VbCentimeters  7      Centímetro.
    
    Printer.PrintQuality = Calidad_Impresion
    'Constante     Valor   Descripción
    'vbPRPQDraft    -1      Resolución borrador
    'vbPRPQLow      -2      Resolución baja
    'vbPRPQMedium   -3      Resolución media
    'vbPRPQHigh     -4      Resolución alta

    Printer.Font = Fuente
    Printer.FontSize = Fuente_Tamano
    
    Printer.FontBold = Fuente_Negrita
    'Valor   Descripción
    'True    Activa el formato de negrita.
    'False   (Predeterminado) Desactiva el formato de negrita.
    
    If Largo_Impresion > 0 Then
       Printer.Height = Largo_Impresion
    End If
    If Ancho_Impresion > 0 Then
       Printer.Width = Ancho_Impresion
    End If
    If Not IsNull(Ancho_Recuadro) And Ancho_Recuadro > 0 Then
       Printer.DrawWidth = Ancho_Recuadro
    End If
    SeteoImpresora = True
    'AjI = 0
    On Error GoTo 0
    Exit Function
    
ErrorPrint:
    SeteoImpresora = False
    On Error GoTo 0
End Function

Public Function AgregoCtaCteCliente(Cliente As String, TipoCom As String, _
                                    NroComp As String, NroSuc As String, _
                                    Rep As String, FechaComp As String, _
                                    TotalCom As String, DebHab As String, FechaCtaCTe As String) As String
                                    
    'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
    sql = "INSERT INTO CTA_CTE_CLIENTE"
    sql = sql & "(CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,REP_CODIGO,COM_FECHA,"
    sql = sql & "COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,"
    sql = sql & "CTA_CTE_DH,CTA_CTE_FECHA,COM_NUMEROTXT)"
    sql = sql & " VALUES ("
    sql = sql & XN(Cliente) & ","
    sql = sql & XN(TipoCom) & ","
    sql = sql & XN(NroComp) & ","
    sql = sql & XN(NroSuc) & ","
    sql = sql & XN(Rep) & ","
    sql = sql & XDQ(FechaComp) & ","
    sql = sql & XN(TotalCom) & ","
    If DebHab = "D" Then
        sql = sql & XN(TotalCom) & ","
        sql = sql & "0.00" & ","
    Else
        sql = sql & "0.00" & ","
        sql = sql & XN(TotalCom) & ","
    End If
    sql = sql & XS(DebHab) & ","
    sql = sql & XDQ(FechaCtaCTe) & ","
    sql = sql & XS(Format(NroComp, "00000000")) & ")"
    
    AgregoCtaCteCliente = sql
End Function

Public Function QuitoCtaCteCliente(Cliente As String, TipoCom As String, _
                                    NroComp As String, NroSuc As String, Rep As String) As String
    'BORO DE LA CUENTA CORRIENTE DEL CLIENTE
    sql = "DELETE FROM CTA_CTE_CLIENTE"
    sql = sql & " WHERE"
    sql = sql & " CLI_CODIGO=" & XN(Cliente)
    sql = sql & " AND TCO_CODIGO=" & XN(TipoCom)
    sql = sql & " AND COM_NUMERO=" & XN(NroComp)
    sql = sql & " AND COM_SUCURSAL=" & XN(NroSuc)
    sql = sql & " AND REP_CODIGO=" & XN(Rep)
    QuitoCtaCteCliente = sql
End Function

Public Function AgregoCtaCteProveedores(TipoProv As String, Proveedor As String, TipoCom As String, _
                                    NroSuc As String, NroComp As String, FechaComp As String, _
                                    TotalCom As String, DebHab As String, FechaCtaCTe As String) As String
                                    
    'ACTUALIZO LA CUENTA CORRIENTE DEL PROVEEDOR
    sql = "INSERT INTO CTA_CTE_PROVEEDORES"
    sql = sql & "(TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,COM_SUCURSAL,COM_NUMERO,"
    sql = sql & "COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,"
    sql = sql & "CTA_CTE_DH,CTA_CTE_FECHA)"
    sql = sql & " VALUES ("
    sql = sql & XN(TipoProv) & ","
    sql = sql & XN(Proveedor) & ","
    sql = sql & XN(TipoCom) & ","
    sql = sql & XS(NroSuc) & ","
    sql = sql & XS(NroComp) & ","
    sql = sql & XDQ(FechaComp) & ","
    sql = sql & XN(TotalCom) & ","
    If DebHab = "D" Then
        sql = sql & XN(TotalCom) & ","
        sql = sql & "0.00" & ","
    Else
        sql = sql & "0.00" & ","
        sql = sql & XN(TotalCom) & ","
    End If
    sql = sql & XS(DebHab) & ","
    sql = sql & XDQ(FechaCtaCTe) & ")"
    AgregoCtaCteProveedores = sql
End Function

Public Function QuitoCtaCteProveedores(TipoProv As String, Proveedor As String, TipoCom As String, _
                                    NroSuc As String, NroComp As String) As String
    'BORO DE LA CUENTA CORRIENTE DEL CLIENTE
    sql = "DELETE FROM CTA_CTE_PROVEEDORES"
    sql = sql & " WHERE"
    sql = sql & " TPR_CODIGO=" & XN(TipoProv)
    sql = sql & " AND PROV_CODIGO=" & XN(Proveedor)
    sql = sql & " AND TCO_CODIGO=" & XN(TipoCom)
    sql = sql & " AND COM_SUCURSAL=" & XS(NroSuc)
    sql = sql & " AND COM_NUMERO=" & XS(NroComp)
    QuitoCtaCteProveedores = sql
End Function

Public Sub BuscoEstado(Codigo As Integer, Control As Label)
    Set Rec4 = New ADODB.Recordset
    sql = "SELECT EST_DESCRI"
    sql = sql & " FROM ESTADO_DOCUMENTO"
    sql = sql & " WHERE"
    sql = sql & " EST_CODIGO=" & Codigo
    Rec4.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec4.EOF = False Then
        Control.Caption = Rec4!EST_DESCRI
        
        Select Case Codigo
        Case 2
            Control.ForeColor = &HFF&
        Case 1
            Control.ForeColor = &HFF0000
        Case 3
            Control.ForeColor = &H0&
        Case Else
            Control.ForeColor = &HFF0000
        End Select
    End If
    Rec4.Close
End Sub

Public Sub BuscoNroSucursal()
    Set rec = New ADODB.Recordset
    sql = "SELECT SUCURSAL"
    sql = sql & " FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Sucursal = Format(rec!Sucursal, "0000")
    End If
    rec.Close
End Sub

Public Sub ActualizoNumeroComprobantes(Repre As Integer, TipCom As Integer, NroCom As String)
    Set Rec3 = New ADODB.Recordset
    sql = "SELECT *"
    sql = sql & " FROM PARAMETROS"
    Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec3.EOF = False Then
        If Rec3!REP_CODIGO = Repre Then
            Select Case TipCom
                Case 0 'REMITO
                    sql = "UPDATE PARAMETROS SET NRO_REMITO=" & XN(NroCom)
                Case 1, 4, 7 'FACTURA A, 'NOTA DE CREDITO A, 'NOTA DE DEBITO A
                    sql = "UPDATE PARAMETROS SET FACTURA_A=" & XN(NroCom)
                Case 2, 5, 8 'FACTURA B, NOTA DE CREDITO B, NOTA DE DEBITO B
                    sql = "UPDATE PARAMETROS SET FACTURA_B=" & XN(NroCom)
                Case 3 'FACTURA C
                'Case 4 'NOTA DE CREDITO A
                '    sql = "UPDATE PARAMETROS SET NOTA_CREDITO_A=" & XN(NroCom)
                'Case 5 'NOTA DE CREDITO B
                '    sql = "UPDATE PARAMETROS SET NOTA_CREDITO_B=" & XN(NroCom)
                'Case 6 'NOTA DE CREDITO C
                'Case 7 'NOTA DE DEBITO A
                '    sql = "UPDATE PARAMETROS SET NOTA_DEBITO_A=" & XN(NroCom)
                'Case 8 'NOTA DE DEBITO B
                '    sql = "UPDATE PARAMETROS SET NOTA_DEBITO_B=" & XN(NroCom)
                Case 9 'NOTA DE DEBITO C
                Case 10 'RECIBO A
                    sql = "UPDATE PARAMETROS SET RECIBO_A=" & XN(NroCom)
                Case 11 'RECIBO B
                    sql = "UPDATE PARAMETROS SET RECIBO_B=" & XN(NroCom)
                Case 12 'RECIBO C
            End Select
                DBConn.Execute sql
        End If
        'SI SE TRATA DE VIÑA MAIPU S.A.
        If Rec3!REP_CODIGO_SUC2 = Repre Then
            Select Case TipCom
                Case 0 'REMITO
                    sql = "UPDATE PARAMETROS SET NRO_REMITO_SUC2=" & XN(NroCom)
                Case 1, 4, 7 'FACTURA A, 'NOTA DE CREDITO A, 'NOTA DE DEBITO A
                    sql = "UPDATE PARAMETROS SET FACTURA_A_SUC2=" & XN(NroCom)
                Case 2, 5, 8 'FACTURA B, NOTA DE CREDITO B, NOTA DE DEBITO B
                    sql = "UPDATE PARAMETROS SET FACTURA_B_SUC2=" & XN(NroCom)
                Case 3 'FACTURA C
                'Case 4 'NOTA DE CREDITO A
                '    sql = "UPDATE PARAMETROS SET NOTA_CREDITO_A_SUC2=" & XN(NroCom)
                'Case 5 'NOTA DE CREDITO B
                '    sql = "UPDATE PARAMETROS SET NOTA_CREDITO_B_SUC2=" & XN(NroCom)
                Case 6 'NOTA DE CREDITO C
                'Case 7 'NOTA DE DEBITO A
                '    sql = "UPDATE PARAMETROS SET NOTA_DEBITO_A_SUC2=" & XN(NroCom)
                'Case 8 'NOTA DE DEBITO B
                '    sql = "UPDATE PARAMETROS SET NOTA_DEBITO_B_SUC2=" & XN(NroCom)
                Case 9 'NOTA DE DEBITO C
                Case 10 'RECIBO A
                    sql = "UPDATE PARAMETROS SET RECIBO_A_SUC2=" & XN(NroCom)
                Case 11 'RECIBO B
                    sql = "UPDATE PARAMETROS SET RECIBO_B_SUC2=" & XN(NroCom)
                Case 12 'RECIBO C
            End Select
                DBConn.Execute sql
        End If
        'SI SE TRATA DE PEDRO CARRICONDO E HIJOS S.R.L.
        If Rec3!REP_CODIGO_SUC3 = Repre Then
            Select Case TipCom
                Case 0 'REMITO
                    sql = "UPDATE PARAMETROS SET NRO_REMITO_SUC3=" & XN(NroCom)
                Case 1, 4, 7 'FACTURA A, 'NOTA DE CREDITO A, 'NOTA DE DEBITO A
                    sql = "UPDATE PARAMETROS SET FACTURA_A_SUC3=" & XN(NroCom)
                Case 2, 5, 8 'FACTURA B, NOTA DE CREDITO B, NOTA DE DEBITO B
                    sql = "UPDATE PARAMETROS SET FACTURA_B_SUC3=" & XN(NroCom)
                Case 3 'FACTURA C
                'Case 4 'NOTA DE CREDITO A
                '    sql = "UPDATE PARAMETROS SET NOTA_CREDITO_A_SUC3=" & XN(NroCom)
                'Case 5 'NOTA DE CREDITO B
                '    sql = "UPDATE PARAMETROS SET NOTA_CREDITO_B_SUC3=" & XN(NroCom)
                'Case 6 'NOTA DE CREDITO C
                'Case 7 'NOTA DE DEBITO A
                '    sql = "UPDATE PARAMETROS SET NOTA_DEBITO_A_SUC3=" & XN(NroCom)
                'Case 8 'NOTA DE DEBITO B
                '    sql = "UPDATE PARAMETROS SET NOTA_DEBITO_B_SUC3=" & XN(NroCom)
                Case 9 'NOTA DE DEBITO C
                Case 10 'RECIBO A
                    sql = "UPDATE PARAMETROS SET RECIBO_A_SUC3=" & XN(NroCom)
                Case 11 'RECIBO B
                    sql = "UPDATE PARAMETROS SET RECIBO_B_SUC3=" & XN(NroCom)
                Case 12 'RECIBO C
            End Select
                DBConn.Execute sql
        End If
    End If
    Rec3.Close
    Set Rec3 = Nothing
End Sub

Public Function BuscoUltimoNumeroComprobante(Repre As Integer, TipCom As Integer) As String
    Set Rec3 = New ADODB.Recordset
    sql = "SELECT *"
    sql = sql & " FROM PARAMETROS"
    Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec3.EOF = False Then
        BuscoUltimoNumeroComprobante = Format("1", "00000000")
        'SI SE TRATA DE ESTILO S.R.L.
        If Rec3!REP_CODIGO = Repre Then
            Select Case TipCom
                Case 0 'REMITO
                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NRO_REMITO + 1), "00000000")
                Case 1, 4, 7 'FACTURA A, NOTA DE CREDITO A, NOTA DE DEBITO A
                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!FACTURA_A + 1), "00000000")
                Case 2, 5, 8 'FACTURA B, NOTA DE CREDITO B, NOTA DE DEBITO B
                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!FACTURA_B + 1), "00000000")
                Case 3 'FACTURA C
                'Case 4 'NOTA DE CREDITO A
                    'BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_CREDITO_A + 1), "00000000")
                'Case 5 'NOTA DE CREDITO B
                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_CREDITO_B + 1), "00000000")
                Case 6 'NOTA DE CREDITO C
                    MsgBox "No hay Notas de Crédito del tipo C", vbExclamation, TIT_MSGBOX
                    BuscoUltimoNumeroComprobante = ""
                'Case 7 'NOTA DE DEBITO A
                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_DEBITO_A + 1), "00000000")
                'Case 8 'NOTA DE DEBITO B
                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_DEBITO_B + 1), "00000000")
                Case 9 'NOTA DE DEBITO C
                Case 10 'RECIBO A
                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!RECIBO_A + 1), "00000000")
                Case 11 'RECIBO B
                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!RECIBO_B + 1), "00000000")
                Case 12 'RECIBO C
            End Select
        End If
        'SI SE TRATA DE PEDRO CARRICONDO E HIJOS S.R.L.
'        If Rec3!REP_CODIGO_SUC2 = Repre Then
'            Select Case TipCom
'                Case 0 'REMITO
'                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NRO_REMITO_SUC2 + 1), "00000000")
'                Case 1, 4, 7 'FACTURA A, NOTA DE CREDITO A, NOTA DE DEBITO A
'                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!FACTURA_A_SUC2 + 1), "00000000")
'                Case 2, 5, 8 'FACTURA B, NOTA DE CREDITO B, NOTA DE DEBITO B
'                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!FACTURA_B_SUC2 + 1), "00000000")
'                Case 3 'FACTURA C
'                'Case 4 'NOTA DE CREDITO A
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_CREDITO_A_SUC2 + 1), "00000000")
'                'Case 5 'NOTA DE CREDITO B
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_CREDITO_B_SUC2 + 1), "00000000")
'                Case 6 'NOTA DE CREDITO C
'                    MsgBox "No hay Notas de Crédito del tipo C", vbExclamation, TIT_MSGBOX
'                    BuscoUltimoNumeroComprobante = ""
'                'Case 7 'NOTA DE DEBITO A
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_DEBITO_A_SUC2 + 1), "00000000")
'                'Case 8 'NOTA DE DEBITO B
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_DEBITO_B_SUC2 + 1), "00000000")
'                Case 9 'NOTA DE DEBITO C
'                Case 10 'RECIBO A
'                    'BuscoUltimoNumeroComprobante = Format(CStr(Rec3!RECIBO_A_SUC2 + 1), "00000000")
'                Case 11 'RECIBO B
'                    'BuscoUltimoNumeroComprobante = Format(CStr(Rec3!RECIBO_B_SUC2 + 1), "00000000")
'                Case 12 'RECIBO C
'            End Select
'        End If
'        'SI SE TRATA DE VIÑA MAIPU S.A.
'        If Rec3!REP_CODIGO_SUC3 = Repre Then
'            Select Case TipCom
'                Case 0 'REMITO
'                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NRO_REMITO_SUC3 + 1), "00000000")
'                Case 1, 4, 7 'FACTURA A, NOTA DE CREDITO A, NOTA DE DEBITO A
'                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!FACTURA_A_SUC3 + 1), "00000000")
'                Case 2, 5, 8 'FACTURA B, NOTA DE CREDITO B, NOTA DE DEBITO B
'                    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!FACTURA_B_SUC3 + 1), "00000000")
'                Case 3 'FACTURA C
'                'Case 4 'NOTA DE CREDITO A
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_CREDITO_A_SUC3 + 1), "00000000")
'                'Case 5 'NOTA DE CREDITO B
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_CREDITO_B_SUC3 + 1), "00000000")
'                Case 6 'NOTA DE CREDITO C
'                    MsgBox "No hay Notas de Crédito del tipo C", vbExclamation, TIT_MSGBOX
'                    BuscoUltimoNumeroComprobante = ""
'                'Case 7 'NOTA DE DEBITO A
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_DEBITO_A_SUC3 + 1), "00000000")
'                'Case 8 'NOTA DE DEBITO B
'                '    BuscoUltimoNumeroComprobante = Format(CStr(Rec3!NOTA_DEBITO_B_SUC3 + 1), "00000000")
'                Case 9 'NOTA DE DEBITO C
'                Case 10 'RECIBO A
'                    'BuscoUltimoNumeroComprobante = Format(CStr(Rec3!RECIBO_A_SUC3 + 1), "00000000")
'                Case 11 'RECIBO B
'                    'BuscoUltimoNumeroComprobante = Format(CStr(Rec3!RECIBO_B_SUC3 + 1), "00000000")
'                Case 12 'RECIBO C
'            End Select
'        End If
    End If
    Rec3.Close
    Set Rec3 = Nothing
End Function

Public Sub CargoComboBox(Combo As ComboBox, Tabla As String, CampoCod As String, CampoDes As String, Optional Ordenar As String)
    Set Rec4 = New ADODB.Recordset
    If Ordenar = "" Then
        Ordenar = "D"
    End If
    sql = "SELECT " & Trim(CampoCod) & "," & Trim(CampoDes)
    sql = sql & " FROM " & Trim(Tabla)
    If Ordenar = "D" Then
        sql = sql & " ORDER BY " & Trim(CampoDes)
    Else
        sql = sql & " ORDER BY " & Trim(CampoCod)
    End If
    Rec4.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec4.EOF = False Then
        Do While Rec4.EOF = False
            Combo.AddItem Trim(Rec4.Fields(1))
            Combo.ItemData(Combo.NewIndex) = Trim(Rec4.Fields(0))
            Rec4.MoveNext
        Loop
    End If
    Rec4.Close
    Set Rec4 = Nothing
End Sub

Public Sub CargoComboBoxItemData(Combo As ComboBox, mConsulta As String)
    Set Rec4 = New ADODB.Recordset
    Rec4.Open mConsulta, DBConn, adOpenStatic, adLockOptimistic
    If Rec4.EOF = False Then
        Do While Rec4.EOF = False
            Combo.AddItem Trim(Rec4.Fields(1))
            Combo.ItemData(Combo.NewIndex) = Trim(Rec4.Fields(0))
            Rec4.MoveNext
        Loop
    End If
    Rec4.Close
    Set Rec4 = Nothing
End Sub



