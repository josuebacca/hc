VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "fecha32.ocx"
Begin VB.Form frmMovimientoStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimiento de Stock"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameImpresora 
      Caption         =   "impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   20
      Top             =   2115
      Width           =   6690
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   210
         Width           =   1665
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2385
         TabIndex        =   7
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   315
         Width           =   585
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1725
      Top             =   3045
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmMovimientoStock.frx":0000
      Height          =   720
      Left            =   5040
      Picture         =   "frmMovimientoStock.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2910
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   720
      Left            =   5910
      Picture         =   "frmMovimientoStock.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2910
      Width           =   825
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   720
      Left            =   4185
      Picture         =   "frmMovimientoStock.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2910
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ver por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   6675
      Begin VB.CheckBox chkFechaRemito 
         Caption         =   "Tener en cuenta fecha para Remitos pendientes"
         Height          =   345
         Left            =   1410
         TabIndex        =   5
         Top             =   1620
         Width           =   4335
      End
      Begin VB.ComboBox cboStock 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   3480
      End
      Begin VB.TextBox txtdescri 
         Height          =   330
         Left            =   1410
         TabIndex        =   1
         Top             =   585
         Width           =   4620
      End
      Begin VB.TextBox txtcodigo 
         Height          =   330
         Left            =   450
         TabIndex        =   0
         Top             =   585
         Width           =   930
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Height          =   330
         Left            =   6090
         MaskColor       =   &H000000FF&
         Picture         =   "frmMovimientoStock.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Buscar Producto"
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   3750
         TabIndex        =   4
         Top             =   1320
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   1410
         TabIndex        =   3
         Top             =   1320
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2415
         TabIndex        =   18
         Top             =   345
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Stock:"
         Height          =   195
         Left            =   870
         TabIndex        =   17
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   255
         TabIndex        =   16
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   2700
         TabIndex        =   15
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1470
         TabIndex        =   14
         Top             =   345
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   525
         TabIndex        =   13
         Top             =   345
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2205
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   19
      Top             =   3135
      Width           =   750
   End
End
Attribute VB_Name = "frmMovimientoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 2
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtcodigo.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtdescri.Text = frmBuscar.grdBuscar.Text
        cboStock.SetFocus
    Else
        txtcodigo.SetFocus
    End If
End Sub

Private Sub Cierre_Stock()
    sql = "SELECT CTK_FECHA,PTO_CODIGO,SUM(CTK_FISICO) AS FISICO"
    sql = sql & " FROM CIERRE_STOCK"
    sql = sql & " WHERE"
    sql = sql & " CTK_FECHA="
    '----INICIO SUBCONCULTA----
    sql = sql & " (SELECT MAX(C.CTK_FECHA)AS MINFECHA FROM CIERRE_STOCK C"
    If FechaDesde <> "" Then sql = sql & " WHERE C.CTK_FECHA<=" & XDQ(FechaDesde.Text)
    If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
        sql = sql & " AND STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex)) & ")"
    Else
        sql = sql & ")"
    End If
    '----FIN SUBCONSULTA-------
    If txtcodigo.Text <> "" Then
        sql = sql & " AND PTO_CODIGO=" & XN(txtcodigo.Text)
    End If
    If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
        sql = sql & " AND STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
    End If
    sql = sql & " GROUP BY CTK_FECHA, PTO_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_MOVIMIENTO_STOCK (FECHA,NRO_REMITO,NRO_ENTRASALE,"
            sql = sql & "ENTRADA,SALIDA,PENDIENTE,CLIENTE,PRODUCTO,"
            sql = sql & "ORDENAR,OBSERVACION) VALUES ("
            sql = sql & XDQ(rec!CTK_FECHA) & ","
            sql = sql & XS("") & ","
            sql = sql & XS("") & ","
            sql = sql & XN(rec!FISICO) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN("0") & ","
            sql = sql & XS("SALDO INICIAL DEL MOVIMIENTO") & ","
            sql = sql & XN(rec!PTO_CODIGO) & ","
            sql = sql & "'C',"
            sql = sql & "'CIERRE DE STOCK DEL - '" & "'" & (rec!CTK_FECHA) & "')"
            DBConn.Execute sql
            
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub Entrada_Deposito()
    sql = "SELECT EP.EPR_CODIGO, EP.EPR_FECHA, EP.REP_CODIGO, EP.EPR_NROSUCREM,"
    sql = sql & " EP.EPR_NROREM, EP.STK_CODIGO, EP.EPR_OBSERVACIONES, EP.EST_CODIGO,"
    sql = sql & " DEP.PTO_CODIGO, P.PTO_DESCRI, DEP.DEP_CANTIDAD,EP.CLI_CODIGO"
    sql = sql & " FROM ENTRADA_PRODUCTO EP, DETALLE_ENTRADA_PRODUCTO DEP, PRODUCTO P"
    sql = sql & " WHERE"
    sql = sql & " EP.EPR_CODIGO = DEP.EPR_CODIGO"
    sql = sql & " AND DEP.PTO_CODIGO = P.PTO_CODIGO"
    sql = sql & " AND EP.EST_CODIGO = 3"
    If txtcodigo.Text <> "" Then
        sql = sql & " AND DEP.PTO_CODIGO=" & XN(txtcodigo.Text)
    End If
    If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
        sql = sql & " AND EP.STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
    End If
    If FechaDesde <> "" Then sql = sql & " AND EP.EPR_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND EP.EPR_FECHA<=" & XDQ(FechaHasta)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_MOVIMIENTO_STOCK (FECHA,NRO_REMITO,NRO_ENTRASALE,"
            sql = sql & "ENTRADA,SALIDA,PENDIENTE,CLIENTE,PRODUCTO,"
            sql = sql & "ORDENAR,OBSERVACION) VALUES ("
            sql = sql & XDQ(rec!EPR_FECHA) & ","
            If Not IsNull(rec!EPR_NROSUCREM) Then
                sql = sql & XS(Format(rec!EPR_NROSUCREM, "0000") & "-" & Format(rec!EPR_NROREM, "00000000")) & ","
            Else
                sql = sql & XS("") & ","
            End If
            sql = sql & XS(Format(rec!EPR_CODIGO, "00000000")) & ","
            sql = sql & XN(rec!DEP_CANTIDAD) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN("0") & ","
            If Not IsNull(rec!REP_CODIGO) Then
                If IsNull(rec!CLI_CODIGO) Then
                    sql = sql & XS(BuscoRepresentada(CStr(rec!REP_CODIGO))) & ","
                Else
                    sql = sql & XS(BuscoCliente(CStr(rec!CLI_CODIGO))) & ","
                End If
            Else
                If IsNull(rec!CLI_CODIGO) Then
                    sql = sql & XS("") & ","
                Else
                    sql = sql & XS(BuscoCliente(CStr(rec!CLI_CODIGO))) & ","
                End If
            End If
            sql = sql & XN(rec!PTO_CODIGO) & ","
            sql = sql & "'E',"
            sql = sql & XS(ChkNull(rec!EPR_OBSERVACIONES)) & ")"
            DBConn.Execute sql
            
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub Salida_Deposito()
    sql = "SELECT EP.EGA_CODIGO, EP.EGA_FECHA, EP.RCL_SUCURSAL, EP.RCL_NUMERO,"
    sql = sql & " EP.EGA_OBSERVACIONES, EP.EST_CODIGO, DRC.PTO_CODIGO, P.PTO_DESCRI,"
    sql = sql & " DRC.DRC_CANTIDAD, C.CLI_CODIGO, C.CLI_RAZSOC, RC.STK_CODIGO"
    sql = sql & " FROM PRODUCTO P, NOTA_PEDIDO NP, CLIENTE C,"
    sql = sql & " ENTREGA_PRODUCTO EP, REMITO_CLIENTE RC, DETALLE_REMITO_CLIENTE DRC"
    sql = sql & " WHERE"
    sql = sql & " NP.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND EP.RCL_NUMERO = RC.RCL_NUMERO"
    sql = sql & " AND EP.RCL_SUCURSAL = RC.RCL_SUCURSAL"
    sql = sql & " AND RC.RCL_SUCURSAL = DRC.RCL_SUCURSAL"
    sql = sql & " AND RC.RCL_NUMERO = DRC.RCL_NUMERO"
    sql = sql & " AND NP.NPE_FECHA = RC.NPE_FECHA "
    sql = sql & " AND NP.NPE_NUMERO = RC.NPE_NUMERO"
    sql = sql & " AND P.PTO_CODIGO = DRC.PTO_CODIGO"
    sql = sql & " AND EP.EST_CODIGO=3"
    If txtcodigo.Text <> "" Then
        sql = sql & " AND DRC.PTO_CODIGO=" & XN(txtcodigo.Text)
    End If
    If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
        sql = sql & " AND RC.STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
    End If
    If FechaDesde <> "" Then sql = sql & " AND EP.EGA_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND EP.EGA_FECHA<=" & XDQ(FechaHasta)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_MOVIMIENTO_STOCK (FECHA,NRO_REMITO,NRO_ENTRASALE,"
            sql = sql & "ENTRADA,SALIDA,PENDIENTE,CLIENTE,PRODUCTO,"
            sql = sql & "ORDENAR,OBSERVACION) VALUES ("
            sql = sql & XDQ(rec!EGA_FECHA) & ","
            sql = sql & XS(Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000")) & ","
            sql = sql & XS(Format(rec!EGA_CODIGO, "00000000")) & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(rec!DRC_CANTIDAD) & ","
            sql = sql & XN("0") & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XN(rec!PTO_CODIGO) & ","
            sql = sql & "'E',"
            sql = sql & XS(ChkNull(rec!EGA_OBSERVACIONES)) & ")"
            DBConn.Execute sql
        
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub Remitos_Pendientes()

    sql = " SELECT RC.RCL_NUMERO, RC.RCL_SUCURSAL, RC.RCL_FECHA,"
    sql = sql & " DR.PTO_CODIGO, DR.DRC_CANTIDAD, C.CLI_RAZSOC"
    sql = sql & " FROM REMITO_CLIENTE RC, DETALLE_REMITO_CLIENTE DR,"
    sql = sql & " NOTA_PEDIDO NP, CLIENTE C"
    sql = sql & " WHERE"
    sql = sql & " RC.NPE_NUMERO=NP.NPE_NUMERO"
    sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
    sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.RCL_SUCURSAL=DR.RCL_SUCURSAL"
    sql = sql & " AND RC.RCL_NUMERO=DR.RCL_NUMERO"
    sql = sql & " AND RC.EST_CODIGO <> 2" 'BUSCA REMITOS NO ANULADOS
    If txtcodigo.Text <> "" Then
        sql = sql & " AND DR.PTO_CODIGO=" & XN(txtcodigo.Text)
    End If
    If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
        sql = sql & " AND RC.STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
    End If
    If chkFechaRemito.Value = Checked Then
        If FechaDesde <> "" Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
    End If
    sql = sql & " AND RC.RCL_NUMERO NOT IN (SELECT RCL_NUMERO FROM ENTREGA_PRODUCTO"
    sql = sql & " WHERE EST_CODIGO=3)"
    sql = sql & " ORDER BY RC.RCL_SUCURSAL,RC.RCL_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_MOVIMIENTO_STOCK (FECHA,NRO_REMITO,NRO_ENTRASALE,"
            sql = sql & "ENTRADA,SALIDA,PENDIENTE,CLIENTE,PRODUCTO,"
            sql = sql & "ORDENAR,OBSERVACION) VALUES ("
            sql = sql & XDQ(rec!RCL_FECHA) & ","
            sql = sql & XS(Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000")) & ","
            sql = sql & "NULL,"
            sql = sql & XN("0") & ","
            sql = sql & XN("0") & ","
            sql = sql & XN(rec!DRC_CANTIDAD) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XN(rec!PTO_CODIGO) & ","
            sql = sql & "'R',"
            sql = sql & XS("REMITO PENDIENTE DE SALIDA") & ")"
            DBConn.Execute sql
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub Pongo_Saldo()
    Dim Saldo As Double
    Dim Orden As Long
    Dim CodProd As Long
    Saldo = 0
    Orden = 1
    CodProd = 0
    
    sql = "SELECT PRODUCTO,FECHA,NRO_REMITO,NRO_ENTRASALE,ENTRADA,SALIDA,"
    sql = sql & " PENDIENTE,SALDO,CLIENTE,OBSERVACION,SECUENCIA,ORDENAR"
    sql = sql & " FROM TMP_MOVIMIENTO_STOCK"
    sql = sql & " ORDER BY PRODUCTO, ORDENAR ASC, FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        CodProd = rec!Producto
        
        Do While rec.EOF = False
            If rec!ENTRADA <> 0 Then
                Saldo = Saldo + CDbl(rec!ENTRADA)
            ElseIf rec!SALIDA <> 0 Then
                Saldo = Saldo - CDbl(rec!SALIDA)
            ElseIf rec!PENDIENTE <> 0 Then
                Saldo = Saldo - CDbl(rec!PENDIENTE)
            End If
            
            sql = "UPDATE TMP_MOVIMIENTO_STOCK"
            sql = sql & " SET SALDO=" & XN(CStr(Saldo))
            sql = sql & " ,SECUENCIA=" & XN(CStr(Orden))
            'ENTRA ACA CUANDO COMIENZA CON EL CIERRE DE STOCK
            If IsNull(rec!NRO_ENTRASALE) And IsNull(rec!NRO_REMITO) Then
                sql = sql & " ,ENTRADA=" & XN("0")
            End If
            sql = sql & " WHERE"
            sql = sql & " FECHA=" & XDQ(rec!Fecha)
            If Not IsNull(rec!NRO_REMITO) Then
                sql = sql & " AND NRO_REMITO=" & XS(ChkNull(rec!NRO_REMITO))
            End If
            If Not IsNull(rec!NRO_ENTRASALE) Then
                sql = sql & " AND NRO_ENTRASALE=" & XS(ChkNull(rec!NRO_ENTRASALE))
            End If
            'ESTO ES PARA CUANDO ES CIERRE DE STOCK
            If IsNull(rec!NRO_ENTRASALE) And IsNull(rec!NRO_REMITO) Then
                sql = sql & " AND ORDENAR='C'"
                sql = sql & " AND FECHA=" & XDQ(rec!Fecha)
            End If
            sql = sql & " AND PRODUCTO=" & XN(rec!Producto)
            DBConn.Execute sql
            
            Orden = Orden + 1
            rec.MoveNext
            If rec.EOF = False Then
                If CodProd <> rec!Producto Then
                    Orden = 1
                    Saldo = 0
                    CodProd = rec!Producto
                End If
            End If
            
        Loop
    End If
    rec.Close
End Sub

Private Function BuscoRepresentada(Codigo As String) As String
    sql = "SELECT REP_RAZSOC FROM REPRESENTADA"
    sql = sql & " WHERE REP_CODIGO=" & XN(Codigo)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscoRepresentada = Rec1!REP_RAZSOC
    Else
        BuscoRepresentada = ""
    End If
    Rec1.Close
End Function

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoCliente = Rec1!CLI_RAZSOC
        Else
            BuscoCliente = ""
        End If
        Rec1.Close
End Function

Private Sub cmdListar_Click()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    If FechaDesde.Text = "" Then
        MsgBox "Debe ingresar una fecha de inicio", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Claveti
    lblEstado.Caption = "Buscando Movimiento..."
    Screen.MousePointer = vbHourglass
    
    sql = "DELETE FROM TMP_MOVIMIENTO_STOCK"
    DBConn.Execute sql
    Cierre_Stock
    Entrada_Deposito
    Salida_Deposito
    Remitos_Pendientes
    Pongo_Saldo
    'BORRO ESTA TABLA PORQUE SINO MUESTRA LOS DATOS EN EL LISTADO COMO QUIERE
    sql = "DELETE FROM TMP_LISTADO_DETALLE_STOCK"
    DBConn.Execute sql
    
    LlamoReporte
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    Exit Sub
    
Claveti:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub LlamoReporte()
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SISESTILO"
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Periodo  Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Periodo  Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Periodo  Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Periodo  Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
        Rep.Formulas(1) = "STOCK='" & "               Stock: " & Trim(cboStock.List(cboStock.ListIndex)) & "'"
    Rep.WindowTitle = "Listado de Movimeinto de Stock"
    
    Rep.ReportFileName = DRIVE & DirReport & "movimientostock.rpt"
    
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
     Rep.Action = 1
     
     lblEstado.Caption = ""
     Rep.Formulas(0) = ""
     Rep.Formulas(1) = ""
End Sub

Private Sub CmdNuevo_Click()
    txtcodigo.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    chkFechaRemito.Value = Unchecked
    cboStock.ListIndex = 0
    txtcodigo.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmMovimientoStock = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    CargocboStock
    
    lblEstado.Caption = ""
End Sub

Private Sub TxtCodigo_Change()
    If txtcodigo.Text = "" Then
        txtcodigo.Text = ""
        txtdescri.Text = ""
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, "
        sql = sql & " R.RUB_DESCRI,RE.REP_RAZSOC,P.PTO_CODIGO"
        sql = sql & " FROM PRODUCTO P,LINEAS L,RUBROS R,REPRESENTADA RE"
        sql = sql & " WHERE P.LNA_CODIGO = L.LNA_CODIGO"
        sql = sql & " AND P.RUB_CODIGO = R.RUB_CODIGO"
        sql = sql & " AND P.REP_CODIGO = RE.REP_CODIGO"
        sql = sql & " AND P.PTO_CODIGO = " & XN(txtcodigo.Text)
        sql = sql & " ORDER BY P.PTO_CODIGO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtdescri.Text = rec!PTO_DESCRI
        Else
            MsgBox "El Código no existe, o no pertenece al stock de " & cboStock.Text & "", vbExclamation, TIT_MSGBOX
            txtcodigo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub CargocboStock()
    sql = "SELECT S.STK_CODIGO,R.REP_RAZSOC FROM STOCK S, REPRESENTADA R "
    sql = sql & " ORDER BY S.STK_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboStock.AddItem "(Todos)"
        Do While rec.EOF = False
            cboStock.AddItem rec!REP_RAZSOC
            cboStock.ItemData(cboStock.NewIndex) = rec!STK_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtDescri_Change()
    If txtdescri.Text = "" Then
        txtcodigo.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtdescri
End Sub

Private Sub txtdescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
           
   If txtcodigo.Text = "" And txtdescri.Text <> "" Then
        Set rec = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, R.RUB_DESCRI, L.LNA_DESCRI,RE.REP_RAZSOC"
        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L,REPRESENTADA RE"
        sql = sql & " WHERE P.RUB_CODIGO = R.RUB_CODIGO"
        sql = sql & " AND P.LNA_CODIGO = L.LNA_CODIGO AND L.LNA_CODIGO = R.LNA_CODIGO"
        sql = sql & " AND RE.REP_CODIGO=P.REP_CODIGO"
        sql = sql & " AND P.PTO_DESCRI LIKE '" & txtdescri.Text & "%'ORDER BY P.PTO_DESCRI"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                'grdGrilla.SetFocus
                frmBuscar.TipoBusqueda = 2
                frmBuscar.CodListaPrecio = 0
                frmBuscar.TxtDescriB.Text = txtdescri.Text
                frmBuscar.Show vbModal
                frmBuscar.grdBuscar.Col = 0
                txtcodigo.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                frmBuscar.grdBuscar.Col = 1
                txtdescri.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
            Else
                txtcodigo.Text = Trim(rec!PTO_CODIGO)
                txtdescri.Text = Trim(rec!PTO_DESCRI)
            End If
        Else
                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                txtdescri.Text = ""
        End If
        rec.Close
        Screen.MousePointer = vbNormal
    End If
    
End Sub

