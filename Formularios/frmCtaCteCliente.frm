VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmCtaCteCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cta-Cte Clientes"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7995
   Begin VB.Frame Frame1 
      Caption         =   "Movimientos.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   4965
      TabIndex        =   18
      Top             =   1245
      Width           =   2970
      Begin VB.OptionButton optSaldosHistoricos 
         Caption         =   "Saldos Historicos"
         Height          =   225
         Left            =   1365
         TabIndex        =   6
         Top             =   315
         Width           =   1545
      End
      Begin VB.OptionButton optSaldos 
         Caption         =   "Saldos"
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton optPendiente 
         Caption         =   "Pendientes"
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   660
         Width           =   1155
      End
      Begin VB.OptionButton optTodo 
         Caption         =   "Todos"
         Height          =   195
         Left            =   1365
         TabIndex        =   7
         Top             =   660
         Width           =   1500
      End
   End
   Begin VB.Frame FrameImpresora 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   45
      TabIndex        =   15
      Top             =   1245
      Width           =   4920
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   345
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   1755
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   1020
         TabIndex        =   11
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2085
         TabIndex        =   12
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   405
      Left            =   4980
      TabIndex        =   8
      Top             =   2325
      Width           =   960
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6945
      TabIndex        =   10
      Top             =   2325
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   405
      Left            =   5955
      TabIndex        =   9
      Top             =   2325
      Width           =   975
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   45
      TabIndex        =   13
      Top             =   0
      Width           =   7890
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   0
         Top             =   345
         Width           =   720
      End
      Begin VB.TextBox txtDesCli 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   345
         Width           =   4575
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   4095
         TabIndex        =   3
         Top             =   690
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   1620
         TabIndex        =   2
         Top             =   690
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3045
         TabIndex        =   20
         Top             =   735
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   525
         TabIndex        =   19
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   525
         TabIndex        =   14
         Top             =   390
         Width           =   555
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3885
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   3390
      Top             =   2145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   21
      Top             =   2355
      Width           =   660
   End
End
Attribute VB_Name = "frmCtaCteCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Saldo As Double
Dim Cliente As Integer
Dim Orden As Integer

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub BuscarCtaCTeClientes()
       
    SQL = "DELETE FROM CTA_CTE_CLIENTE"
    DBConn.Execute SQL
    
    If optPendiente.Value = True Or optSaldos.Value = True Then
        
        'FACTURAS PENDIENTES
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,COM_NUMEROTXT)"
        SQL = SQL & " SELECT F.CLI_CODIGO,F.TCO_CODIGO,F.FCL_NUMERO,F.FCL_SUCURSAL,"
        SQL = SQL & " F.FCL_FECHA,F.FCL_TOTAL,F.FCL_SALDO,0 AS HABER,'D' AS DEBE,FCL_NUMEROTXT"
        SQL = SQL & " FROM SALDO_FACTURAS_CLIENTE_V F"
        SQL = SQL & " WHERE"
        SQL = SQL & " F.EST_CODIGO=3"
        SQL = SQL & " AND F.FCL_SALDO > 0"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND F.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND F.FCL_FECHA>=" & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND F.FCL_FECHA<=" & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL

        'NOTA DEBITOS CLIENTE PENDIENTES
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        SQL = SQL & " COM_NUMEROTXT)"
        SQL = SQL & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NDC_NUMERO,N.NDC_SUCURSAL,"
        SQL = SQL & " N.NDC_FECHA,N.NDC_TOTAL,N.NDC_SALDO,0 AS HABER,'D' AS DEBE,N.NDC_NUMEROTXT"
        SQL = SQL & " FROM NOTA_DEBITO_CLIENTE N"
        SQL = SQL & " WHERE N.EST_CODIGO=3"
        SQL = SQL & " AND N.NDC_SALDO > 0"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND N.NDC_FECHA>=" & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND N.NDC_FECHA<=" & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
        
        'NOTA CREDITO CLIENTE PENDIENTES
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        SQL = SQL & " COM_NUMEROTXT)"
        SQL = SQL & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NCC_NUMERO,N.NCC_SUCURSAL,"
        SQL = SQL & " N.NCC_FECHA,N.NCC_TOTAL,0 AS DEBE,NCC_SALDO,'C' AS CREDITO,N.NCC_NUMEROTXT"
        SQL = SQL & " FROM NOTA_CREDITO_CLIENTE N"
        SQL = SQL & " WHERE N.EST_CODIGO=3"
        SQL = SQL & " AND N.NCC_SALDO > 0"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND N.NCC_FECHA>=" & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND N.NCC_FECHA<=" & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
        
        'TODOS LOS RECIBOS CON SALDOS
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        SQL = SQL & " COM_NUMEROTXT)"
        SQL = SQL & " SELECT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,R.REC_FECHA,"
        SQL = SQL & " S.REC_SALDO AS TOTAL,0 AS DEBE,S.REC_SALDO AS HABER,'C' AS CREDITO,R.REC_NUMEROTXT"
        SQL = SQL & " FROM RECIBO_CLIENTE R , RECIBO_CLIENTE_SALDO S"
        SQL = SQL & " WHERE R.EST_CODIGO=3"
        SQL = SQL & " AND R.TCO_CODIGO=S.TCO_CODIGO"
        SQL = SQL & " AND R.REC_SUCURSAL=S.REC_SUCURSAL"
        SQL = SQL & " AND R.REC_NUMERO=S.REC_NUMERO"
        SQL = SQL & " AND S.REC_SALDO > 0"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND R.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND R.REC_FECHA >= " & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND R.REC_FECHA <= " & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
    End If
    
    If optTodo.Value = True Or optSaldosHistoricos.Value = True Then
        'TODAS LAS FACTURAS
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,COM_NUMEROTXT)"
        SQL = SQL & " SELECT F.CLI_CODIGO,F.TCO_CODIGO,F.FCL_NUMERO,F.FCL_SUCURSAL,"
        SQL = SQL & " F.FCL_FECHA,F.FCL_TOTAL,F.FCL_TOTAL,0 AS HABER,'D' AS DEBE,FCL_NUMEROTXT"
        SQL = SQL & " FROM FACTURA_CLIENTE F"
        SQL = SQL & " WHERE F.EST_CODIGO=3"
        SQL = SQL & " AND FPG_CODIGO=2"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND F.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND F.FCL_FECHA >= " & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND F.FCL_FECHA <= " & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
    
        'TODAS LAS NOTAS DEBITOS CLIENTE
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        SQL = SQL & " COM_NUMEROTXT)"
        SQL = SQL & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NDC_NUMERO,N.NDC_SUCURSAL,"
        SQL = SQL & " N.NDC_FECHA,N.NDC_TOTAL,N.NDC_TOTAL,0 AS HABER,'D' AS DEBE,N.NDC_NUMEROTXT"
        SQL = SQL & " FROM NOTA_DEBITO_CLIENTE N"
        SQL = SQL & " WHERE N.EST_CODIGO=3"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND N.NDC_FECHA >= " & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND N.NDC_FECHA <= " & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
        
        'TODAS LAS NOTAS CREDITO CLIENTE
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        SQL = SQL & " COM_NUMEROTXT)"
        SQL = SQL & " SELECT DISTINCT N.CLI_CODIGO,N.TCO_CODIGO,N.NCC_NUMERO,N.NCC_SUCURSAL,"
        SQL = SQL & " N.NCC_FECHA,N.NCC_TOTAL,0 AS DEBE,NCC_TOTAL,'C' AS CREDITO,N.NCC_NUMEROTXT"
        SQL = SQL & " FROM NOTA_CREDITO_CLIENTE N"
        SQL = SQL & " WHERE N.EST_CODIGO=3"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND N.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND N.NCC_FECHA >= " & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND N.NCC_FECHA <= " & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
        
        'TODOS LOS RECIBOS
        SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
        SQL = SQL & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
        SQL = SQL & " COM_NUMEROTXT)"
        SQL = SQL & " SELECT DISTINCT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,"
        SQL = SQL & " R.REC_FECHA,R.REC_TOTAL,0 AS DEBE,REC_TOTAL,'C' AS CREDITO,R.REC_NUMEROTXT"
        SQL = SQL & " FROM RECIBO_CLIENTE R"
        SQL = SQL & " WHERE R.EST_CODIGO=3"
        If txtCliente.Text <> "" Then
            SQL = SQL & " AND R.CLI_CODIGO=" & XN(txtCliente.Text)
        End If
        If FechaDesde.Text <> "" Then
            SQL = SQL & " AND R.REC_FECHA >= " & XDQ(FechaDesde.Text)
        End If
        If FechaHasta.Text <> "" Then
            SQL = SQL & " AND R.REC_FECHA <= " & XDQ(FechaHasta.Text)
        End If
        DBConn.Execute SQL
        
        'TODOS LOS RECIBOS CON SALDOS
'        sql = " SELECT R.CLI_CODIGO,R.TCO_CODIGO,R.REC_NUMERO,R.REC_SUCURSAL,"
'        sql = sql & " R.REC_FECHA,(R.REC_TOTAL+S.REC_SALDO) AS TOTAL,R.REC_NUMEROTXT"
'        sql = sql & " FROM RECIBO_CLIENTE R , RECIBO_CLIENTE_SALDO S"
'        sql = sql & " WHERE R.EST_CODIGO=3"
'        sql = sql & " AND R.TCO_CODIGO=S.TCO_CODIGO"
'        sql = sql & " AND R.REC_SUCURSAL=S.REC_SUCURSAL"
'        sql = sql & " AND R.REC_NUMERO=S.REC_NUMERO"
'        If txtCliente.Text <> "" Then
'            sql = sql & " AND R.CLI_CODIGO=" & XN(txtCliente.Text)
'        End If
'        If FechaDesde.Text <> "" Then
'            sql = sql & " AND R.REC_FECHA >= " & XDQ(FechaDesde.Text)
'        End If
'        If FechaHasta.Text <> "" Then
'            sql = sql & " AND R.REC_FECHA <= " & XDQ(FechaHasta.Text)
'        End If
'        Rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec.EOF = False Then
'            Do While Rec.EOF = False
'                sql = "DELETE FROM CTA_CTE_CLIENTE"
'                sql = sql & " WHERE"
'                sql = sql & " CLI_CODIGO=" & XN(Rec!CLI_CODIGO)
'                sql = sql & " AND TCO_CODIGO=" & XN(Rec!TCO_CODIGO)
'                sql = sql & " AND COM_NUMERO=" & XN(Rec!REC_NUMERO)
'                sql = sql & " AND COM_SUCURSAL=" & XN(Rec!REC_SUCURSAL)
'                DBConn.Execute sql
'
'                sql = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,COM_SUCURSAL,"
'                sql = sql & " COM_FECHA,COM_IMPORTE,COM_IMP_DEBE,COM_IMP_HABER,CTA_CTE_DH,"
'                sql = sql & " COM_NUMEROTXT)"
'                sql = sql & " VALUES ("
'                sql = sql & XN(Rec!CLI_CODIGO) & ","
'                sql = sql & XN(Rec!TCO_CODIGO) & ","
'                sql = sql & XN(Rec!REC_NUMERO) & ","
'                sql = sql & XN(Rec!REC_SUCURSAL) & ","
'                sql = sql & XDQ(Rec!REC_FECHA) & ","
'                sql = sql & XN(Rec!TOTAL) & ","
'                sql = sql & XN("0") & ","
'                sql = sql & XN(Rec!TOTAL) & ","
'                sql = sql & XS("C") & ","
'                sql = sql & XS(Rec!REC_NUMEROTXT) & ")"
'                DBConn.Execute sql
                
'                Rec.MoveNext
'            Loop
'        End If
'        Rec.Close
    End If
    If optSaldos.Value = True Or optSaldosHistoricos.Value = True Then
        BuscaSaldosGeneral
    Else
        BuscaSaldosDetalle
    End If
End Sub

Private Sub BuscaSaldosDetalle()
    'CONFIGURO EL SALDO
    SQL = "SELECT * FROM CTA_CTE_CLIENTE"
    SQL = SQL & " ORDER BY CLI_CODIGO,COM_FECHA"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Cliente = rec!CLI_CODIGO
        Saldo = 0
        Orden = 1
        Do While rec.EOF = False
            If rec!CTA_CTE_DH = "D" Then
                Saldo = Saldo + CDbl(Chk0(rec!COM_IMP_DEBE))
            Else
                Saldo = Saldo - CDbl(Chk0(rec!COM_IMP_HABER))
            End If
            SQL = "UPDATE CTA_CTE_CLIENTE SET CTA_CTA_SALDO=" & XN(CStr(Saldo))
            SQL = SQL & " ,CTA_CTE_ORDEN=" & XN(CStr(Orden))
            SQL = SQL & " ,COM_NUMEROTXT=" & XS(Format(rec!COM_NUMERO, "00000000"))
            SQL = SQL & " WHERE CLI_CODIGO=" & XN(rec!CLI_CODIGO)
            SQL = SQL & " AND TCO_CODIGO=" & XN(rec!TCO_CODIGO)
            SQL = SQL & " AND COM_NUMERO=" & XN(rec!COM_NUMERO)
            SQL = SQL & " AND COM_SUCURSAL=" & XN(rec!COM_SUCURSAL)
            DBConn.Execute SQL
            
            Orden = Orden + 1
            rec.MoveNext
            If rec.EOF = False Then
                'SI NO VA DETALLADO POR REPRESENTADA
                If Cliente <> rec!CLI_CODIGO Then
                    Cliente = rec!CLI_CODIGO
                    Saldo = 0
                    Orden = 1
                End If
            End If
        Loop
    End If
    rec.Close
End Sub

Private Sub BuscaSaldosGeneral()
    'CONFIGURO EL SALDO
    SQL = "SELECT SUM(COM_IMP_DEBE) AS DEBE ,SUM(COM_IMP_HABER)AS HABER "
    SQL = SQL & " ,CLI_CODIGO"
    SQL = SQL & " FROM CTA_CTE_CLIENTE"
    SQL = SQL & " GROUP BY CLI_CODIGO"
    SQL = SQL & " ORDER BY CLI_CODIGO"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Saldo = 0
        Do While rec.EOF = False
             Saldo = CDbl(rec!DEBE) - CDbl(rec!HABER)
             SQL = "DELETE FROM CTA_CTE_CLIENTE"
             SQL = SQL & " WHERE CLI_CODIGO=" & XN(rec!CLI_CODIGO)
             DBConn.Execute SQL
             
            'If Saldo > 0 Then
                SQL = "INSERT INTO CTA_CTE_CLIENTE (CLI_CODIGO,TCO_CODIGO,COM_NUMERO,"
                SQL = SQL & "COM_SUCURSAL,CTA_CTE_SALDOFINAL)"
                SQL = SQL & " VALUES ("
                SQL = SQL & XN(rec!CLI_CODIGO) & ","
                SQL = SQL & XN("1") & ","
                SQL = SQL & XN("1") & ","
                SQL = SQL & XN("1") & ","
                SQL = SQL & XN(CStr(Saldo)) & ")"
                DBConn.Execute SQL
            'End If
            Saldo = 0
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub cmdListar_Click()
    On Error GoTo CLAVOSE
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Buscando..."
    'LLENO LA TABLA CTA_CTE_CLIENTE
    BuscarCtaCTeClientes
    
    DBConn.Execute "DELETE FROM TMP_RESUMEN_CUENTA_BANCO"
        
    Rep.WindowState = crptMaximized
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""

    If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & " Al: " & Date & "'"
    End If
    
    Rep.WindowTitle = "CTA-CTE de Clientes..."
    If optPendiente.Value = True Or optTodo.Value = True Then
        Rep.ReportFileName = DRIVE & DirReport & "ctacte_clientes.rpt"
    Else
        Rep.ReportFileName = DRIVE & DirReport & "ctacte_clientes_Saldos.rpt"
    End If
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
    Rep.Action = 1
     
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdNuevo_Click()
    txtCliente.Text = ""
    lblEstado.Caption = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    optSaldos.Value = True
End Sub

Private Sub cmdSalir_Click()
    Set frmCtaCteCliente = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset

    Me.Left = 0
    Me.Top = 0
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "", "CODIGO"
    End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = "SELECT CLI_CODIGO, CLI_RAZSOC"
        SQL = SQL & " FROM CLIENTE"
        SQL = SQL & " WHERE "
        If txtCliente.Text <> "" Then
            SQL = SQL & " CLI_CODIGO=" & XN(txtCliente)
        End If
        If rec.State = 1 Then rec.Close
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = Trim(rec!CLI_RAZSOC)
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "", "CODIGO"
    End If
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = "SELECT CLI_CODIGO, CLI_RAZSOC"
        SQL = SQL & " FROM CLIENTE"
        SQL = SQL & " WHERE "
        SQL = SQL & " CLI_RAZSOC LIKE '" & Trim(txtDesCli) & "%'"
        If rec.State = 1 Then rec.Close
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "", "CADENA", Trim(txtDesCli.Text)
                If rec.State = 1 Then rec.Close
                txtDesCli.SetFocus
            Else
                txtCliente.Text = rec!CLI_CODIGO
                txtDesCli.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Function BuscoCliente(Cli As String) As String
    SQL = "SELECT CLI_CODIGO, CLI_RAZSOC"
    SQL = SQL & " FROM CLIENTE"
    SQL = SQL & " WHERE "
    If txtCliente.Text <> "" Then
        SQL = SQL & " CLI_CODIGO=" & XN(Cli)
    Else
        SQL = SQL & " CLI_RAZSOC LIKE '" & Cli & "%'"
    End If
    BuscoCliente = SQL
End Function

Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCliente.Text = .ResultFields(2)
            txtCliente_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub
