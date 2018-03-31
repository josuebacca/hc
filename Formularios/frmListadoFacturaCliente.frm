VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmListadoFacturaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Factura de Cliente"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoFacturaCliente.frx":0000
      Height          =   750
      Left            =   8865
      Picture         =   "frmListadoFacturaCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6495
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   9270
      Top             =   6045
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   9930
      Top             =   6090
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9750
      Picture         =   "frmListadoFacturaCliente.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6495
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7995
      Picture         =   "frmListadoFacturaCliente.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6495
      Width           =   855
   End
   Begin VB.Frame Frame1 
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
      Height          =   600
      Left            =   90
      TabIndex        =   28
      Top             =   5415
      Width           =   10425
      Begin VB.OptionButton optDetalladoTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado "
         Height          =   255
         Left            =   3945
         TabIndex        =   9
         Top             =   240
         Width           =   1770
      End
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado Seleccionado"
         Height          =   255
         Left            =   6720
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   1140
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1650
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   90
      TabIndex        =   24
      Top             =   6030
      Width           =   7845
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   12
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblImpresora 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1965
         TabIndex        =   26
         Top             =   840
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Facturas de Cliente por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   105
      TabIndex        =   16
      Top             =   75
      Width           =   10395
      Begin VB.TextBox txtOrden 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   33
         Text            =   "A"
         Top             =   1575
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cboRep 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1695
         Width           =   3630
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1020
         Width           =   3630
      End
      Begin VB.CommandButton cmdBuscarVendedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoFacturaCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Buscar Vendedor"
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.TextBox txtCliente 
         Height          =   330
         Left            =   2520
         MaxLength       =   40
         TabIndex        =   0
         Top             =   300
         Width           =   1005
      End
      Begin VB.TextBox txtDesCli 
         BackColor       =   &H8000000F&
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
         Height          =   330
         Left            =   4005
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "Descripción"
         Top             =   300
         Width           =   5175
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   360
         Left            =   7515
         MaskColor       =   &H000000FF&
         TabIndex        =   6
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   1650
         UseMaskColor    =   -1  'True
         Width           =   1665
      End
      Begin VB.TextBox txtDesVen 
         BackColor       =   &H8000000F&
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
         Height          =   330
         Left            =   4005
         TabIndex        =   18
         Top             =   660
         Width           =   5175
      End
      Begin VB.TextBox txtVendedor 
         Height          =   330
         Left            =   2520
         TabIndex        =   1
         Top             =   660
         Width           =   1005
      End
      Begin VB.CommandButton cmdBuscarCli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoFacturaCliente.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Buscar"
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   5025
         TabIndex        =   4
         Top             =   1380
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   2520
         TabIndex        =   3
         Top             =   1380
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Representada:"
         Height          =   195
         Left            =   1410
         TabIndex        =   32
         Top             =   1755
         Width           =   1080
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   1410
         TabIndex        =   31
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   3
         Left            =   1410
         TabIndex        =   23
         Top             =   345
         Width           =   570
      End
      Begin VB.Label lblFechaDesde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1410
         TabIndex        =   22
         Top             =   1410
         Width           =   990
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3975
         TabIndex        =   21
         Top             =   1425
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   1410
         TabIndex        =   20
         Top             =   690
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3105
      Left            =   90
      TabIndex        =   7
      Top             =   2220
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   5477
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   8010
      TabIndex        =   29
      Top             =   6180
      Width           =   660
   End
End
Attribute VB_Name = "frmListadoFacturaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub Insertar_Temporal()
    
    sql = "DELETE FROM TMP_FACTURA_CLIENTE"
    DBConn.Execute sql
    
    'AGREGO LAS FACTURAS POR REMITO
    sql = "INSERT INTO TMP_FACTURA_CLIENTE"
    sql = sql & " SELECT FC.TCO_CODIGO,FC.FCL_NUMERO, FC.FCL_SUCURSAL,FC.REP_CODIGO,"
    sql = sql & " FC.FCL_FECHA,NP.CLI_CODIGO, NP.VEN_CODIGO"
    sql = sql & " FROM FACTURA_CLIENTE FC,REMITO_CLIENTE RC, NOTA_PEDIDO NP"
    sql = sql & " WHERE"
    sql = sql & " FC.RCL_NUMERO=RC.RCL_NUMERO"
    sql = sql & " AND FC.RCL_SUCURSAL=RC.RCL_SUCURSAL"
    sql = sql & " AND RC.NPE_NUMERO=NP.NPE_NUMERO"
    sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
    sql = sql & " AND FC.FCL_TIPO='R'"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente.Text)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor.Text)
    If FechaDesde <> "" Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde.Text)
    If FechaHasta <> "" Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta.Text)
    If cboEstado.List(cboEstado.ListIndex) <> "(Todos)" Then sql = sql & " AND FC.EST_CODIGO=" & XN(cboEstado.ItemData(cboEstado.ListIndex))
    If cboRep.List(cboRep.ListIndex) <> "(Todas)" Then sql = sql & " AND FC.REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    DBConn.Execute sql
    
    'AGREGO LAS FACTURAS POR CONCEPTO Y PRODUCTO
    sql = "INSERT INTO TMP_FACTURA_CLIENTE"
    sql = sql & " SELECT FC.TCO_CODIGO,FC.FCL_NUMERO, FC.FCL_SUCURSAL,FC.REP_CODIGO,"
    sql = sql & " FC.FCL_FECHA,FC.CLI_CODIGO, FC.VEN_CODIGO"
    sql = sql & " FROM FACTURA_CLIENTE FC"
    sql = sql & " WHERE"
    sql = sql & " FC.FCL_TIPO<>'R'"
    If txtCliente.Text <> "" Then sql = sql & " AND FC.CLI_CODIGO=" & XN(txtCliente.Text)
    If txtVendedor.Text <> "" Then sql = sql & " AND FC.VEN_CODIGO=" & XN(txtVendedor.Text)
    If FechaDesde <> "" Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde.Text)
    If FechaHasta <> "" Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta.Text)
    If cboEstado.List(cboEstado.ListIndex) <> "(Todos)" Then sql = sql & " AND FC.EST_CODIGO=" & XN(cboEstado.ItemData(cboEstado.ListIndex))
    If cboRep.List(cboRep.ListIndex) <> "(Todas)" Then sql = sql & " AND FC.REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    DBConn.Execute sql
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    
    Insertar_Temporal
    
    sql = "SELECT T.FCL_NUMERO, T.FCL_SUCURSAL, T.FCL_FECHA, C.CLI_RAZSOC,"
    sql = sql & " R.REP_RAZSOC, T.TCO_CODIGO, TC.TCO_ABREVIA"
    sql = sql & " FROM TMP_FACTURA_CLIENTE T, CLIENTE C, REPRESENTADA R, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " T.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND T.REP_CODIGO=R.REP_CODIGO"
    sql = sql & " AND T.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " ORDER BY T.FCL_FECHA"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        GrdModulos.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            GrdModulos.AddItem Trim(rec!TCO_ABREVIA) & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                            & Chr(9) & rec!FCL_FECHA & Chr(9) & rec!CLI_RAZSOC _
                            & Chr(9) & rec!REP_RAZSOC & Chr(9) & rec!TCO_CODIGO
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
        GrdModulos.Col = 0
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
    End If
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdBuscarVendedor_Click()
    frmBuscar.TipoBusqueda = 4
    frmBuscar.txtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtVendedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtDesVen.Text = frmBuscar.grdBuscar.Text
        txtVendedor.SetFocus
    Else
        txtVendedor.SetFocus
    End If
End Sub

Private Sub cmdListar_Click()
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Buscando Listado..."
    
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SISESTILO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    
    If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
        Rep.Formulas(1) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
        Rep.Formulas(1) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
        Rep.Formulas(1) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
        Rep.Formulas(1) = "FECHA='" & "Al: " & Date & "'"
    End If
    If optGeneralTodos.Value = True Or optDetalladoTodos.Value = True Then
        Rep.SelectionFormula = ""
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{TMP_FACTURA_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
        Else
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
        End If
        If FechaDesde.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            End If
        End If
        
        If FechaHasta.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            End If
        End If
        
        If cboEstado.List(cboEstado.ListIndex) <> "(Todos)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=" & XN(cboEstado.ItemData(cboEstado.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.EST_CODIGO}=" & XN(cboEstado.ItemData(cboEstado.ListIndex))
            End If
        End If
        
        If cboRep.List(cboRep.ListIndex) <> "(Todas)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {FACTURA_CLIENTE.REP_CODIGO}=" & XN(cboRep.ItemData(cboRep.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.REP_CODIGO}=" & XN(cboRep.ItemData(cboRep.ListIndex))
            End If
        End If
        If optDetalladoTodos.Value = True Then
            Rep.Formulas(0) = ""
            Rep.WindowTitle = "Factura Cliente - Detallado..."
            Rep.ReportFileName = DRIVE & DirReport & "facturaclientedetalle.rpt"
        Else
            Rep.WindowTitle = "Factura Cliente - General..."
            Rep.ReportFileName = DRIVE & DirReport & "facturaclientegeneral.rpt"
        End If
    End If
    
    If optDetallado.Value = True Then
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
            MsgBox "Debe seleccionar una Factura", vbExclamation, TIT_MSGBOX
            GrdModulos.SetFocus
            Exit Sub
        End If
        Rep.SelectionFormula = ""
        Rep.SelectionFormula = "{FACTURA_CLIENTE.FCL_NUMERO}=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)) _
                               & " AND {FACTURA_CLIENTE.FCL_SUCURSAL}=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)) _
                               & " AND {FACTURA_CLIENTE.TCO_CODIGO}=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 5))
                               
                               
        Rep.WindowTitle = "Factura Cliente - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "facturaclientedetalle.rpt"
    End If
    
    If optPantalla.Value = True Then
        Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
    End If
    Rep.Action = 1
     
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.txtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub

Private Sub CmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtVendedor.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    cboEstado.ListIndex = 0
    cboRep.ListIndex = 0
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    optGeneralTodos.Value = True
    optPantalla.Value = True
    txtCliente.SetFocus
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    GrdModulos.Rows = 1
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""

    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = "^Comp.|^Número|^Fecha|Cliente|Representada|TIPO DE FACTURA"
    GrdModulos.ColWidth(0) = 800
    GrdModulos.ColWidth(1) = 1300
    GrdModulos.ColWidth(2) = 1100
    GrdModulos.ColWidth(3) = 3800
    GrdModulos.ColWidth(4) = 3000
    GrdModulos.ColWidth(5) = 0
    GrdModulos.Rows = 1
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.row = 0
    Dim I As Integer
    For I = 0 To 4
        GrdModulos.Col = I
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080 'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    '------------------------------------
    '------------------------------------
    'CARGO COMBO ESTADO
    CargoComboEstado
    
    'CRAGO COMBO REPRESENTADA
    cboRep.AddItem "(Todas)"
    Call CargoComboBox(cboRep, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboRep.ListIndex = 0
End Sub

Private Sub CargoComboEstado()
    sql = "SELECT * FROM ESTADO_DOCUMENTO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboEstado.AddItem "(Todos)"
        Do While rec.EOF = False
            cboEstado.AddItem rec!EST_DESCRI
            cboEstado.ItemData(cboEstado.NewIndex) = rec!EST_CODIGO
            rec.MoveNext
        Loop
        cboEstado.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub GrdModulos_Click()
    If GrdModulos.MouseRow = 0 Then
        GrdModulos.Col = GrdModulos.MouseCol
        GrdModulos.ColSel = GrdModulos.MouseCol
        
        If txtOrden.Text = "A" Then
            GrdModulos.Sort = 2
            txtOrden.Text = "B"
        Else
            GrdModulos.Sort = 1
            txtOrden.Text = "A"
        End If
    End If
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoFacturaCliente = Nothing
    Unload Me
End Sub

Private Sub txtVendedor_Change()
    If txtVendedor.Text = "" Then
        txtDesVen.Text = ""
    End If
End Sub

Private Sub txtVendedor_GotFocus()
    SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtVendedor_LostFocus()
    If txtVendedor.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT VEN_NOMBRE"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE VEN_CODIGO=" & XN(txtVendedor)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtDesVen.Text = Trim(rec!VEN_NOMBRE)
        Else
            MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
            txtDesVen.Text = ""
            txtVendedor.SetFocus
        End If
        rec.Close
    End If
End Sub

