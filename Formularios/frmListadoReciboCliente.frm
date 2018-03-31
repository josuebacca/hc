VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmListadoReciboCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Recibo de Cliente"
   ClientHeight    =   6315
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
   ScaleHeight     =   6315
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoReciboCliente.frx":0000
      Height          =   750
      Left            =   8790
      Picture         =   "frmListadoReciboCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5535
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6615
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7290
      Top             =   4935
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9690
      Picture         =   "frmListadoReciboCliente.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5535
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7905
      Picture         =   "frmListadoReciboCliente.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5535
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
      TabIndex        =   27
      Top             =   4440
      Width           =   10425
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado"
         Height          =   255
         Left            =   5475
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1665
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
      TabIndex        =   23
      Top             =   5055
      Width           =   7665
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   11
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   840
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Recibo de Cliente por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   15
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
         Left            =   6870
         TabIndex        =   31
         Text            =   "A"
         Top             =   1005
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cboBuscaRep 
         Height          =   315
         Left            =   2655
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1605
         Width           =   3090
      End
      Begin VB.ComboBox cboRecibo 
         Height          =   315
         Left            =   2655
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   930
         Width           =   2400
      End
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   2655
         MaxLength       =   40
         TabIndex        =   0
         Top             =   255
         Width           =   990
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
         Height          =   300
         Left            =   4140
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "Descripción"
         Top             =   255
         Width           =   4620
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   360
         Left            =   7095
         MaskColor       =   &H000000FF&
         TabIndex        =   6
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   1560
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
         Height          =   285
         Left            =   3705
         TabIndex        =   17
         Top             =   615
         Width           =   5055
      End
      Begin VB.TextBox txtVendedor 
         Height          =   300
         Left            =   2655
         TabIndex        =   1
         Top             =   600
         Width           =   990
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
         Height          =   315
         Left            =   3705
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoReciboCliente.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Buscar"
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   5160
         TabIndex        =   4
         Top             =   1290
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   2655
         TabIndex        =   3
         Top             =   1290
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Representada:"
         Height          =   195
         Left            =   1515
         TabIndex        =   30
         Top             =   1635
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   2205
         TabIndex        =   29
         Top             =   960
         Width           =   360
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
         Left            =   2040
         TabIndex        =   22
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1560
         TabIndex        =   21
         Top             =   1335
         Width           =   990
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   4110
         TabIndex        =   20
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   1830
         TabIndex        =   19
         Top             =   645
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   2295
      Left            =   90
      TabIndex        =   7
      Top             =   2115
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4048
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
      Left            =   7920
      TabIndex        =   28
      Top             =   5205
      Width           =   660
   End
End
Attribute VB_Name = "frmListadoReciboCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub CmdBuscAprox_Click()
    Dim Representada As String
    Representada = ""
    
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT RC.REC_NUMERO, RC.REC_SUCURSAL, RC.REC_FECHA, RC.REP_CODIGO,"
    sql = sql & " RC.TCO_CODIGO, TC.TCO_ABREVIA,"
    sql = sql & " C.CLI_RAZSOC, V.VEN_NOMBRE, R.REP_RAZSOC"
    sql = sql & " FROM RECIBO_CLIENTE RC, CLIENTE C, VENDEDOR V, TIPO_COMPROBANTE TC, REPRESENTADA R"
    sql = sql & " WHERE"
    sql = sql & " RC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND RC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND RC.VEN_CODIGO=V.VEN_CODIGO"
    sql = sql & " AND RC.REP_CODIGO=R.REP_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND RC.CLI_CODIGO=" & XN(txtCliente)
    If txtVendedor.Text <> "" Then sql = sql & " AND RC.VEN_CODIGO=" & XN(txtVendedor)
    If FechaDesde <> "" Then sql = sql & " AND RC.REC_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND RC.REC_FECHA<=" & XDQ(FechaHasta)
    If cboRecibo.List(cboRecibo.ListIndex) <> "(Todos)" Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboRecibo.ItemData(cboRecibo.ListIndex))
    If cboBuscaRep.List(cboBuscaRep.ListIndex) <> "(Todas)" Then sql = sql & " AND RC.REP_CODIGO=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
    sql = sql & " ORDER BY RC.REP_CODIGO,RC.REC_SUCURSAL,RC.REC_NUMERO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!TCO_ABREVIA & Chr(9) & Format(Rec1!REC_SUCURSAL, "0000") & "-" & Format(Rec1!REC_NUMERO, "00000000") _
                               & Chr(9) & Rec1!REC_FECHA & Chr(9) & Rec1!CLI_RAZSOC _
                               & Chr(9) & Rec1!VEN_NOMBRE & Chr(9) & Rec1!REP_RAZSOC _
                               & Chr(9) & Rec1!TCO_CODIGO
            Rec1.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
        GrdModulos.Col = 0
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Recibos...", vbExclamation, TIT_MSGBOX
        txtCliente.SetFocus
    End If
    Rec1.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdListar_Click()
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SISESTILO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    
    If optGeneralTodos.Value = True Then
        Rep.SelectionFormula = ""
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{RECIBO_CLIENTE.CLI_CODIGO}=" & txtCliente.Text
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
        Else
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
        End If
        
        If FechaDesde.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {RECIBO_CLIENTE.REC_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO_CLIENTE.REC_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            End If
        End If
        
        If FechaHasta.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {RECIBO_CLIENTE.REC_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO_CLIENTE.REC_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            End If
        End If
        
        If txtVendedor.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {RECIBO_CLIENTE.VEN_CODIGO}=" & XN(txtVendedor.Text)
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO_CLIENTE.VEN_CODIGO}=" & XN(txtVendedor.Text)
            End If
        End If
        
        If cboBuscaRep.List(cboBuscaRep.ListIndex) <> "(Todas)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {RECIBO_CLIENTE.REP_CODIGO}=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {RECIBO_CLIENTE.REP_CODIGO}=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
            End If
        End If
        
        If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
        ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
        ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
            Rep.Formulas(2) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
        ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
            Rep.Formulas(2) = "FECHA='" & "Al: " & Date & "'"
        End If

        Rep.WindowTitle = "Recibo de Cliente - General..."
        Rep.ReportFileName = DRIVE & DirReport & "rptreciboclientegeneral.rpt"
    End If
    
    If optDetallado.Value = True Then
         Exit Sub
'        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
'            MsgBox "Debe seleccionar un Recibo", vbExclamation, TIT_MSGBOX
'            chkCliente.SetFocus
'            Exit Sub
'        End If
'        Rep.SelectionFormula = ""
'        Rep.SelectionFormula = "{RECIBO_CLIENTE.REC_NUMERO}=" & GrdModulos.TextMatrix(GrdModulos.RowSel, 0) _
'                               & " AND DAY({RECIBO_CLIENTE.REC_FECHA})=" & Day(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) _
'                               & " AND MONTH({RECIBO_CLIENTE.REC_FECHA})=" & Month(GrdModulos.TextMatrix(GrdModulos.RowSel, 1)) _
'                               & " AND YEAR({RECIBO_CLIENTE.REC_FECHA})=" & Year(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
'        Rep.WindowTitle = "Recibo de Cliente - Detallado..."
'        Rep.ReportFileName = DRIVE & DirReport & "rptreciboclientedetalle.rpt"
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
     Rep.Formulas(2) = ""
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
    cboRecibo.ListIndex = -1
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    optGeneralTodos.Value = True
    optPantalla.Value = True
    txtCliente.SetFocus
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    
    optGeneralTodos.Value = True
    
    GrdModulos.Rows = 1
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    lblEstado.Caption = ""

    Call Centrar_pantalla(Me)
    GrdModulos.FormatString = "^Tipo|^Nro Recibo|^Fecha|Cliente|Vendedor|Representada|TIPO RECIBO"
    GrdModulos.ColWidth(0) = 900  'TIPO_RECIBO
    GrdModulos.ColWidth(1) = 1300 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1100 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 3500 'CLIENTE
    GrdModulos.ColWidth(4) = 2500 'VENDEDOR
    GrdModulos.ColWidth(5) = 2500 'REPRESENTADA
    GrdModulos.ColWidth(6) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.Rows = 1
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For I = 0 To 6
        GrdModulos.Col = I
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    '------------------------------------
    'LLENAR COMBO RECIBO
    LlenarComboRecibo
    'CARGO COMBO REPRESENTADA
    CargoComboRepresentada
End Sub

Private Sub LlenarComboRecibo()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'RECIB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboRecibo.AddItem "(Todos)"
        Do While rec.EOF = False
            cboRecibo.AddItem rec!TCO_DESCRI
            cboRecibo.ItemData(cboRecibo.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboRecibo.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CargoComboRepresentada()
    sql = "SELECT REP_RAZSOC,REP_CODIGO FROM REPRESENTADA ORDER BY REP_RAZSOC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboBuscaRep.AddItem "(Todas)"
        Do While rec.EOF = False
            cboBuscaRep.AddItem rec!REP_RAZSOC
            cboBuscaRep.ItemData(cboBuscaRep.NewIndex) = rec!REP_CODIGO
            rec.MoveNext
        Loop
        cboBuscaRep.ListIndex = 0
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
    Set frmListadoReciboCliente = Nothing
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
            txtVendedor.Text = ""
            txtVendedor.SetFocus
        End If
        rec.Close
    End If
End Sub

