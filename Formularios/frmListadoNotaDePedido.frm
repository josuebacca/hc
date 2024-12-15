VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmListadoNotaDePedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Nota de Pedido"
   ClientHeight    =   6975
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
   ScaleHeight     =   6975
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoNotaDePedido.frx":0000
      Height          =   750
      Left            =   8865
      Picture         =   "frmListadoNotaDePedido.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6180
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   6615
      Top             =   5100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   7275
      Top             =   5145
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9750
      Picture         =   "frmListadoNotaDePedido.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6180
      Width           =   840
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   750
      Left            =   7995
      Picture         =   "frmListadoNotaDePedido.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6180
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
      TabIndex        =   32
      Top             =   5070
      Width           =   10425
      Begin VB.OptionButton optDetalladoVarios 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado (Varios)"
         Height          =   255
         Left            =   3645
         TabIndex        =   10
         Top             =   240
         Width           =   2370
      End
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado Detallado (Una)"
         Height          =   255
         Left            =   735
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2205
      End
      Begin VB.OptionButton optGeneralTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "... Listado General"
         Height          =   210
         Left            =   7155
         TabIndex        =   11
         Top             =   240
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
      TabIndex        =   28
      Top             =   5700
      Width           =   7845
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   13
         Top             =   360
         Width           =   1050
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   660
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   840
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nota de Pedido por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   90
      TabIndex        =   17
      Top             =   15
      Width           =   10395
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   2415
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1635
         Width           =   3630
      End
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
         Left            =   6705
         TabIndex        =   35
         Text            =   "A"
         Top             =   1365
         Visible         =   0   'False
         Width           =   630
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
         Height          =   315
         Left            =   3300
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoNotaDePedido.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Buscar Vendedor"
         Top             =   930
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.ComboBox cboRepresentada 
         Height          =   315
         Left            =   2415
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1275
         Width           =   3630
      End
      Begin VB.TextBox txtDesSuc 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3735
         MaxLength       =   50
         TabIndex        =   22
         Tag             =   "Descripción"
         Top             =   585
         Width           =   4410
      End
      Begin VB.TextBox txtSucursal 
         Height          =   315
         Left            =   2415
         MaxLength       =   40
         TabIndex        =   1
         Top             =   570
         Width           =   855
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   2415
         MaxLength       =   40
         TabIndex        =   0
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox txtDesCli 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3735
         MaxLength       =   50
         TabIndex        =   21
         Tag             =   "Descripción"
         Top             =   225
         Width           =   4410
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "&Buscar"
         Height          =   450
         Left            =   6120
         MaskColor       =   &H000000FF&
         TabIndex        =   7
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   1845
         UseMaskColor    =   -1  'True
         Width           =   2040
      End
      Begin VB.TextBox txtDesVen 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3735
         TabIndex        =   20
         Top             =   930
         Width           =   4410
      End
      Begin VB.TextBox txtVendedor 
         Height          =   315
         Left            =   2415
         TabIndex        =   2
         Top             =   930
         Width           =   855
      End
      Begin VB.CommandButton cmdBuscarSuc 
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
         Left            =   3300
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoNotaDePedido.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Buscar"
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   405
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
         Left            =   3300
         MaskColor       =   &H000000FF&
         Picture         =   "frmListadoNotaDePedido.frx":17FC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Top             =   1995
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   2415
         TabIndex        =   5
         Top             =   1995
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   1275
         TabIndex        =   36
         Top             =   1650
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Representada:"
         Height          =   195
         Left            =   1275
         TabIndex        =   33
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1275
         TabIndex        =   27
         Top             =   630
         Width           =   660
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
         Left            =   1275
         TabIndex        =   26
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1275
         TabIndex        =   25
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3870
         TabIndex        =   24
         Top             =   2055
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   1275
         TabIndex        =   23
         Top             =   975
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   2595
      Left            =   90
      TabIndex        =   8
      Top             =   2460
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4577
      _Version        =   393216
      Cols            =   5
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
      Left            =   8040
      TabIndex        =   37
      Top             =   5865
      Width           =   660
   End
End
Attribute VB_Name = "frmListadoNotaDePedido"
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

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    sql = "SELECT NP.*, C.CLI_RAZSOC, S.SUC_DESCRI"
    sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, SUCURSAL S"
    sql = sql & " WHERE"
    sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
    sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente.Text)
    If txtSucursal.Text <> "" Then sql = sql & " AND NP.SUC_CODIGO=" & XN(txtSucursal.Text)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor.Text)
    If cboRepresentada.List(cboRepresentada.ListIndex) <> "(Todas)" Then
        sql = sql & " AND NP.REP_CODIGO=" & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex))
    End If
    If cboEstado.List(cboEstado.ListIndex) <> "(Todos)" Then
        sql = sql & " AND NP.EST_CODIGO=" & XN(cboEstado.ItemData(cboEstado.ListIndex))
    End If
    If FechaDesde <> "" Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde.Text)
    If FechaHasta <> "" Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta.Text)
    sql = sql & " ORDER BY NPE_NUMERO,NPE_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.HighLight = flexHighlightAlways
        Do While rec.EOF = False
            GrdModulos.AddItem Format(rec!NPE_NUMERO, "00000000") & Chr(9) & rec!NPE_FECHA _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!SUC_DESCRI & Chr(9) & rec!FPG_CODIGO
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
    'Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SISESTILO"
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    
    'NOTA DE PEDIDO GENERAL
    If optGeneralTodos.Value = True Then
        Rep.SelectionFormula = ""
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{NOTA_PEDIDO.CLI_CODIGO}=" & txtCliente.Text
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: " & txtDesCli & "'"
        Else
            Rep.Formulas(0) = "CLIENTE='" & "Cliente: Todos'"
        End If
        
        If txtSucursal.Text <> "" Then
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.SUC_CODIGO}=" & txtSucursal.Text
            Rep.Formulas(1) = "SUCURSAL='" & "Sucursal: " & txtDesSuc & "'"
        Else
            Rep.Formulas(1) = "SUCURSAL='" & "Sucursal: Todas'"
        End If
        
        If txtVendedor.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.VEN_CODIGO}= " & XN(txtVendedor.Text)
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.VEN_CODIGO}= " & XN(txtVendedor.Text)
            End If
        End If
        
        If cboEstado.List(cboEstado.ListIndex) <> "(Todos)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.EST_CODIGO}= " & XN(cboEstado.ItemData(cboEstado.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.EST_CODIGO}= " & XN(cboEstado.ItemData(cboEstado.ListIndex))
            End If
        End If
        
        If cboRepresentada.List(cboRepresentada.ListIndex) <> "(Todas)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.REP_CODIGO}= " & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.REP_CODIGO}= " & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex))
            End If
            Rep.Formulas(3) = "REPRESENTADA='" & "Representada: " & cboRepresentada.List(cboRepresentada.ListIndex) & "'"
        End If
        
        If FechaDesde.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.NPE_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.NPE_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            End If
        End If
        
        If FechaHasta.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.NPE_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.NPE_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
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
    
        Rep.WindowTitle = "Nota de Pedido - General..."
        Rep.ReportFileName = DRIVE & DirReport & "rptnotapedidogeneral.rpt"
    End If
    
    'NOTA DE PEDIDO DETALLADO (UNA NOTA DE PEDIDO SOLA)
    If optDetallado.Value = True Then
        Rep.Formulas(0) = ""
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 0) = "" Then
            MsgBox "Debe seleccionar una Nota de Pedido", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
            Exit Sub
        End If
        Rep.SelectionFormula = ""
        Rep.SelectionFormula = "{NOTA_PEDIDO.NPE_NUMERO}=" & GrdModulos.TextMatrix(GrdModulos.RowSel, 0) _
                               & " AND {NOTA_PEDIDO.NPE_FECHA}= DATE (" & Mid(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 7, 4) & "," & Mid(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4, 2) & "," & Mid(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 1, 2) & ")"
        
        Rep.WindowTitle = "Nota de Pedido - Detallado..."
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "" Then 'SI VA SIN DETALLE
            Rep.ReportFileName = DRIVE & DirReport & "rptnotapedidodetalle.rpt"
        Else 'SI VA CON DETALLE
            Rep.ReportFileName = DRIVE & DirReport & "rptnotapedidodetallePrecio.rpt"
        End If
    End If
    
    'NOTA DE PEDIDO DETALLE (VARIOS)
    If optDetalladoVarios.Value = True Then
        Rep.Formulas(0) = ""
        Rep.SelectionFormula = ""
        If txtCliente.Text <> "" Then
            Rep.SelectionFormula = "{NOTA_PEDIDO.CLI_CODIGO}=" & txtCliente.Text
        End If
        If txtSucursal.Text <> "" Then
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.SUC_CODIGO}=" & txtSucursal.Text
        End If
                
        If txtVendedor.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.VEN_CODIGO}= " & XN(txtVendedor.Text)
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.VEN_CODIGO}= " & XN(txtVendedor.Text)
            End If
        End If
        
        If cboEstado.List(cboEstado.ListIndex) <> "(Todos)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.EST_CODIGO}= " & XN(cboEstado.ItemData(cboEstado.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.EST_CODIGO}= " & XN(cboEstado.ItemData(cboEstado.ListIndex))
            End If
        End If

        If cboRepresentada.List(cboRepresentada.ListIndex) <> "(Todas)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.REP_CODIGO}= " & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex))
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.REP_CODIGO}= " & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex))
            End If
        End If
        
        If FechaDesde.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.NPE_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.NPE_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
            End If
        End If
        
        If FechaHasta.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {NOTA_PEDIDO.NPE_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {NOTA_PEDIDO.NPE_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
            End If
        End If
        
        If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
            Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
        ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
            Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
        ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
            Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
        ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
            Rep.Formulas(0) = "FECHA='" & "Al: " & Date & "'"
        End If
    
        Rep.WindowTitle = "Nota de Pedido - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "rptnotapedidodetalle.rpt"
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

Private Sub cmdBuscarSuc_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.txtDescriB = ""
    If txtCliente.Text <> "" Then
        frmBuscar.CodigoCli = txtCliente.Text
    Else
        frmBuscar.CodigoCli = ""
    End If
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 3
        txtCliente.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 0
        txtSucursal.Text = frmBuscar.grdBuscar.Text
        txtSucursal.SetFocus
        txtSucursal_LostFocus
    Else
        txtSucursal.SetFocus
    End If
End Sub

Private Sub CmdNuevo_Click()
    txtSucursal.Text = ""
    txtDesSuc.Text = ""
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtVendedor.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.Rows = 1
    cboRepresentada.ListIndex = 0
    cboEstado.ListIndex = 0
    optDetallado.Value = True
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
    GrdModulos.FormatString = "^Número|^Fecha|Cliente|Sucursal|Forma de pago"
    GrdModulos.ColWidth(0) = 1000
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 4050
    GrdModulos.ColWidth(3) = 4050
    GrdModulos.ColWidth(4) = 0
    'GrdModulos.Rows = 2
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    Dim I As Integer
    For I = 0 To 4
        GrdModulos.Col = I
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080 'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    '-----------------------------------
    'CRAGO COMBO REPRESENTADA
    cboRepresentada.AddItem "(Todas)"
    Call CargoComboBox(cboRepresentada, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboRepresentada.ListIndex = 0
    
    'CARGO COMBO ESTADO
    CargoComboEstado
    optDetallado.Value = True
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

Private Sub txtSucursal_Change()
    If txtSucursal.Text = "" Then
        txtDesSuc.Text = ""
    End If
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto txtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    
    If txtSucursal.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, SUC_DESCRI FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtSucursal)
        If txtCliente.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(txtCliente)
        End If
        lblEstado.Caption = "Buscando..."
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliente.Text = Rec1!CLI_CODIGO
            txtCliente_LostFocus
            txtDesSuc.Text = Rec1!SUC_DESCRI
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDesSuc.Text = ""
            txtSucursal.SetFocus
             Rec1.Close
            Exit Sub
        End If
        Rec1.Close
    End If
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoNotaDePedido = Nothing
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

