VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmControlStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Stock"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
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
   ScaleHeight     =   6480
   ScaleWidth      =   7755
   Begin VB.Frame Frame3 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   18
      Top             =   5265
      Width           =   7620
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   360
         Left            =   5550
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1800
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmControlStock.frx":0000
         Left            =   450
         List            =   "frmControlStock.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   1635
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "frmControlStock.frx":002F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         Picture         =   "frmControlStock.frx":0131
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         Picture         =   "frmControlStock.frx":0233
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   405
      Left            =   4440
      TabIndex        =   6
      Top             =   6030
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6615
      TabIndex        =   8
      Top             =   6030
      Width           =   1035
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmControlStock.frx":0335
      Height          =   405
      Left            =   5520
      TabIndex        =   7
      Top             =   6030
      Width           =   1065
   End
   Begin VB.Frame Frame4 
      Caption         =   "Consulta de Stock por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   75
      TabIndex        =   11
      Top             =   45
      Width           =   7590
      Begin VB.ComboBox cborubro 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1020
         Width           =   3420
      End
      Begin VB.ComboBox cbolinea 
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   3420
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
         Left            =   4665
         TabIndex        =   14
         Text            =   "A"
         Top             =   705
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtProducto 
         Height          =   315
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   0
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox txtDescri 
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
         Left            =   2685
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   330
         Width           =   4740
      End
      Begin VB.CommandButton cmdBuscarProd 
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
         Left            =   2235
         MaskColor       =   &H000000FF&
         Picture         =   "frmControlStock.frx":063F
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Buscar"
         Top             =   330
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   390
         Left            =   5325
         MaskColor       =   &H00404040&
         TabIndex        =   4
         ToolTipText     =   "Buscar Nota de Pedido"
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   2100
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Left            =   435
         TabIndex        =   17
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Left            =   435
         TabIndex        =   16
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   12
         Top             =   375
         Width           =   705
      End
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   3135
      Top             =   6030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3675
      Top             =   6090
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3630
      Left            =   60
      TabIndex        =   5
      Top             =   1560
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   6403
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   15
      Top             =   6120
      Width           =   660
   End
End
Attribute VB_Name = "frmControlStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VStockRep As String
Dim i As Integer

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    Frame3.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cboLinea_LostFocus()
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        cboRubro.Clear
        cargocboRubro
    Else
        cboRubro.Clear
        cboRubro.AddItem "(Todos)"
        cboRubro.ListIndex = 0
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    Dim J As Integer
    
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    GrdModulos.HighLight = flexHighlightNever
    
    sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI, D.DST_STKFIS, P.PTO_CODBARRAS"
    sql = sql & " FROM PRODUCTO P, STOCK D"
    sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
    sql = sql & " AND PTO_ESTADO='N'" 'MUESTRA LOS PRODUCTOS NO DADOS DE BAJA
    If txtProducto.Text <> "" Then
        sql = sql & " AND P.PTO_CODIGO=" & XN(txtProducto.Text)
    End If
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        sql = sql & " AND P.LNA_CODIGO=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
        sql = sql & " AND P.RUB_CODIGO=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
    End If
    sql = sql & " ORDER BY P.PTO_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not rec.EOF
            GrdModulos.AddItem IIf(IsNull(rec!PTO_CODBARRAS), rec!PTO_CODIGO, rec!PTO_CODBARRAS) & Chr(9) & _
                               rec!PTO_DESCRI & Chr(9) & IIf(IsNull(rec!DST_STKFIS), "0", rec!DST_STKFIS)
    
            rec.MoveNext
        Loop
        rec.Close
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.Col = 0
        GrdModulos.SetFocus
    Else
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
        GrdModulos.Rows = 1
        GrdModulos.HighLight = flexHighlightNever
    End If
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdBuscarProd_Click()
    frmBuscar.TipoBusqueda = 2
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtProducto.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtDescri.Text = frmBuscar.grdBuscar.Text
    Else
        txtProducto.SetFocus
    End If
End Sub

Private Sub cmdListar_Click()
    
    lblEstado.Caption = "Buscando Listado..."
    Screen.MousePointer = vbHourglass
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.WindowTitle = "Listado de Stock"
    
    Rep.SelectionFormula = "{STOCK.STK_CODIGO}=" & XN(Sucursal)
    
    If txtProducto.Text <> "" Then
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.PTO_CODIGO}=" & XN(txtProducto.Text)
    End If
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
    End If
    
    Rep.ReportFileName = DRIVE & DirReport & "listadostock.rpt"
    Screen.MousePointer = vbNormal
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    lblEstado.Caption = ""
    
End Sub

Private Sub LlenoTablaTemporal()

    sql = "DELETE FROM TMP_LISTADO_DETALLE_STOCK"
    DBConn.Execute sql
    
'    If chkStockValuado.Value = Unchecked Then
'        If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
'            sql = "INSERT INTO TMP_LISTADO_DETALLE_STOCK"
'            sql = sql & " (STK_CODIGO,PTO_CODIGO,DST_STKFIS,DST_STKPEN)"
'            sql = sql & " SELECT D.STK_CODIGO, P.PTO_CODIGO, D.DST_STKFIS, D.DST_STKPEN"
'            sql = sql & " FROM PRODUCTO P, DETALLE_STOCK D"
'            sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
'            sql = sql & " AND D.STK_CODIGO = " & XN(cboStock.ItemData(cboStock.ListIndex))
'            If txtProducto.Text <> "" Then sql = sql & " AND P.PTO_CODIGO=" & XN(txtProducto.Text)
'            If txtLinea.Text <> "" Then sql = sql & " AND P.LNA_CODIGO=" & XN(txtLinea.Text)
'            If txtRubro.Text <> "" Then sql = sql & " AND P.RUB_CODIGO=" & XN(txtRubro.Text)
'            If txtRepresentada.Text <> "" Then sql = sql & " AND P.REP_CODIGO=" & XN(txtRepresentada.Text)
'        Else
'            sql = "INSERT INTO TMP_LISTADO_DETALLE_STOCK"
'            sql = sql & " (STK_CODIGO,PTO_CODIGO,DST_STKFIS,DST_STKPEN)"
'            sql = sql & " SELECT 1, P.PTO_CODIGO,"
'            sql = sql & " SUM(D.DST_STKFIS) AS DST_STKFIS, SUM(D.DST_STKPEN) AS DST_STKPEN"
'            sql = sql & " FROM PRODUCTO P, DETALLE_STOCK D"
'            sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
'            If txtProducto.Text <> "" Then sql = sql & " AND P.PTO_CODIGO=" & XN(txtProducto.Text)
'            If txtLinea.Text <> "" Then sql = sql & " AND P.LNA_CODIGO=" & XN(txtLinea.Text)
'            If txtRubro.Text <> "" Then sql = sql & " AND P.RUB_CODIGO=" & XN(txtRubro.Text)
'            If txtRepresentada.Text <> "" Then sql = sql & " AND P.REP_CODIGO=" & XN(txtRepresentada.Text)
'            sql = sql & " GROUP BY P.PTO_CODIGO"
'        End If
'    Else
'        'SI VA VALUADO EL STOCK
'        If cboStock.List(cboStock.ListIndex) <> "(Todos)" Then
'            sql = "INSERT INTO TMP_LISTADO_DETALLE_STOCK"
'            sql = sql & " (STK_CODIGO,PTO_CODIGO,DST_STKFIS,DST_STKPEN,DTS_COSTO)"
'            sql = sql & " SELECT D.STK_CODIGO, P.PTO_CODIGO, D.DST_STKFIS, D.DST_STKPEN,L.LIS_COSTO"
'            sql = sql & " FROM PRODUCTO P, DETALLE_STOCK D, DETALLE_LISTA_PRECIO L"
'            sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
'            sql = sql & " AND D.STK_CODIGO = " & XN(cboStock.ItemData(cboStock.ListIndex))
'            sql = sql & " AND L.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
'            sql = sql & " AND P.PTO_CODIGO = L.PTO_CODIGO"
'            If txtProducto.Text <> "" Then sql = sql & " AND P.PTO_CODIGO=" & XN(txtProducto.Text)
'            If txtLinea.Text <> "" Then sql = sql & " AND P.LNA_CODIGO=" & XN(txtLinea.Text)
'            If txtRubro.Text <> "" Then sql = sql & " AND P.RUB_CODIGO=" & XN(txtRubro.Text)
'            If txtRepresentada.Text <> "" Then sql = sql & " AND P.REP_CODIGO=" & XN(txtRepresentada.Text)
'        Else
'            sql = "INSERT INTO TMP_LISTADO_DETALLE_STOCK"
'            sql = sql & " (STK_CODIGO,PTO_CODIGO,DST_STKFIS,DST_STKPEN,DTS_COSTO)"
'            sql = sql & " SELECT 1, P.PTO_CODIGO,"
'            sql = sql & " SUM(D.DST_STKFIS) AS DST_STKFIS, SUM(D.DST_STKPEN) AS DST_STKPEN,L.LIS_COSTO"
'            sql = sql & " FROM PRODUCTO P, DETALLE_STOCK D,DETALLE_LISTA_PRECIO L"
'            sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
'            sql = sql & " AND L.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
'            sql = sql & " AND P.PTO_CODIGO = L.PTO_CODIGO"
'            If txtProducto.Text <> "" Then sql = sql & " AND P.PTO_CODIGO=" & XN(txtProducto.Text)
'            If txtLinea.Text <> "" Then sql = sql & " AND P.LNA_CODIGO=" & XN(txtLinea.Text)
'            If txtRubro.Text <> "" Then sql = sql & " AND P.RUB_CODIGO=" & XN(txtRubro.Text)
'            If txtRepresentada.Text <> "" Then sql = sql & " AND P.REP_CODIGO=" & XN(txtRepresentada.Text)
'            sql = sql & " GROUP BY P.PTO_CODIGO"
'        End If
'    End If
    DBConn.Execute sql
        
    'ACA CALCULO Y AGREGO LOS PEDIDOS PENDIENTES
    sql = "SELECT DISTINCT DNP.PTO_CODIGO, SUM(DNP.DNP_CANTIDAD) AS PEDPEN"
    sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, NOTA_PEDIDO NP"
    sql = sql & " WHERE NP.NPE_NUMERO = DNP.NPE_NUMERO AND NP.EST_CODIGO = 1"
    sql = sql & " AND DNP_MARCA IS NULL"
    sql = sql & " GROUP BY DNP.PTO_CODIGO"
    sql = sql & " ORDER BY DNP.PTO_CODIGO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While Not rec.EOF
            sql = "UPDATE TMP_LISTADO_DETALLE_STOCK"
            sql = sql & " SET DST_PEDPEN=" & XN(Chk0(rec!PEDPEN))
            sql = sql & " WHERE PTO_CODIGO=" & XN(rec!PTO_CODIGO)
            DBConn.Execute sql
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub CmdNuevo_Click()
    txtProducto.Text = ""
    cboLinea.ListIndex = 0
    cboRubro.ListIndex = 0
    txtDescri.Text = ""
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    txtProducto.SetFocus
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmControlStock = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    'Call Centrar_pantalla(Me)
    Me.Left = 0
    Me.Top = 0

    cargocboLinea
    preparogrilla
    cboLinea_LostFocus
    
    lblEstado.Caption = ""
    Frame3.Caption = "Impresora Actual: " & Printer.DeviceName
    cboDestino.ListIndex = 0
End Sub

Private Sub preparogrilla()
    GrdModulos.FormatString = "^Código|<Producto|>Stk Fis."
    GrdModulos.ColWidth(0) = 1200  'CODIGO
    GrdModulos.ColWidth(1) = 5000 'PRODUCTO
    GrdModulos.ColWidth(2) = 1000 'STOCK FISICO
    GrdModulos.Rows = 1
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To 2
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    GrdModulos.HighLight = flexHighlightNever
End Sub

Private Sub cargocboLinea()
    cboLinea.Clear
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboLinea.AddItem "(Todas)"
        Do While rec.EOF = False
            cboLinea.AddItem rec!LNA_DESCRI
            cboLinea.ItemData(cboLinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        cboLinea.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub cargocboRubro()
    cboRubro.Clear
    sql = "SELECT * FROM RUBROS"
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        sql = sql & " WHERE LNA_CODIGO= " & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboRubro.AddItem "(Todos)"
        Do While rec.EOF = False
            cboRubro.AddItem rec!RUB_DESCRI
            cboRubro.ItemData(cboRubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        cboRubro.ListIndex = 0
    Else
        cboRubro.AddItem "(Todos)"
        cboRubro.ListIndex = 0
    End If
    rec.Close
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

Private Sub txtdescri_Change()
    If txtDescri.Text = "" Then
        txtProducto.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtDescri
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
           
   If txtProducto.Text = "" And txtDescri.Text <> "" Then
        Set rec = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI"
        sql = sql & " FROM PRODUCTO P"
        sql = sql & " WHERE P.PTO_DESCRI LIKE '" & Trim(txtDescri.Text) & "%' ORDER BY P.PTO_DESCRI"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                'grdGrilla.SetFocus
                frmBuscar.TipoBusqueda = 2
                frmBuscar.TxtDescriB.Text = txtDescri.Text
                frmBuscar.Show vbModal
                frmBuscar.grdBuscar.Col = 0
                txtProducto.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                frmBuscar.grdBuscar.Col = 1
                txtDescri.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
            Else
                txtProducto.Text = Trim(rec!PTO_CODIGO)
                txtDescri.Text = Trim(rec!PTO_DESCRI)
            End If
        Else
            MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
            txtDescri.Text = ""
        End If
        rec.Close
        Screen.MousePointer = vbNormal
    End If
End Sub

Private Sub txtproducto_Change()
    If txtProducto.Text = "" Then
        txtProducto.Text = ""
        txtDescri.Text = ""
    End If
End Sub

Private Sub txtproducto_GotFocus()
    SelecTexto txtProducto
End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProducto_LostFocus()
    If txtProducto.Text <> "" Then
        sql = "SELECT P.PTO_CODIGO,P.PTO_DESCRI"
        sql = sql & " FROM PRODUCTO P, STOCK D "
        sql = sql & " WHERE P.PTO_CODIGO = D.PTO_CODIGO"
        sql = sql & " AND P.PTO_CODIGO = " & XN(txtProducto.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtDescri.Text = rec!PTO_DESCRI
        Else
            MsgBox "El código no existe", vbInformation
            txtProducto.SetFocus
        End If
        rec.Close
    End If
End Sub


