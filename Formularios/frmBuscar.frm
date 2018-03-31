VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   4815
   ClientLeft      =   1005
   ClientTop       =   1095
   ClientWidth     =   8385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBuscar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8385
   Begin VB.TextBox txtDescriB 
      Height          =   315
      Left            =   765
      TabIndex        =   0
      Top             =   105
      Width           =   2280
   End
   Begin VB.CommandButton cmdBuscaAprox 
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
      Left            =   3135
      MaskColor       =   &H8000000F&
      Picture         =   "frmBuscar.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ejecutar Búsqueda"
      Top             =   105
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3645
      Picture         =   "frmBuscar.frx":2AAC
      TabIndex        =   3
      ToolTipText     =   " Salir "
      Top             =   90
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSFlexGridLib.MSFlexGrid grdBuscar 
      Height          =   4260
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   7514
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      RowHeightMin    =   262
      BackColorSel    =   16761024
      ForeColorSel    =   16777215
      GridColor       =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "<Esc> Salir"
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
      Left            =   4710
      TabIndex        =   6
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<F3> Buscar"
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
      Left            =   3210
      TabIndex        =   5
      Top             =   150
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Buscar:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoBusqueda As Integer
Public TipoEnt As Integer
Public CodigoCli As String
Dim Importe As Double
Public CodListaPrecio As Integer

Public Sub ArmaSQL()
    Select Case TipoBusqueda
    
    Case 1 'CLIENTE
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE "
        sql = sql & " WHERE "
        sql = sql & " CLI_RAZSOC LIKE '" & Trim(TxtDescriB) & "%'"
        sql = sql & " AND CLI_ESTADO=1"
        sql = sql & " ORDER BY CLI_RAZSOC"
    
    Case 2 'PRODUCTOS
        sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, "
        sql = sql & " P.PTO_PRECTO, L.LNA_DESCRI, R.RUB_DESCRI"
        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L"
        sql = sql & " WHERE TPR_CODIGO = " & XN(frmFacturaProveedores.txtCodTipoProv.Text) & _
                    "   AND PROV_CODIGO = " & XN(frmFacturaProveedores.txtCodProveedor.Text) & _
                    "   AND "
        If IsNumeric(TxtDescriB) Then
            sql = sql & " P.PTO_CODIGO=" & XN(TxtDescriB)
        Else
            sql = sql & " P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
        End If
        sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO AND P.RUB_CODIGO=R.RUB_CODIGO"
        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
        sql = sql & " AND PTO_ESTADO='N'"
        sql = sql & " ORDER BY PTO_DESCRI"
        
    Case 3 'SUCURSALES
        sql = "SELECT S.SUC_CODIGO,S.SUC_DESCRI,C.CLI_RAZSOC,S.CLI_CODIGO"
        sql = sql & " FROM SUCURSAL S, CLIENTE C"
        sql = sql & " WHERE S.CLI_CODIGO=C.CLI_CODIGO"
        If IsNumeric(TxtDescriB) Then
            sql = sql & " AND  S.SUC_CODIGO=" & XN(TxtDescriB)
        Else
            sql = sql & " AND S.SUC_DESCRI LIKE '" & Trim(TxtDescriB) & "%' "
        End If
        If CodigoCli <> "" Then
            sql = sql & " AND C.CLI_CODIGO=" & XN(CodigoCli)
        End If
        sql = sql & " AND C.CLI_ESTADO=1"
        sql = sql & " ORDER BY S.SUC_DESCRI"
        
    Case 4 'VENDEDORES
        sql = "SELECT VEN_CODIGO,VEN_NOMBRE,VEN_DOMICI"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE"
        sql = sql & " VEN_NOMBRE LIKE '" & Trim(TxtDescriB) & "%' "
        sql = sql & " ORDER BY VEN_NOMBRE"
    
    Case 5 'PROVEEDORES
        sql = "SELECT TP.TPR_CODIGO,TP.TPR_DESCRI,P.PROV_CODIGO,P.PROV_RAZSOC"
        sql = sql & " FROM TIPO_PROVEEDOR TP, PROVEEDOR P"
        sql = sql & " WHERE"
        sql = sql & " TP.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND PROV_RAZSOC LIKE '" & Trim(TxtDescriB) & "%' "
        sql = sql & " ORDER BY TP.TPR_CODIGO, PROV_RAZSOC"
    
    Case 6 'CHEQUES EN CARTERA
        sql = "SELECT CHE_NUMERO, CHE_IMPORT, CHE_FECVTO, BAN_CODINT, BAN_BANCO, BAN_LOCALIDAD,"
        sql = sql & " BAN_SUCURSAL, BAN_CODIGO, BAN_DESCRI,CES_DESCRI"
        sql = sql & " FROM ChequeEstadoVigente"
        sql = sql & " Where ECH_CODIGO=1" 'CODIGO (1) ES CHEQUE EN CARTERA
        
    Case 7 'PRODUCTOS
        sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, DST_STKCON "
        sql = sql & " FROM PRODUCTO P, RUBROS R, LINEAS L, STOCK S "
        sql = sql & " WHERE "
        If IsNumeric(TxtDescriB) Then
            sql = sql & " P.PTO_CODIGO=" & XN(TxtDescriB)
        Else
            sql = sql & " P.PTO_DESCRI LIKE '" & Trim(TxtDescriB) & "%'"
        End If
        sql = sql & " AND P.LNA_CODIGO=L.LNA_CODIGO "
        sql = sql & " AND P.RUB_CODIGO=R.RUB_CODIGO"
        sql = sql & " AND L.LNA_CODIGO=R.LNA_CODIGO"
        sql = sql & " AND PTO_ESTADO='N'"
        sql = sql & " AND P.PTO_CODIGO=S.PTO_CODIGO"
        sql = sql & " AND DST_STKCON > 0"
        sql = sql & " ORDER BY PTO_DESCRI"
    
    Case 8 'obra social
        sql = "SELECT OS_NUMERO,OS_NOMBRE,OS_DOMICI,OS_TELEFONO"
        sql = sql & " FROM OBRA_SOCIAL"
        sql = sql & " WHERE"
        sql = sql & " OS_NOMBRE LIKE '" & Trim(TxtDescriB) & "%' "
        sql = sql & " ORDER BY OS_NOMBRE"
    End Select
        
End Sub

Public Sub RellenaGrilla(Registro As ADODB.Recordset)
    Select Case TipoBusqueda
    
    Case 1 'CLIENTES
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!CLI_CODIGO) & Chr(9) & _
                Trim(Registro!CLI_RAZSOC)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 2 'PRODUCTOS
        Do While Not Registro.EOF
            grdBuscar.AddItem Registro!PTO_CODIGO & Chr(9) & _
            Trim(Registro!PTO_DESCRI) & Chr(9) & _
            Format(Chk0(Registro!PTO_PRECTO), "0.00") & Chr(9) & _
            ChkNull(Registro!LNA_DESCRI) & Chr(9) & _
            ChkNull(Registro!RUB_DESCRI)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 3 'SUCURSAL
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!SUC_CODIGO) & Chr(9) & _
                Trim(Registro!SUC_DESCRI) & Chr(9) & _
                Trim(Registro!CLI_RAZSOC) & Chr(9) & _
                Trim(Registro!CLI_CODIGO)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
        
    Case 4 'VENDEDORES
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!VEN_CODIGO) & Chr(9) & _
                Trim(Registro!VEN_NOMBRE) & Chr(9) & _
                Trim(Registro!VEN_DOMICI) & Chr(9) & _
                Trim(Registro!VEN_CODIGO)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 5 'PROVEEDORES
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!PROV_CODIGO) & Chr(9) & _
                Trim(Registro!PROV_RAZSOC) & Chr(9) & _
                Trim(Registro!TPR_CODIGO) & Chr(9) & _
                Trim(Registro!TPR_DESCRI)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    Case 6 'CHEQUES EN CARTERA
        Do While Not Registro.EOF
            grdBuscar.AddItem Trim(Registro!BAN_DESCRI) & Chr(9) & _
                Trim(Registro!CHE_NUMERO) & Chr(9) & _
                Trim(Registro!CHE_FECVTO) & Chr(9) & _
                Trim(Valido_Importe(Registro!CHE_IMPORT)) & Chr(9) & _
                Trim(Registro!BAN_CODINT) & Chr(9) & _
                Trim(Registro!BAN_BANCO) & Chr(9) & _
                Trim(Registro!BAN_LOCALIDAD) & Chr(9) & _
                Trim(Registro!BAN_SUCURSAL) & Chr(9) & _
                Trim(Registro!BAN_CODIGO) & Chr(9) & _
                Trim(Registro!CES_DESCRI)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    
    Case 7 'PRODUCTOS en CONSIGNACION
        Do While Not Registro.EOF
            grdBuscar.AddItem Registro!PTO_CODIGO & Chr(9) & _
                         Trim(Registro!PTO_DESCRI) & Chr(9) & _
                         Trim(Registro!DST_STKCON)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    Case 8 'OBRAS SOCIALES
        Do While Not Registro.EOF
            grdBuscar.AddItem Registro!OS_NUMERO & Chr(9) & _
                         Trim(Registro!OS_NOMBRE) & Chr(9) & _
                         Trim(Registro!OS_DOMICI) & Chr(9) & _
                         Trim(Registro!OS_TELEFONO)
            Registro.MoveNext
            grdBuscar.Refresh
        Loop
    End Select
End Sub

Private Sub cmdBuscaAprox_Click()
'    If Trim(TxtDescriB) = "" Then
'        MsgBox "Debe especificar un detalle de Búsqueda"
'        Exit Sub
'    End If
    Screen.MousePointer = vbHourglass
    Set Rec1 = New ADODB.Recordset
    ArmaSQL
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    EliminarFilasDeGrilla grdBuscar
    If Rec1.EOF = False Then
        RellenaGrilla Rec1
        grdBuscar.SetFocus
        If grdBuscar.Rows > 1 Then grdBuscar.Col = 0
        grdBuscar.HighLight = flexHighlightAlways
    Else
        MsgBox "No se han encontrado datos relacionados"
        SelecTexto TxtDescriB
        TxtDescriB.SetFocus
        grdBuscar.HighLight = flexHighlightNever
    End If
    Rec1.Close
    Set Rec1 = Nothing
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdSalir_Click()
    grdBuscar.Clear
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    grdBuscar.Rows = 1
    Importe = 0
    TxtDescriB.Text = ""
    TxtDescriB.SetFocus
    
    Select Case TipoBusqueda
    Case 1 'CLIENTES
        Me.Caption = "Buscar:[Clientes]"
        grdBuscar.Width = 6400
        grdBuscar.Cols = 2
        grdBuscar.FormatString = "Código|Razón Social"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 5000
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 2 'PRODUCTOS de Kiosco
        Me.Caption = "Buscar:[Productos]"
        grdBuscar.Width = 11400
        grdBuscar.Cols = 5
        grdBuscar.FormatString = ">Código|Descripción|Precio|Línea|Rubro"
        grdBuscar.ColWidth(0) = 800
        grdBuscar.ColWidth(1) = 3000
        grdBuscar.ColWidth(2) = 1000
        grdBuscar.ColWidth(3) = 2500
        grdBuscar.ColWidth(4) = 2500
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 3 'SUCURSALES
        Me.Caption = "Buscar:[Sucursales]"
        grdBuscar.Width = 11400
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "Código|Razón Social|Cliente|CODCLI"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 5000
        grdBuscar.ColWidth(2) = 5000
        grdBuscar.ColWidth(3) = 0
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 4 'VENDEDORES
        Me.Caption = "Buscar:[Vendedores]"
        grdBuscar.Width = 7900
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "Número|Nombre|Domicilio|NUMVENDEDOR"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 3500
        grdBuscar.ColWidth(2) = 3000
        grdBuscar.ColWidth(3) = 0
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
        
    Case 5 'PROVEDORES  de Kiosco
        Me.Caption = "Buscar:[Proveedores]"
        grdBuscar.Width = 8200
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "Número|Razón Social|Cod Tipo Prov|Tipo Proveedor"
        grdBuscar.ColWidth(0) = 800
        grdBuscar.ColWidth(1) = 3000
        grdBuscar.ColWidth(2) = 800
        grdBuscar.ColWidth(3) = 3000
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    
    Case 6 'CHEQUES EN CARTERA
        Me.Caption = "Buscar:[Cheques en Cartera (de Terceros)]"
        grdBuscar.Width = 10200
        grdBuscar.Cols = 10
        grdBuscar.FormatString = "Banco|^Cheuqe Nro|^Fecha Vto|>Importe|BAN_CODINT|BAN_BANCO" _
                                & "|BAN_LOCALIDAD|BAN_SUCURSAL|BAN_CODIGO|Estado"
        grdBuscar.ColWidth(0) = 4000 'Banco
        grdBuscar.ColWidth(1) = 1200 'Cheuqe Nro
        grdBuscar.ColWidth(2) = 1100 'Fecha Vto
        grdBuscar.ColWidth(3) = 1100 'Importe
        grdBuscar.ColWidth(4) = 0    'BAN_CODINT
        grdBuscar.ColWidth(5) = 0    'BAN_BANCO
        grdBuscar.ColWidth(6) = 0    'BAN_LOCALIDAD
        grdBuscar.ColWidth(7) = 0    'BAN_SUCURSAL
        grdBuscar.ColWidth(8) = 0    'BAN_CODIGO
        grdBuscar.ColWidth(9) = 2100 'CES_DESCRI
        cmdBuscaAprox_Click
    
    Case 7 'Productos en Consignacion del Kiosco
        Me.Caption = "Buscar:[Productos en Consignación]"
        grdBuscar.Width = 6200
        grdBuscar.Cols = 3
        grdBuscar.FormatString = ">Código|Descripción|Cantidad"
        grdBuscar.ColWidth(0) = 800
        grdBuscar.ColWidth(1) = 4000
        grdBuscar.ColWidth(2) = 1000
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
        
    Case 8 'OBRA SOCIAL
        Me.Caption = "Buscar:[Obras Sociales]"
        grdBuscar.Width = 7900
        grdBuscar.Cols = 4
        grdBuscar.FormatString = "Número|Nombre|Domicilio|Telefono"
        grdBuscar.ColWidth(0) = 1000
        grdBuscar.ColWidth(1) = 3000
        grdBuscar.ColWidth(2) = 2500
        grdBuscar.ColWidth(3) = 1500
        If TxtDescriB.Text <> "" Then cmdBuscaAprox_Click
    End Select
    Me.Width = grdBuscar.Width + 120
    
    'PARA DARLE FORMATO A LA GRILLA
    grdBuscar.BorderStyle = flexBorderNone
    grdBuscar.row = 0
    For i = 0 To grdBuscar.Cols - 1
        grdBuscar.Col = i
        grdBuscar.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdBuscar.CellBackColor = &H808080    'GRIS OSCURO
        grdBuscar.CellFontBold = True
    Next
    
    Call Centrar_pantalla(Me)
    If grdBuscar.Rows > 1 Then
        grdBuscar.SetFocus
        grdBuscar.HighLight = flexHighlightAlways
    Else
        TxtDescriB.SetFocus
        grdBuscar.HighLight = flexHighlightNever
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub grdBuscar_Click()
    OrdenarGrilla grdBuscar
End Sub

Private Sub GrdBuscar_dblClick()
    If grdBuscar.Rows > 0 Then
        Select Case TipoBusqueda
        Case 1
            grdBuscar.Col = 0
        Case 2
            grdBuscar.Col = 0
        Case 3
            grdBuscar.Col = 0
        Case 4
            grdBuscar.Col = 0
        Case 5
            grdBuscar.Col = 0
'        Case 5
'            grdBuscar.Col = 0
'        Case 6
'            grdBuscar.Col = 0
'        Case 7
'            grdBuscar.Col = 0
'        Case 8
'            grdBuscar.Col = 0
'        Case 99
'            grdBuscar.Col = 0
        End Select
        Me.Hide
    End If
End Sub

Private Sub grdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GrdBuscar_dblClick
    End If
End Sub

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub txtDescriB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        cmdBuscaAprox_Click
    End If
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
