VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form FrmListadePrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Precios"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
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
   ScaleHeight     =   6345
   ScaleWidth      =   11760
   Begin VB.Frame FrameBuscaProducto 
      Caption         =   "Agergar Productos a la Lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   465
      TabIndex        =   39
      Top             =   1065
      Visible         =   0   'False
      Width           =   9435
      Begin VB.CommandButton cmdBuscarAgregar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7605
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   375
         Width           =   1590
      End
      Begin VB.TextBox txtDescriAgergar 
         Height          =   315
         Left            =   1260
         TabIndex        =   45
         Top             =   375
         Width           =   5385
      End
      Begin VB.CommandButton cmdSalirFrame 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7665
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3720
         Width           =   1590
      End
      Begin VB.CommandButton CmdSelec 
         Caption         =   "&Seleccionar todo"
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   3720
         Width           =   1590
      End
      Begin VB.CommandButton CmdDeselec 
         Caption         =   "&Deseleccionar todo"
         Height          =   315
         Left            =   1725
         TabIndex        =   41
         Top             =   3720
         Width           =   1590
      End
      Begin VB.CommandButton cmdAceptarViajes 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6045
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3720
         Width           =   1590
      End
      Begin MSFlexGridLib.MSFlexGrid grdGrilla2 
         Height          =   2850
         Left            =   105
         TabIndex        =   44
         Top             =   795
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   5027
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         ForeColorSel    =   16777215
         GridColor       =   -2147483633
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   225
         TabIndex        =   46
         Top             =   435
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   465
      Left            =   7335
      Picture         =   "FrmListadePrecios.frx":0000
      TabIndex        =   38
      ToolTipText     =   "Nueva Lista de Precios"
      Top             =   5820
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   480
      Left            =   6465
      Picture         =   "FrmListadePrecios.frx":0E42
      TabIndex        =   37
      ToolTipText     =   "Guardar Lista de Precios"
      Top             =   5805
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   465
      Left            =   10830
      Picture         =   "FrmListadePrecios.frx":114C
      TabIndex        =   36
      Top             =   5820
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   465
      Left            =   8205
      Picture         =   "FrmListadePrecios.frx":1456
      TabIndex        =   35
      ToolTipText     =   "Eliminar Lista de Precios"
      Top             =   5820
      Width           =   870
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   465
      Left            =   9960
      Picture         =   "FrmListadePrecios.frx":1760
      TabIndex        =   34
      ToolTipText     =   "Imprimir lista de Precios"
      Top             =   5820
      Width           =   870
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   465
      Left            =   9090
      Picture         =   "FrmListadePrecios.frx":3F02
      TabIndex        =   33
      ToolTipText     =   "Limpiar"
      Top             =   5820
      Width           =   870
   End
   Begin VB.CommandButton cmdPrecios 
      Height          =   1155
      Left            =   11220
      Picture         =   "FrmListadePrecios.frx":420C
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Modificar Precios"
      Top             =   2040
      Width           =   465
   End
   Begin VB.CommandButton cmdQuitar 
      Height          =   690
      Left            =   11235
      Picture         =   "FrmListadePrecios.frx":4AD6
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Quitar Producto de la Lista de precios"
      Top             =   4335
      Width           =   465
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   615
      Left            =   11235
      Picture         =   "FrmListadePrecios.frx":4DE0
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Agegar Producto a la Lista de Precios"
      Top             =   3720
      Width           =   465
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
      Left            =   3555
      TabIndex        =   25
      Text            =   "A"
      Top             =   5910
      Visible         =   0   'False
      Width           =   390
   End
   Begin TabDlg.SSTab TabPrecios 
      Height          =   1860
      Left            =   5220
      TabIndex        =   21
      Top             =   2415
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   3281
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmListadePrecios.frx":50EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Cambiar Precio..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   90
         TabIndex        =   22
         Top             =   60
         Width           =   2985
         Begin VB.CommandButton cmdSalirP 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   1860
            TabIndex        =   14
            ToolTipText     =   "Cancelar"
            Top             =   1230
            Width           =   1035
         End
         Begin VB.TextBox txtActual 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1620
            TabIndex        =   12
            Top             =   780
            Width           =   1245
         End
         Begin VB.CommandButton cmdAceptarP 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   810
            TabIndex        =   13
            ToolTipText     =   "Guardar Precio"
            Top             =   1230
            Width           =   1035
         End
         Begin VB.TextBox txtAnterior 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   405
            Width           =   1245
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Precio Actual:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   150
            TabIndex        =   24
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Precio Anterior:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   150
            TabIndex        =   23
            Top             =   450
            Width           =   1275
         End
      End
   End
   Begin VB.Frame freLista 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   30
      TabIndex        =   15
      Top             =   0
      Width           =   11670
      Begin VB.ComboBox cbodescri 
         Height          =   315
         Left            =   5655
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2450
      End
      Begin FechaCtl.Fecha Fecha1 
         Height          =   285
         Left            =   3105
         TabIndex        =   10
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.TextBox txtcodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   825
         TabIndex        =   9
         Top             =   255
         Width           =   750
      End
      Begin VB.TextBox TxtDescriB 
         Height          =   285
         Left            =   5655
         MaxLength       =   40
         TabIndex        =   1
         Top             =   255
         Width           =   2450
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   9615
         TabIndex        =   2
         ToolTipText     =   "Buscar Datos de la Lista de Precios"
         Top             =   195
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vigencia desde:"
         Height          =   195
         Left            =   1935
         TabIndex        =   19
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   18
         Top             =   285
         Width           =   870
      End
   End
   Begin VB.Frame freOpciones 
      Caption         =   "Opciones de Consulta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   30
      TabIndex        =   16
      Top             =   675
      Width           =   11670
      Begin VB.ComboBox cboListaPrecio 
         Height          =   315
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   225
         Width           =   3420
      End
      Begin VB.CommandButton cmdfiltrar 
         Caption         =   "&Filtrar"
         Height          =   375
         Left            =   9660
         TabIndex        =   7
         ToolTipText     =   "Buscar Productos"
         Top             =   525
         Width           =   1455
      End
      Begin VB.ComboBox cbolinea 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   585
         Width           =   3420
      End
      Begin VB.ComboBox cborubro 
         Height          =   315
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   585
         Width           =   3420
      End
      Begin VB.TextBox txtproducto 
         Height          =   315
         Left            =   930
         TabIndex        =   3
         Top             =   225
         Width           =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista Precio:"
         Height          =   195
         Index           =   1
         Left            =   4605
         TabIndex        =   32
         Top             =   300
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Left            =   45
         TabIndex        =   28
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   45
         TabIndex        =   27
         Top             =   285
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Left            =   4605
         TabIndex        =   26
         Top             =   630
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   4095
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Haciendo doble click sobre la grilla puede modificar el precio del Producto"
      Top             =   1665
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   6
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
   Begin Crystal.CrystalReport Rep 
      Left            =   6090
      Top             =   5940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   105
      TabIndex        =   17
      Top             =   5910
      Width           =   660
   End
End
Attribute VB_Name = "FrmListadePrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodigoProducto As String
Dim J As Integer
Dim i As Integer

Private Sub cbodescri_Click()
    If cbodescri.ListIndex <> -1 Then
        txtcodigo.Text = cbodescri.ItemData(cbodescri.ListIndex)
    End If
End Sub

Private Sub cbodescri_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub cbodescri_LostFocus()
    If cbodescri.ListIndex <> -1 Then
        txtcodigo.Text = cbodescri.ItemData(cbodescri.ListIndex)
    End If
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

Private Sub cmdAceptarP_Click()
    TabPrecios.Visible = False
    freLista.Enabled = True
    freOpciones.Enabled = True
    'frebotones.Enabled = True
    
    On Error GoTo HayError
        If txtcodigo = "" Then
            'ENTRA ACA CUANDO ES UNA LISTA NUEVA
            GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = Valido_Importe(txtActual.Text)
            'GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = Valido_Importe(txtCostoActual.Text)
        Else
            'ENTRA ACA CUANDO ACTUALIZO UN PRECIO DE LISTA
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Actualizando ..."
            GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = Valido_Importe(Chk0(txtActual.Text))
            
            DBConn.BeginTrans
            SQL = "UPDATE DETALLE_LISTA_PRECIO"
            SQL = SQL & " SET LIS_PRECIO=" & XN(txtActual.Text)
            'sql = sql & " ,LIS_COSTO=" & XN(txtCostoActual.Text)
            SQL = SQL & " WHERE LIS_CODIGO=" & XN(txtcodigo.Text)
            SQL = SQL & " AND PTO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 0))
            DBConn.Execute SQL
            
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            DBConn.CommitTrans
        End If
    Exit Sub
            
HayError:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub Agregoproducto()
    On Error GoTo HayError
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    
    SQL = "SELECT PTO_DESCRI, PTO_PREVTA, PTO_CODIGO, L.LNA_DESCRI, M.MAR_DESCRI, P.PTO_CODBARRAS"
    SQL = SQL & " FROM PRODUCTO P, LINEAS L, MARCAS M"
    SQL = SQL & " WHERE"
    SQL = SQL & " L.LNA_CODIGO=P.LNA_CODIGO"
    SQL = SQL & " AND M.MAR_CODIGO = P.MAR_CODIGO"
    SQL = SQL & " AND P.PTO_CODIGO = " & XN(CodigoProducto)
    SQL = SQL & " ORDER BY PTO_DESCRI"
        
    lblEstado.Caption = "Buscando..."
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
            If txtcodigo.Text = "" Then
                'ACA ENTRA CUANDO ESTOY CREANDO UNA NUEVA LISTA DE PRECIO
                GrdModulos.AddItem Trim(rec!PTO_CODIGO) & Chr(9) & IIf(IsNull(rec!PTO_CODBARRAS), Trim(rec!PTO_CODIGO), Trim(rec!PTO_CODBARRAS)) _
                                   & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                                   Trim(rec!LNA_DESCRI) & Chr(9) & Trim(rec!MAR_DESCRI) & Chr(9) & Valido_Importe(Chk0(rec!PTO_PREVTA))

            Else
                 'INSERTO EN LA LISTA DE PRECIO Y EN DETALLE DE LISTA DE PRECIO
                 SQL = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,PTO_CODIGO,LIS_PRECIO)"
                 SQL = SQL & " VALUES ("
                 SQL = SQL & XN(txtcodigo) & ","
                 SQL = SQL & XN(CodigoProducto) & ","
                 SQL = SQL & XN(Chk0(rec!PTO_PREVTA)) & " )"
                 DBConn.Execute SQL
                
                'INSERTO EN LA GRILLA
                GrdModulos.AddItem Trim(rec!PTO_CODIGO) & Chr(9) & IIf(IsNull(rec!PTO_CODBARRAS), Trim(rec!PTO_CODIGO), Trim(rec!PTO_CODBARRAS)) _
                                   & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                                   Trim(rec!LNA_DESCRI) & Chr(9) & Trim(rec!MAR_DESCRI) & Chr(9) & Valido_Importe(Chk0(rec!PTO_PREVTA))
            End If
    End If
    rec.Close
        
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    Exit Sub
    
HayError:
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdAceptarViajes_Click()
    If grdGrilla2.Rows > 1 Then
        For i = 1 To grdGrilla2.Rows - 1
            If grdGrilla2.TextMatrix(i, 5) = "SI" Then
                CodigoProducto = grdGrilla2.TextMatrix(i, 0)
                Agregoproducto
            End If
        Next
        cmdSalirFrame_Click
    End If
End Sub

Private Sub cmdAgregar_Click()
    CodigoProducto = ""
    grdGrilla2.Rows = 1
    txtDescriAgergar.Text = ""
    FrameBuscaProducto.Visible = True
'    Dim cSQL As String
'    Dim hSQL As String
'    Dim B As CBusqueda
'    'Dim posicion As Integer
'    Dim cadena As String
'
'    Set B = New CBusqueda
'    With B
'        'Set .Conn = DBConn
'        cSQL = "SELECT P.PTO_DESCRI, P.PTO_CODIGO, P.PTO_PREVTA, L.LNA_DESCRI"
'        cSQL = cSQL & " FROM PRODUCTO P, LINEAS L"
'        cSQL = cSQL & " WHERE"
'        cSQL = cSQL & " L.LNA_CODIGO=P.LNA_CODIGO"
'        cSQL = cSQL & " AND P.PTO_CODIGO NOT IN ("
'        cSQL = cSQL & " SELECT D.PTO_CODIGO"
'        cSQL = cSQL & " FROM DETALLE_LISTA_PRECIO D"
'        cSQL = cSQL & " WHERE D.LIS_CODIGO=" & cbodescri.ItemData(cbodescri.ListIndex) & ")"
'
'        hSQL = "Descripción, Código, Precio, Linea"
'        .sql = cSQL
'        .Headers = hSQL
'        .Field = "PTO_DESCRI"
'        campo1 = .Field
'        .Field = "PTO_CODIGO"
'        campo2 = .Field
'        .Field = "PTO_PREVTA"
'        campo3 = .Field
'        .Field = "LNA_DESCRI"
'        campo4 = .Field
'        .OrderBy = "PTO_DESCRI"
'        camponumerico = False
'        .Titulo = "Busqueda de Productos :"
'        .MaxRecords = 1
'        .Show
'
'        ' utilizar la coleccion de datos devueltos
'        If .ResultFields.Count > 0 Then
'            CodigoProducto = .ResultFields(2)
'            Agregoproducto
'        End If
'    End With
'
'    Set B = Nothing
  
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    If Trim(txtcodigo.Text) <> "" Then
        If MsgBox("Seguro desea eliminar La Lista de Precios: " & Trim(cbodescri.Text) & "? ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Eliminando ..."
            
            DBConn.Execute "DELETE FROM DETALLE_LISTA_PRECIO WHERE LIS_CODIGO = " & XN(txtcodigo)
            DBConn.Execute "DELETE FROM LISTA_PRECIO WHERE LIS_CODIGO = " & XN(txtcodigo)
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            cmdCancelar_Click
        End If
    End If
    Exit Sub
    
CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    
    LimpiarOpciones
    Screen.MousePointer = vbHourglass
    SQL = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, LP.LIS_FECHA, P.PTO_CODBARRAS,"
    SQL = SQL & " D.LIS_PRECIO,P.PTO_CODIGO, LP.LIS_CODIGO, M.MAR_DESCRI"
    SQL = SQL & " FROM PRODUCTO P, LINEAS L, MARCAS M,"
    SQL = SQL & " LISTA_PRECIO LP, DETALLE_LISTA_PRECIO D"
    SQL = SQL & " WHERE P.LNA_CODIGO = L.LNA_CODIGO AND D.PTO_CODIGO = P.PTO_CODIGO"
    SQL = SQL & " AND LP.LIS_CODIGO = D.LIS_CODIGO"
    SQL = SQL & " AND M.MAR_CODIGO = P.MAR_CODIGO"
    SQL = SQL & " AND LP.LIS_CODIGO = " & XN(cbodescri.ItemData(cbodescri.ListIndex))
    'sql = sql & " AND LP.LIS_DESCRI LIKE '" & Trim(cbodescri.List(cbodescri.ListIndex)) & "%'"
    SQL = SQL & " ORDER BY P.PTO_DESCRI"
        
    lblEstado.Caption = "Buscando..."
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        GrdModulos.Rows = 1
        Do While Not rec.EOF
           GrdModulos.AddItem Trim(rec!PTO_CODIGO) & Chr(9) & IIf(IsNull(rec!PTO_CODBARRAS), Trim(rec!PTO_CODIGO), Trim(rec!PTO_CODBARRAS)) _
                              & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                              Trim(rec!LNA_DESCRI) & Chr(9) & Trim(rec!MAR_DESCRI) & Chr(9) & Valido_Importe(rec!LIS_PRECIO)
            rec.MoveNext
        Loop
        rec.MoveFirst
        txtcodigo.Text = rec!LIS_CODIGO
        Fecha1.Text = ChkNull(rec!LIS_FECHA)
        cmdImprimir.Enabled = True
        CmdBorrar.Enabled = True
    Else
        lblEstado.Caption = ""
        MsgBox "No hay Productos Asignados a la lista de Precio", vbInformation, TIT_MSGBOX
        Me.cbodescri.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdBuscarAgregar_Click()

    cSQL = "SELECT P.PTO_DESCRI, P.PTO_CODIGO, P.PTO_PREVTA, L.LNA_DESCRI, P.PTO_CODBARRAS"
    cSQL = cSQL & " FROM PRODUCTO P, LINEAS L"
    cSQL = cSQL & " WHERE"
    cSQL = cSQL & " L.LNA_CODIGO=P.LNA_CODIGO"
    If txtDescriAgergar.Text <> "" Then
        cSQL = cSQL & " AND P.PTO_DESCRI LIKE '" & Trim(txtDescriAgergar.Text) & "%'"
    End If
    cSQL = cSQL & " AND P.PTO_CODIGO NOT IN ("
    cSQL = cSQL & " SELECT D.PTO_CODIGO"
    cSQL = cSQL & " FROM DETALLE_LISTA_PRECIO D"
    cSQL = cSQL & " WHERE D.LIS_CODIGO=" & cbodescri.ItemData(cbodescri.ListIndex) & ")"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    grdGrilla2.Rows = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            grdGrilla2.AddItem rec!PTO_CODIGO & Chr(9) & Trim(ChkNull(rec!PTO_CODBARRAS)) & Chr(9) & _
                               Trim(rec!PTO_DESCRI) & Chr(9) & Trim(rec!LNA_DESCRI) & Chr(9) & _
                               Valido_Importe(Chk0(rec!PTO_PREVTA))
            rec.MoveNext
        Loop
    Else
        MsgBox "No se Encontraron Productos", vbExclamation, TIT_MSGBOX
        txtDescriAgergar.Text = ""
        txtDescriAgergar.SetFocus
    End If
    rec.Close
End Sub

Private Sub cmdCancelar_Click()
    TxtDescriB.Text = ""
    CodigoProducto = ""
    TxtDescriB.Visible = False
    cbodescri.Visible = True
    cmdGrabar.Enabled = False
    CmdBuscAprox.Enabled = True
    SeteoInicial
    freOpciones.Caption = ""
    freOpciones.Caption = "Opciones de Consulta"
    Fecha1.Enabled = False
    cmdGrabar.Enabled = False
    CmdBorrar.Enabled = False
    cmdImprimir.Enabled = False
    
    cboListaPrecio.ListIndex = 0
    cboListaPrecio.Enabled = False
    
    LimpiarOpciones
    cbodescri.SetFocus
End Sub

Private Sub CmdDeselec_Click()
    For i = 1 To grdGrilla2.Rows - 1
        grdGrilla2.TextMatrix(i, 5) = "NO"
        Call CambiaColorAFilaDeGrilla(grdGrilla2, i, vbBlack, vbWhite)
    Next
    grdGrilla2.SetFocus
End Sub

Private Sub cmdfiltrar_Click()
    If txtcodigo.Text <> "" Then
        'ENTRA ACA CUANDO CONSULTA UNA LISTA
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Buscando..."
        SQL = " SELECT P.PTO_DESCRI,L.LNA_DESCRI, P.PTO_CODBARRAS,"
        SQL = SQL & " D.LIS_PRECIO,P.PTO_CODIGO, D.LIS_COSTO, M.MAR_DESCRI"
        SQL = SQL & " FROM PRODUCTO P, LINEAS L, MARCAS M,"
        SQL = SQL & " LISTA_PRECIO LP, DETALLE_LISTA_PRECIO D"
        SQL = SQL & " WHERE P.LNA_CODIGO = L.LNA_CODIGO"
        SQL = SQL & " AND D.PTO_CODIGO = P.PTO_CODIGO"
        SQL = SQL & " AND LP.LIS_CODIGO = D.LIS_CODIGO"
        SQL = SQL & " AND M.MAR_CODIGO = P.MAR_CODIGO"
        SQL = SQL & " AND LP.LIS_CODIGO = " & XN(txtcodigo.Text)
        If txtProducto.Text <> "" Then
            SQL = SQL & " AND P.PTO_DESCRI LIKE '" & Trim(txtProducto.Text) & "%' "
        End If
        If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
            SQL = SQL & " AND P.LNA_CODIGO = " & XN(cboLinea.ItemData(cboLinea.ListIndex))
        End If
        If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
            SQL = SQL & " AND P.RUB_CODIGO = " & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
        SQL = SQL & " ORDER BY P.PTO_DESCRI"
             
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            GrdModulos.Rows = 1
            Do While Not rec.EOF
                
               GrdModulos.AddItem Trim(rec!PTO_CODIGO) & Chr(9) & IIf(IsNull(rec!PTO_CODBARRAS), Trim(rec!PTO_CODIGO), Trim(rec!PTO_CODBARRAS)) _
                                   & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                                   Trim(rec!LNA_DESCRI) & Chr(9) & Trim(rec!MAR_DESCRI) & Chr(9) & Valido_Importe(Chk0(rec!LIS_PRECIO))
                                
               rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
            GrdModulos.Rows = 1
            Me.cmdfiltrar.SetFocus
        End If
    Else
        'ENTRA ACA CUANDO CARGO UNA NUEVA LISTA
        Screen.MousePointer = vbHourglass
        
        SQL = "SELECT P.PTO_DESCRI, L.LNA_DESCRI, P.PTO_CODBARRAS, M.MAR_DESCRI,"
        If cboListaPrecio.List(cboListaPrecio.ListIndex) = "(Todas)" Then
            SQL = SQL & " P.PTO_PRECTO, P.PTO_PREVTA, P.PTO_CODIGO"
            SQL = SQL & " FROM PRODUCTO P, LINEAS L, MARCAS M"
            SQL = SQL & " WHERE P.LNA_CODIGO = L.LNA_CODIGO"
            SQL = SQL & " AND M.MAR_CODIGO = P.MAR_CODIGO"
        Else
            'LIS_PRECIO
            SQL = SQL & " P.PTO_PRECTO, D.LIS_PRECIO AS PTO_PREVTA, P.PTO_CODIGO, P.PTO_CODBARRAS, M.MAR_DESCRI"
            SQL = SQL & " FROM PRODUCTO P, LINEAS L, DETALLE_LISTA_PRECIO D, MARCAS M"
            SQL = SQL & " WHERE P.LNA_CODIGO = L.LNA_CODIGO"
            SQL = SQL & " AND D.LIS_CODIGO=" & cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
            SQL = SQL & " AND P.PTO_CODIGO = D.PTO_CODIGO"
            SQL = SQL & " AND M.MAR_CODIGO = P.MAR_CODIGO"
        End If
        If txtProducto.Text <> "" Then
            SQL = SQL & " AND P.PTO_DESCRI LIKE '" & txtProducto.Text & "%' "
        End If
        If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
            SQL = SQL & " AND P.LNA_CODIGO = " & XN(cboLinea.ItemData(cboLinea.ListIndex))
        End If
        If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
            SQL = SQL & " AND P.RUB_CODIGO = " & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
        SQL = SQL & " ORDER BY P.PTO_DESCRI "
        
        lblEstado.Caption = "Buscando..."
        
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            GrdModulos.Rows = 1
            Do While Not rec.EOF
            
               GrdModulos.AddItem Trim(rec!PTO_CODIGO) & Chr(9) & IIf(IsNull(rec!PTO_CODBARRAS), Trim(rec!PTO_CODIGO), Trim(rec!PTO_CODBARRAS)) _
                                   & Chr(9) & Trim(rec!PTO_DESCRI) & Chr(9) & _
                                   Trim(rec!LNA_DESCRI) & Chr(9) & Trim(rec!MAR_DESCRI) & Chr(9) & Valido_Importe(Chk0(rec!PTO_PREVTA))
                                   
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
        Else
            lblEstado.Caption = ""
            MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
            GrdModulos.Rows = 1
            Me.cmdfiltrar.SetFocus
        End If
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Function ValidarLista()
    If TxtDescriB.Text = "" Then
        MsgBox "No ha ingresado la Descricpción de la Lista de Precios", vbExclamation, TIT_MSGBOX
        TxtDescriB.SetFocus
        ValidarLista = False
        Exit Function
    End If
    If Fecha1.Text = "" Then
        MsgBox "No ha ingresado la Fecha de Vigencia", vbExclamation, TIT_MSGBOX
        Fecha1.SetFocus
        ValidarLista = False
        Exit Function
    End If
    If GrdModulos.Rows = 1 Then
        MsgBox "Debe haber al menos un producto en la Lista de Precios", vbExclamation, TIT_MSGBOX
        ValidarLista = False
        Exit Function
    End If
    ValidarLista = True
    
End Function

Private Sub cmdGrabar_Click()
    
    On Error GoTo HayError
         
    If ValidarLista = False Then Exit Sub

    ' Entra aca cuando hago una nueva lista
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando ..."
    DBConn.BeginTrans
    
    txtcodigo = "1"
    SQL = "SELECT MAX(LIS_CODIGO) as maximo FROM LISTA_PRECIO"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Not IsNull(rec.Fields!Maximo) Then txtcodigo = XN(rec.Fields!Maximo) + 1
    rec.Close
    
    SQL = "INSERT INTO LISTA_PRECIO(LIS_CODIGO,LIS_FECHA,LIS_DESCRI)    "
    SQL = SQL & " VALUES ("
    SQL = SQL & XN(txtcodigo) & ","
    SQL = SQL & XDQ(Fecha1) & ","
    SQL = SQL & XS(TxtDescriB) & ")"
    DBConn.Execute SQL
    
    For J = 1 To GrdModulos.Rows - 1
        SQL = "INSERT INTO DETALLE_LISTA_PRECIO(LIS_CODIGO,PTO_CODIGO,LIS_PRECIO)"
        SQL = SQL & " VALUES ("
        SQL = SQL & XN(txtcodigo) & ","
        SQL = SQL & XN(GrdModulos.TextMatrix(J, 0)) & ","
        SQL = SQL & XN(GrdModulos.TextMatrix(J, 5)) & " )"
        DBConn.Execute SQL
    Next
    
    Screen.MousePointer = vbNormal
    DBConn.CommitTrans
    cmdCancelar_Click
    CmdBorrar.Enabled = True

    cbodescri.Visible = True
    TxtDescriB.Text = ""
    TxtDescriB.Visible = False
    cmdGrabar.Enabled = False
    CmdBuscAprox.Enabled = True
    freOpciones.Caption = ""
    freOpciones.Caption = "Opciones de Consulta"
    LimpiarOpciones
    Exit Sub
        
HayError:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub cmdImprimir_Click()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    If txtcodigo.Text <> "" Then
        Rep.SelectionFormula = "{LISTA_PRECIO.LIS_CODIGO}=" & XN(txtcodigo.Text)
    Else
        Exit Sub
    End If
    If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
    End If
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = "{PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
        End If
    End If
    'NO MUESTRO LOS PRODUCTOS DADOS DE BAJA
    'ESTADO 1 SON LOS PRODUCTOS HABILITADOS
    'ESTADO 2 SON LOS PRODUCTOS DADOS DE BAJA
    If Rep.SelectionFormula = "" Then
        Rep.SelectionFormula = "{PRODUCTO.PTO_ESTADO}='N'"
    Else
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.PTO_ESTADO}='N'"
    End If
    Rep.WindowTitle = "Lista de Precios..."
       
    Rep.ReportFileName = DRIVE & DirReport & "rptlistaprecio.rpt"
    
    Rep.Destination = crptToWindow
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    lblEstado.Caption = ""
End Sub

Private Sub cmdNuevo_Click()
    cmdGrabar.Enabled = True
    CmdBorrar.Enabled = False
    Fecha1.Text = Date
    CodigoProducto = ""
    txtcodigo.Text = ""
    TxtDescriB.Text = ""
    TxtDescriB.Visible = True
    cbodescri.Visible = False
    freOpciones.Caption = ""
    freOpciones.Caption = "Opciones de Carga"
    'NuevaLista 'Carga los productos de la Tabla producto
    GrdModulos.Rows = 1
    CmdBuscAprox.Enabled = False
       
    LimpiarOpciones
    
    cboListaPrecio.ListIndex = 0
    cboListaPrecio.Enabled = True

    TxtDescriB.SetFocus
End Sub

Function NuevaLista()
    GrdModulos.Rows = 1
    Screen.MousePointer = vbHourglass
    
    SQL = " SELECT P.PTO_DESCRI,L.LNA_DESCRI,R.RUB_DESCRI,"
    SQL = SQL & " RE.REP_RAZSOC,P.PTO_PRECIO,P.PTO_CODIGO "
    SQL = SQL & " FROM PRODUCTO P,LINEAS L,RUBROS R,REPRESENTADA RE, TIPO_PRESENTACION TP"
    SQL = SQL & " WHERE P.LNA_CODIGO = L.LNA_CODIGO "
    SQL = SQL & " AND P.RUB_CODIGO = R.RUB_CODIGO AND P.REP_CODIGO = RE.REP_CODIGO ORDER BY P.PTO_DESCRI"
    
    lblEstado.Caption = " Creando Nueva Lista de Precios..."
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
           GrdModulos.AddItem rec.Fields(0) & Chr(9) & rec.Fields(1) & Chr(9) & _
                              rec.Fields(2) & Chr(9) & rec.Fields(3) & Chr(9) & _
                              Valido_Importe(rec.Fields(4)) & Chr(9) & rec.Fields(5)
            rec.MoveNext
        Loop
      '  If GrdModulos.Enabled Then GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        MsgBox "No hay Productos cargados", vbOKOnly + vbCritical, TIT_MSGBOX
        Me.TxtDescriB.SetFocus
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Function

Private Sub cmdPrecios_Click()
    If GrdModulos.Rows <> 1 Then
        frmModificoPrecios.Show vbModal
        Set frmModificoPrecios = Nothing
        If txtcodigo.Text <> "" Then
            CmdBuscAprox_Click
        End If
    Else
        MsgBox "Debe haber al menos un producto en la Lista de Precios", vbExclamation, TIT_MSGBOX
    End If
End Sub

Private Sub cmdQuitar_Click()
    If GrdModulos.Rows = 1 Then
        MsgBox "Debe seleccinar un producto de la Lista", vbCritical, TIT_MSGBOX
    Else
        On Error GoTo CLAVOSE
        If MsgBox("Seguro desea quitar el Producto " & GrdModulos.TextMatrix(GrdModulos.RowSel, 2) & " de la Lista de Precios? ", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar") = vbYes Then
            If txtcodigo.Text = "" Then
                'CUANDO CARGO UNO NUEVO, SOLO ELIMINO EN LA GRILLA
                If GrdModulos.Rows = 2 Then
                    GrdModulos.Rows = 1
                Else
                    GrdModulos.RemoveItem (GrdModulos.RowSel)
                End If
            Else
                Screen.MousePointer = vbHourglass
                lblEstado.Caption = "Borrando..."
                ' CUANDO ELIMINO UN ITEM DE LA LISTA DE PRECIO YA CARGADA
                DBConn.BeginTrans
                SQL = "DELETE FROM DETALLE_LISTA_PRECIO WHERE LIS_CODIGO = " & XN(txtcodigo.Text)
                SQL = SQL & " AND PTO_CODIGO = " & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 0))
                DBConn.Execute SQL
                If GrdModulos.Rows = 2 Then
                    GrdModulos.Rows = 1
                Else
                    GrdModulos.RemoveItem (GrdModulos.RowSel)
                End If
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                DBConn.CommitTrans
            End If
        End If
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set FrmListadePrecios = Nothing
        Unload Me
    End If
End Sub

Private Sub cmdSalirFrame_Click()
    FrameBuscaProducto.Visible = False
End Sub

Private Sub cmdSalirP_Click()
    FrmListadePrecios.Enabled = True
    TabPrecios.Visible = False
    freLista.Enabled = True
    freOpciones.Enabled = True
End Sub

Private Sub CmdSelec_Click()
    For i = 1 To grdGrilla2.Rows - 1
        grdGrilla2.TextMatrix(i, 5) = "SI"
        Call CambiaColorAFilaDeGrilla(grdGrilla2, i, vbRed, vbWhite)
    Next
    grdGrilla2.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    'Call Centrar_pantalla(Me)
    Me.Left = 0
    Me.Top = 0
    SeteoInicial
    FrameBuscaProducto.Visible = False
    FrameBuscaProducto.Top = 1590
    FrameBuscaProducto.Left = 630
    FrameBuscaProducto.Height = 4185
    FrameBuscaProducto.Width = 9435
End Sub

Private Sub SeteoInicial()
    'CONFIGURO GRILLA
    GrdModulos.FormatString = "^Código Interno|^Código|<Descripción|<Linea|<Marca|>Importe"
    GrdModulos.ColWidth(0) = 0    'CODIGO INTERNO PRODUCTO
    GrdModulos.ColWidth(1) = 1350 'CODIGO DE BARRAS PRODUCTO
    GrdModulos.ColWidth(2) = 4200 'PRODUCTO - DESCRIPCION
    GrdModulos.ColWidth(3) = 2000 'LINEA
    GrdModulos.ColWidth(4) = 2000 'MARCA
    GrdModulos.ColWidth(5) = 1200 'IMPORTE
    GrdModulos.Rows = 1
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To 5
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    GrdModulos.HighLight = flexHighlightWithFocus
    
    'CONFIGURO GRILLA AGREGAR PRODUCTO
    grdGrilla2.FormatString = "^Código Interno|^Código|<Descripción|<Linea|>Importe|^Agrega"
    grdGrilla2.ColWidth(0) = 0    'CODIGO INTERNO PRODUCTO
    grdGrilla2.ColWidth(1) = 1350 'CODIGO DE BARRAS PRODUCTO
    grdGrilla2.ColWidth(2) = 3500 'PRODUCTO - DESCRIPCION
    grdGrilla2.ColWidth(3) = 2000 'LINEA
    grdGrilla2.ColWidth(4) = 1000 'IMPORTE
    grdGrilla2.ColWidth(5) = 900  'AGREGA (SI/NO)
    grdGrilla2.Rows = 1
    grdGrilla2.BorderStyle = flexBorderNone
    grdGrilla2.row = 0
    For i = 0 To 5
        grdGrilla2.Col = i
        grdGrilla2.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla2.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla2.CellFontBold = True
    Next
    grdGrilla2.HighLight = flexHighlightWithFocus
    
    cargocboLinea
    'CARGA LAS LISTA DE PRECIOS EXISTENTES
    cargocboLista
    cboListaPrecio.Enabled = False
    
    cboRubro.AddItem "(Todos)"
    cboRubro.ListIndex = 0
    
    TxtDescriB.Visible = False
    cmdGrabar.Enabled = False
    TabPrecios.Visible = False
    Fecha1.Enabled = False
    lblEstado.Caption = ""
End Sub

Private Sub cargocboLinea()
    cboLinea.Clear
    SQL = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
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
    SQL = "SELECT * FROM RUBROS"
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        SQL = SQL & " WHERE LNA_CODIGO= " & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    SQL = SQL & " ORDER BY RUB_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
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

Private Sub cargocboLista()
    cbodescri.Clear
    cboListaPrecio.Clear
    cboListaPrecio.AddItem "(Todas)"
    SQL = "SELECT LIS_CODIGO,LIS_DESCRI,LIS_FECHA "
    SQL = SQL & " FROM LISTA_PRECIO ORDER BY LIS_CODIGO DESC"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboListaPrecio.AddItem rec!LIS_DESCRI
            cboListaPrecio.ItemData(cboListaPrecio.NewIndex) = rec!LIS_CODIGO
            cbodescri.AddItem rec!LIS_DESCRI
            cbodescri.ItemData(cbodescri.NewIndex) = rec!LIS_CODIGO
            rec.MoveNext
        Loop
        cbodescri.ListIndex = 0
        cboListaPrecio.ListIndex = 0
        rec.MoveFirst
        txtcodigo.Text = rec!LIS_CODIGO
        Fecha1.Text = rec!LIS_FECHA
    Else
        cmdGrabar.Enabled = False
        CmdBorrar.Enabled = False
        cmdImprimir.Enabled = False
    End If
    rec.Close
End Sub

Private Sub grdGrilla2_DblClick()
    If Trim(grdGrilla2.TextMatrix(grdGrilla2.RowSel, 5)) = "NO" Or _
       Trim(grdGrilla2.TextMatrix(grdGrilla2.RowSel, 5)) = "" Then 'NO IMPRIME
        Call CambiaColorAFilaDeGrilla(grdGrilla2, grdGrilla2.RowSel, vbRed, vbWhite)
        grdGrilla2.TextMatrix(grdGrilla2.RowSel, 5) = "SI"
    Else
        Call CambiaColorAFilaDeGrilla(grdGrilla2, grdGrilla2.RowSel, vbBlack, vbWhite)
        grdGrilla2.TextMatrix(grdGrilla2.RowSel, 5) = "NO"
    End If
End Sub

Private Sub grdGrilla2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        grdGrilla2_DblClick
    End If
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

Private Sub GrdModulos_dblClick()
    TabPrecios.Visible = True
    txtAnterior.Text = Valido_Importe(GrdModulos.TextMatrix(GrdModulos.RowSel, 5))
    'txtCostoAnterior.Text = Valido_Importe(Chk0(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)))
    txtActual.Text = ""
    'txtCostoActual.Text = ""
    txtActual.SetFocus
    freLista.Enabled = False
    freOpciones.Enabled = False
'    frebotones.Enabled = False
End Sub

Private Sub txtActual_GotFocus()
    SelecTexto txtActual
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtActual, KeyAscii)
End Sub

Private Sub txtActual_LostFocus()
    txtActual.Text = Valido_Importe(Chk0(txtActual))
End Sub

Private Sub txtAnterior_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtAnterior, KeyAscii)
End Sub

Private Sub txtAnterior_LostFocus()
    txtAnterior.Text = Valido_Importe(txtAnterior)
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtDescriAgergar_GotFocus()
    SelecTexto txtDescriAgergar
End Sub

Private Sub txtDescriAgergar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtDescriB_GotFocus()
    SelecTexto TxtDescriB
End Sub

Private Sub TxtDescriB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtObservaciones1_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtObservaciones2_KeyPress(KeyAscii As Integer)
     KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtproducto_GotFocus()
    SelecTexto txtProducto
End Sub

Private Sub txtproducto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub LimpiarOpciones()
    cboLinea.ListIndex = 0
    txtProducto.Text = ""
    cboRubro.Clear
    cboRubro.AddItem "(Todos)"
    cboRubro.ListIndex = 0
    cboListaPrecio.ListIndex = 0
    cboListaPrecio.Enabled = False
End Sub
