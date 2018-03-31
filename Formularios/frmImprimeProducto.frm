VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmImprimeProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Productos"
   ClientHeight    =   2640
   ClientLeft      =   1515
   ClientTop       =   1740
   ClientWidth     =   5145
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImprimeProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2520
      TabIndex        =   5
      Top             =   2190
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3825
      TabIndex        =   6
      Top             =   2190
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   45
      TabIndex        =   24
      Top             =   30
      Width           =   5055
      Begin VB.TextBox txtCodMarca 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1110
         Width           =   585
      End
      Begin VB.TextBox txtDescriMarca 
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
         Left            =   1815
         TabIndex        =   3
         Top             =   1110
         Width           =   2745
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         ItemData        =   "frmImprimeProducto.frx":0442
         Left            =   1200
         List            =   "frmImprimeProducto.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   765
         Width           =   3375
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         ItemData        =   "frmImprimeProducto.frx":0446
         Left            =   1200
         List            =   "frmImprimeProducto.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Index           =   8
         Left            =   615
         TabIndex        =   27
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Index           =   0
         Left            =   615
         TabIndex        =   26
         Top             =   795
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Index           =   0
         Left            =   615
         TabIndex        =   25
         Top             =   450
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   6735
      TabIndex        =   14
      Top             =   210
      Visible         =   0   'False
      Width           =   6915
      Begin VB.TextBox txtEmpresaCuit 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   22
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox txtEmp_Id 
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
         Left            =   4365
         TabIndex        =   21
         Top             =   1065
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   75
         Left            =   90
         TabIndex        =   20
         Top             =   1560
         Width           =   6795
      End
      Begin VB.TextBox txtTipoLibro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4365
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEmpresa 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   18
         Top             =   375
         Width           =   3075
      End
      Begin VB.TextBox txtMes_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         MaxLength       =   2
         TabIndex        =   7
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtAnio_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   8
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txtLibro_IdI 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4380
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "C.U.I.T."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1005
         Width           =   540
      End
   End
   Begin VB.Frame fraImpresion 
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
      Left            =   45
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
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
         Picture         =   "frmImprimeProducto.frx":044A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
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
         Picture         =   "frmImprimeProducto.frx":054C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
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
         Index           =   0
         Left            =   135
         Picture         =   "frmImprimeProducto.frx":064E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmImprimeProducto.frx":0750
         Left            =   450
         List            =   "frmImprimeProducto.frx":075D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   1635
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2340
      Top             =   1830
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Modo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6690
      TabIndex        =   10
      Top             =   2985
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmImprimeProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cboListar_Click()
    If frmImprimeProducto.Visible = True Then
        If cboListar.ListIndex = 0 Then
            cboAgrupar.Enabled = True
            cboAgrupar.ListIndex = 0
        Else
            cboAgrupar.Enabled = False
            cboAgrupar.ListIndex = -1
        End If
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

Private Sub cmdAceptar_Click()
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
    '------------------------
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
        End If
    End If
    If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
    End If
    If txtCodMarca.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PRODUCTO.MAR_CODIGO}=" & XN(txtCodMarca.Text)
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.MAR_CODIGO}=" & XN(txtCodMarca.Text)
        End If
    End If
    Rep.WindowTitle = "Listado de Productos"
    Rep.ReportFileName = DRIVE & DirReport & "listadoproductos.rpt"
    Rep.Action = 1
End Sub

Private Sub cmdCancelar_Click()
    mQuienLlamo = ""
    Set frmImprimeProducto = Nothing
    'mQuienLlamo = "ABMProducto"
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    'CARGO COMBO LINEA
    cargocboLinea
    cboLinea_LostFocus
    cboDestino.ListIndex = 0
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

Private Sub Form_Unload(Cancel As Integer)
    mQuienLlamo = ""
End Sub

Private Sub txtCodMarca_Change()
    If txtCodMarca.Text = "" Then
        txtDescriMarca.Text = ""
    End If
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCodMarca_GotFocus()
    SelecTexto txtCodMarca
End Sub

Private Sub txtCodMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarMarcas txtCodMarca, "CODIGO"
    End If
End Sub

Private Sub txtCodMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodMarca_LostFocus()
    If txtCodMarca.Text <> "" Then
        SQL = "SELECT MAR_CODIGO, MAR_DESCRI"
        SQL = SQL & " FROM MARCAS"
        SQL = SQL & " WHERE MAR_CODIGO =" & XN(txtCodMarca.Text)
        If rec.State = 1 Then rec.Close
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDescriMarca.Text = ChkNull(rec!MAR_DESCRI)
        Else
            MsgBox "El Código no existe", vbInformation
            txtDescriMarca.Text = ""
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDescriMarca_Change()
    If txtDescriMarca.Text = "" Then
        txtCodMarca.Text = ""
    End If
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDescriMarca_GotFocus()
    SelecTexto txtDescriMarca
End Sub

Private Sub txtDescriMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarMarcas txtCodMarca, "CODIGO"
    End If
End Sub

Private Sub txtDescriMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescriMarca_LostFocus()
    If txtCodMarca.Text = "" And txtDescriMarca.Text <> "" Then
        SQL = "SELECT MAR_CODIGO, MAR_DESCRI"
        SQL = SQL & " FROM MARCAS"
        SQL = SQL & " WHERE MAR_DESCRI LIKE '" & XN(Trim(txtDescriMarca.Text)) & "%'"
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarMarcas txtCodMarca, "CADENA", Trim(txtDescriMarca.Text)
                If rec.State = 1 Then rec.Close
                txtDescriMarca.SetFocus
            Else
                txtCodMarca.Text = rec!MAR_CODIGO
                txtDescriMarca.Text = rec!MAR_DESCRI
            End If
        Else
            MsgBox "La Marca no existe", vbExclamation, TIT_MSGBOX
            txtCodMarca.Text = ""
            txtDescriMarca.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Public Sub BuscarMarcas(Txt As Control, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT MAR_DESCRI, MAR_CODIGO"
        cSQL = cSQL & " FROM MARCAS"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE MAR_DESCRI LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Descripción, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "MAR_DESCRI"
        campo1 = .Field
        .Field = "MAR_CODIGO"
        campo2 = .Field
        .OrderBy = "MAR_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Marcas :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCodMarca.Text = .ResultFields(2)
            txtCodMarca_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

