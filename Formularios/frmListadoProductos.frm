VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Productos"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
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
   ScaleHeight     =   2880
   ScaleWidth      =   6555
   Begin VB.Frame FrameImpresora 
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
      TabIndex        =   11
      Top             =   1545
      Width           =   6435
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   4635
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   195
         Width           =   1665
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoProductos.frx":0000
         Left            =   450
         List            =   "frmListadoProductos.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   285
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
         Picture         =   "frmListadoProductos.frx":002F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
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
         Picture         =   "frmListadoProductos.frx":0131
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
         Index           =   2
         Left            =   135
         Picture         =   "frmListadoProductos.frx":0233
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   3885
      TabIndex        =   4
      Top             =   2355
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   5640
      TabIndex        =   6
      Top             =   2355
      Width           =   840
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoProductos.frx":0335
      Height          =   450
      Left            =   4755
      TabIndex        =   5
      Top             =   2355
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   60
      TabIndex        =   7
      Top             =   -15
      Width           =   6435
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
         Left            =   2265
         TabIndex        =   3
         Top             =   1050
         Width           =   2790
      End
      Begin VB.TextBox txtCodMarca 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   1050
         Width           =   585
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   690
         Width           =   3420
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   3420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marca:"
         Height          =   195
         Index           =   8
         Left            =   1080
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rubro:"
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linea:"
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   390
         Width           =   435
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3330
      Top             =   2460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2850
      Top             =   2370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   120
      TabIndex        =   10
      Top             =   2415
      Width           =   660
   End
End
Attribute VB_Name = "frmListadoProductos"
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

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cboLinea_Click()
    cboRubro.Clear
End Sub

Private Sub cboLinea_LostFocus()
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        cargocboRubro
    Else
        cboRubro.Clear
        cboRubro.AddItem "(Todos)"
        cboRubro.ListIndex = 0
    End If
End Sub

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
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
    
    Rep.SelectionFormula = ""
    If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
        Rep.SelectionFormula = " {PRODUCTO.LNA_CODIGO}=" & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PRODUCTO.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {RUBROS.RUB_CODIGO}=" & XN(cboRubro.ItemData(cboRubro.ListIndex))
        End If
    End If
    
    If txtCodMarca.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {PRODUCTO.MAR_CODIGO}=" & XN(txtCodMarca.Text)
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.MAR_CODIGO}=" & XN(txtCodMarca.Text)
        End If
    End If
    'MUESTRO LOS PRODUCTOS NO DADOS DE BAJA
    If Rep.SelectionFormula = "" Then
        Rep.SelectionFormula = " {PRODUCTO.PTO_ESTADO}='N'"
    Else
        Rep.SelectionFormula = Rep.SelectionFormula & " AND {PRODUCTO.PTO_ESTADO}='N'"
    End If
    
    Rep.WindowTitle = "Listado de Productos"
    Rep.ReportFileName = DRIVE & DirReport & "listadoproductos.rpt"
    Rep.Action = 1
    
    lblEstado.Caption = ""
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
End Sub

Private Sub cmdNuevo_Click()
    cboLinea.ListIndex = 0
    cboRubro.Clear
    cboRubro.AddItem "(Todos)"
    cboRubro.ListIndex = 0
    cboLinea.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmListadoProductos = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    'Call Centrar_pantalla(Me)
    Me.Top = 0
    Me.Left = 0
    cargocboLinea
    cboRubro.AddItem "(Todos)"
    cboRubro.ListIndex = 0
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    cboDestino.ListIndex = 0
End Sub

Private Sub cargocboLinea()
    lblEstado.Caption = ""
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
    SQL = "SELECT * FROM RUBROS "
    SQL = SQL & " WHERE LNA_CODIGO= " & XN(cboLinea.ItemData(cboLinea.ListIndex))
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
    End If
    rec.Close
End Sub

Private Sub txtCodMarca_Change()
    If txtCodMarca.Text = "" Then
        txtDescriMarca.Text = ""
    End If
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


