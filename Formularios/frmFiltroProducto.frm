VERSION 5.00
Begin VB.Form frmFiltroProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro Búsqueda Producto"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFiltroProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLinea 
      Height          =   315
      ItemData        =   "frmFiltroProducto.frx":27A2
      Left            =   870
      List            =   "frmFiltroProducto.frx":27A4
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   495
      Width           =   3375
   End
   Begin VB.ComboBox cboRubro 
      Height          =   315
      ItemData        =   "frmFiltroProducto.frx":27A6
      Left            =   870
      List            =   "frmFiltroProducto.frx":27A8
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtDescriMarca 
      Height          =   315
      Left            =   1485
      TabIndex        =   3
      Top             =   1185
      Width           =   2745
   End
   Begin VB.TextBox txtCodMarca 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   870
      TabIndex        =   2
      Top             =   1185
      Width           =   585
   End
   Begin VB.CommandButton cbmCerrarFiltro 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3135
      TabIndex        =   6
      Top             =   1980
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptarFiltro 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1995
      TabIndex        =   5
      Top             =   1980
      Width           =   1110
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   345
      Left            =   870
      TabIndex        =   4
      Top             =   1530
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descri.::"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Linea:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   885
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   8
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Filtro de Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   105
      Width           =   3690
   End
End
Attribute VB_Name = "frmFiltroProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbmCerrarFiltro_Click()
    Unload Me
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

Private Sub cmdAceptarFiltro_Click()
    Dim auxListView As ListView
    Screen.MousePointer = vbHourglass
    With auxDllActiva
        Set auxListView = .FormBase.lstvLista
        If txtBusqueda.Text <> "" Or cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Or _
           cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Or txtCodMarca.Text <> "" Then
            
            If .Caption = "Actualización de Productos" Then
                .SQL = "SELECT P.PTO_DESCRI, P.PTO_CODIGO, R.RUB_DESCRI, L.LNA_DESCRI, M.MAR_DESCRI" & _
                       " FROM PRODUCTO P, RUBROS R, LINEAS L, MARCAS M" & _
                       " WHERE R.LNA_CODIGO=L.LNA_CODIGO AND P.LNA_CODIGO=L.LNA_CODIGO" & _
                       " AND P.RUB_CODIGO=R.RUB_CODIGO" & _
                       " AND M.MAR_CODIGO=P.MAR_CODIGO"
                .SQL = .SQL & " AND P.PTO_DESCRI like " & XS(txtBusqueda.Text & "%")
                If cboLinea.List(cboLinea.ListIndex) <> "(Todas)" Then
                    .SQL = .SQL & " AND P.LNA_CODIGO=" & cboLinea.ItemData(cboLinea.ListIndex)
                End If
                If cboRubro.List(cboRubro.ListIndex) <> "(Todos)" Then
                    .SQL = .SQL & " AND P.LNA_CODIGO=" & cboRubro.ItemData(cboRubro.ListIndex)
                End If
                If txtCodMarca.Text <> "" Then
                    .SQL = .SQL & " AND P.MAR_CODIGO=" & XN(txtCodMarca.Text)
                End If
            End If
        Else
            If .Caption = "Actualización de Productos" Then
                .SQL = "SELECT P.PTO_DESCRI, P.PTO_CODIGO, R.RUB_DESCRI, L.LNA_DESCRI, M.MAR_DESCRI" & _
                       " FROM PRODUCTO P, RUBROS R, LINEAS L, MARCAS M" & _
                       " WHERE R.LNA_CODIGO=L.LNA_CODIGO AND P.LNA_CODIGO=L.LNA_CODIGO" & _
                       " AND P.RUB_CODIGO=R.RUB_CODIGO" & _
                       " AND M.MAR_CODIGO=P.MAR_CODIGO"
            End If
        End If
        CargarListView .FormBase, auxListView, .SQL, .FieldID, .HeaderSQL, .FormBase.ImgLstLista
        .FormBase.sBarEstado.Panels(1).Text = auxListView.ListItems.Count & " Registro(s)"
    End With
    Screen.MousePointer = vbDefault
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    cmdAceptarFiltro_Click
    'End If
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    cargocboLinea
    cboLinea_LostFocus
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

Private Sub txtBusqueda_GotFocus()
    SelecTexto txtBusqueda
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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

