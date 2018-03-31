VERSION 5.00
Begin VB.Form frmAjusteStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste de Stock - Productos"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
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
   ScaleHeight     =   2640
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   4950
      TabIndex        =   11
      Top             =   2130
      Width           =   960
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   450
      Left            =   3000
      TabIndex        =   4
      Top             =   2130
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   45
      TabIndex        =   6
      Top             =   30
      Width           =   5880
      Begin VB.TextBox txtCodInt 
         Height          =   345
         Left            =   4785
         TabIndex        =   14
         Top             =   930
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txtdescri 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         Top             =   555
         Width           =   4320
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txtStockFisicoReal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   3
         Top             =   1575
         Width           =   1170
      End
      Begin VB.TextBox txtStockFisicoSis 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         TabIndex        =   2
         Top             =   1245
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   13
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Producto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1455
         TabIndex        =   12
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Stock Fisico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   630
         TabIndex        =   10
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Real:"
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
         Left            =   930
         TabIndex        =   9
         Top             =   1620
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sistema:"
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
         Left            =   630
         TabIndex        =   8
         Top             =   1305
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   3975
      TabIndex        =   5
      Top             =   2130
      Width           =   960
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
      Left            =   90
      TabIndex        =   7
      Top             =   2220
      Width           =   660
   End
End
Attribute VB_Name = "frmAjusteStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrabar_Click()
    If txtcodigo.Text = "" Then
        MsgBox "Falta Ingresar el Producto", vbCritical, TIT_MSGBOX
        txtcodigo.SetFocus
        Exit Sub
    End If
    If txtStockFisicoReal.Text = "" Then
        MsgBox "El Stock Real no puede estar en blanco", vbCritical, TIT_MSGBOX
        txtStockFisicoReal.SetFocus
        Exit Sub
    End If
    If MsgBox("Confirma ajuste de Stock", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        lblEstado.Caption = "Actualizando..."
        SQL = "UPDATE STOCK"
        SQL = SQL & " SET DST_STKFIS=" & XN(txtStockFisicoReal.Text)
        SQL = SQL & " WHERE STK_CODIGO=" & XN(Sucursal)
        SQL = SQL & " AND PTO_CODIGO=" & XN(txtCodInt.Text)
        DBConn.Execute SQL
        lblEstado.Caption = ""
        cmdNuevo_Click
    End If
End Sub

Private Sub cmdNuevo_Click()
    txtcodigo.Text = ""
    txtStockFisicoReal.Text = ""
    txtStockFisicoSis.Text = ""
    txtcodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Set frmAjusteStock = Nothing
    Unload Me
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
    SQL = "SELECT SUC_CODIGO, SUC_DESCRI "
    SQL = SQL & " FROM SUCURSAL R "
    SQL = SQL & " WHERE SUC_CODIGO = " & XN(Sucursal)
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Frame1.Caption = "Ajuste de Stock Sucursal  - " & Trim(rec!SUC_DESCRI)
    End If
    rec.Close
    lblEstado.Caption = ""
End Sub

Private Sub TxtCodigo_Change()
    If txtcodigo.Text = "" Then
        txtcodigo.Text = ""
        txtDescri.Text = ""
        txtCodInt.Text = ""
        txtStockFisicoReal.Text = ""
        txtStockFisicoSis.Text = ""
        cmdGrabar.Enabled = False
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        txtcodigo.SetFocus
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = " SELECT P.PTO_DESCRI, P.PTO_CODIGO, S.DST_STKFIS"
        SQL = SQL & " FROM PRODUCTO P, STOCK S"
        SQL = SQL & " WHERE"
        SQL = SQL & " P.PTO_CODIGO=S.PTO_CODIGO"
        SQL = SQL & " AND S.STK_CODIGO=" & XN(Sucursal)
        If IsNumeric(txtcodigo.Text) Then
            SQL = SQL & " AND P.PTO_CODIGO =" & XN(txtcodigo.Text) & " OR P.PTO_CODBARRAS=" & XS(txtcodigo.Text)
        Else
            SQL = SQL & " AND P.PTO_CODBARRAS=" & XS(txtcodigo.Text)
        End If
        SQL = SQL & " ORDER BY P.PTO_CODIGO"
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDescri.Text = Trim(rec!PTO_DESCRI)
            txtCodInt.Text = rec!PTO_CODIGO
            txtStockFisicoSis.Text = Chk0(rec!DST_STKFIS)
            cmdGrabar.Enabled = True
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            txtcodigo.SetFocus
            cmdGrabar.Enabled = False
        End If
        rec.Close
    End If
End Sub

Private Sub txtdescri_Change()
    If txtDescri.Text = "" Then
        txtcodigo.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtDescri
End Sub

Private Sub txtdescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        txtDescri.SetFocus
    End If
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
   If txtcodigo.Text = "" And txtDescri.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        SQL = "SELECT P.PTO_CODIGO,P.PTO_DESCRI,P.PTO_CODBARRAS,S.DST_STKFIS"
        SQL = SQL & " FROM PRODUCTO P, STOCK S"
        SQL = SQL & " WHERE P.PTO_DESCRI LIKE '" & txtDescri.Text & "%'"
        SQL = SQL & " AND P.PTO_CODIGO=S.PTO_CODIGO"
        SQL = SQL & " AND S.STK_CODIGO=" & XN(Sucursal)
        Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1.RecordCount > 1 Then
                'grdGrilla.SetFocus
                BuscarProducto "CADENA", Trim(txtDescri.Text)
                txtDescri.SetFocus
            Else
                txtcodigo.Text = Trim(Rec1!PTO_CODBARRAS)
                txtDescri.Text = Trim(Rec1!PTO_DESCRI)
                txtCodInt.Text = Trim(Rec1!PTO_CODIGO)
                txtStockFisicoSis.Text = Chk0(Rec1!DST_STKFIS)
                cmdGrabar.Enabled = True
            End If
        Else
                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                txtDescri.Text = ""
        End If
        Rec1.Close
        Screen.MousePointer = vbNormal
    ElseIf txtcodigo.Text = "" And txtDescri.Text = "" Then
        cmdGrabar.Enabled = False
    End If
End Sub

Private Sub txtStockFisicoReal_GotFocus()
    SelecTexto txtStockFisicoReal
End Sub

Private Sub txtStockFisicoReal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Public Sub BuscarProducto(mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        'Set .Conn = DBConn
        cSQL = "SELECT PTO_DESCRI, PTO_CODIGO"
        cSQL = cSQL & " FROM PRODUCTO"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE"
            cSQL = cSQL & " PTO_DESCRI LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Descripción, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "PTO_DESCRI"
        campo1 = .Field
        .Field = "PTO_CODIGO"
        campo2 = .Field
        .OrderBy = "PTO_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Productos :"
        .MaxRecords = 1
        .Show
        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
                txtcodigo.Text = .ResultFields(2)
                TxtCodigo_LostFocus
        End If
    End With
    Set B = Nothing
End Sub

