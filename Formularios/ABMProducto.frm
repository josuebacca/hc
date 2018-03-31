VERSION 5.00
Begin VB.Form ABMProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Producto..."
   ClientHeight    =   3750
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
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
      Left            =   1665
      TabIndex        =   5
      Top             =   1545
      Width           =   2745
   End
   Begin VB.TextBox txtCodMarca 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Top             =   1545
      Width           =   585
   End
   Begin VB.TextBox txtCodBarras 
      Height          =   300
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2580
      Width           =   3375
   End
   Begin VB.TextBox TxtPrecioVta 
      Height          =   300
      Left            =   3495
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1890
      Width           =   930
   End
   Begin VB.TextBox TxtPrecioCto 
      Height          =   300
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1890
      Width           =   930
   End
   Begin VB.CheckBox chkPtoEstado 
      Caption         =   "Dar de Baja"
      Height          =   285
      Left            =   1065
      TabIndex        =   10
      Top             =   2955
      Width           =   1140
   End
   Begin VB.TextBox txtStockMinimo 
      Height          =   300
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2235
      Width           =   930
   End
   Begin VB.ComboBox cboRubro 
      Height          =   315
      ItemData        =   "ABMProducto.frx":000C
      Left            =   1050
      List            =   "ABMProducto.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ComboBox cboLinea 
      Height          =   315
      ItemData        =   "ABMProducto.frx":0010
      Left            =   1050
      List            =   "ABMProducto.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMProducto.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3345
      Width           =   330
   End
   Begin VB.TextBox txtDescri 
      Height          =   300
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   1
      Top             =   495
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   1050
      TabIndex        =   0
      Top             =   150
      Width           =   840
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   12
      Top             =   3345
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   11
      Top             =   3345
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      Height          =   195
      Index           =   8
      Left            =   60
      TabIndex        =   22
      Top             =   1575
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cód. Barras:"
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   21
      Top             =   2625
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Precio Vta:"
      Height          =   195
      Index           =   6
      Left            =   2565
      TabIndex        =   20
      Top             =   1935
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Precio Cto:"
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   19
      Top             =   1935
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Stock Min.:"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   18
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Línea:"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   17
      Top             =   900
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rubro:"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   16
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   540
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   180
      Width           =   270
   End
End
Attribute VB_Name = "ABMProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Dim vFieldID As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim mLinea As String

'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "PRODUCTO"
Const cCampoID = "PTO_CODIGO"
Const cDesRegistro = "Producto"

Function ActualizarListaBase(pMode As Integer)
    On Error GoTo moco
    Dim rec As ADODB.Recordset
    Dim cSQL As String
    Dim i As Integer
    Dim auxListItem As ListItem
    Dim IndiceCampoID As Integer
    Dim OrdenCampo As Integer
    Dim f As ADODB.Field
    Set rec = New ADODB.Recordset
    
    'armo la cadena a ejecutar
    If InStr(1, vStringSQL, "WHERE") = 0 Then
        cSQL = vStringSQL & " WHERE " & cCampoID & " = " & txtID.Text
    Else
        cSQL = vStringSQL & " AND " & cCampoID & " = " & txtID.Text
    End If
    
    If pMode = 4 Then
        vListView.ListItems.Remove vListView.SelectedItem.Index
        Exit Function
    End If
    
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
        If rec.EOF = False Then
        
'            'busco el indce del campo identificador
            OrdenCampo = 0
            IndiceCampoID = 0
            For Each f In rec.Fields
                OrdenCampo = OrdenCampo + 1
                If UCase(f.Name) = UCase(vDesFieldID) Then
                    IndiceCampoID = OrdenCampo - 1
                End If
            Next f
        
            'recorro la coleción de campos a actualizar
            For i = 0 To rec.Fields.Count - 1
                If i = 0 Then
                    Select Case pMode
                        Case 1
                            Set auxListItem = vListView.ListItems.Add(, "'" & rec.Fields(IndiceCampoID) & "'", CStr(IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))), 1)
                            auxListItem.Icon = 1
                            auxListItem.SmallIcon = 1
                            
                        Case 2
                            Set auxListItem = vListView.SelectedItem
                            auxListItem.Text = rec.Fields(i)
                    End Select
                Else
                    auxListItem.SubItems(i) = IIf(IsNull(rec.Fields(i)), "", rec.Fields(i))
                End If
            Next i
        End If
    End If
    Exit Function
moco:
    If Err.Number = 35613 Then
        Call Menu.mnuContextABM_Click(4)
    End If
End Function

Function SetMode(pMode As Integer)

    'Configura los controles del form segun el parametro pMode
    'Parametro: pMode indica el modo en que se utilizará este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    
    Select Case pMode
        Case 1, 2
            AcCtrl txtID
            AcCtrl txtDescri
            AcCtrl cboLinea
            AcCtrl cboRubro
            AcCtrl txtCodMarca
            AcCtrl txtDescriMarca
            AcCtrl txtStockMinimo
            AcCtrl chkPtoEstado
            AcCtrl TxtPrecioCto
            AcCtrl TxtPrecioVta
            AcCtrl txtCodBarras
        Case 3, 4
            DesacCtrl txtID
            DesacCtrl txtDescri
            DesacCtrl cboLinea
            DesacCtrl cboRubro
            DesacCtrl txtCodMarca
            DesacCtrl txtDescriMarca
            DesacCtrl txtStockMinimo
            DesacCtrl chkPtoEstado
            DesacCtrl TxtPrecioCto
            DesacCtrl TxtPrecioVta
            DesacCtrl txtCodBarras
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo Producto.."
            txtID_LostFocus
            DesacCtrl txtID
            
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Producto..."
            DesacCtrl txtID

        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del Producto..."
            DesacCtrl txtID
            
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Producto..."
            DesacCtrl txtID
    End Select
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizará la ventana de datos
    Set vListView = pListview 'objeto listview que se está editando
    vDesFieldID = pDesID 'nombre del campo identificador
    
    'valor del campo identificador de registro seleccionado (0 si es un reg. nuevo)
    If vMode <> 1 Then
        If vListView.SelectedItem.Selected = True Then
            vFieldID = vListView.SelectedItem.Key
        Else
            vFieldID = 0
        End If
    Else
        vFieldID = 0
    End If

End Function


Function Validar(pMode As Integer) As Boolean

    If pMode <> 4 Then
        Validar = False
        If txtID.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Identificación del Producto antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
            
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la descripción del Producto antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
            
        ElseIf cboLinea.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Linea del Producto antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboLinea.SetFocus
            Exit Function
            
        ElseIf cboRubro.ListCount = 0 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Rubro del Producto antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboRubro.SetFocus
            Exit Function
        
        ElseIf txtCodMarca.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Marca del Producto antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtCodMarca.SetFocus
            Exit Function
            
        ElseIf txtCodBarras.Text <> "" Then
            SQL = "SELECT PTO_DESCRI FROM PRODUCTO"
            SQL = SQL & " WHERE PTO_CODBARRAS=" & XS(txtCodBarras.Text)
            If pMode = 2 Then
                SQL = SQL & " AND PTO_CODIGO<>" & XN(txtID.Text)
            End If
            If rec.State = 1 Then rec.Close
            rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Beep
                MsgBox "Código de Barras Existente." & Chr(13) & _
                                 "El Mismo fue ingresado en el Producto: " & Trim(rec!PTO_DESCRI), vbCritical + vbOKOnly, App.Title
                txtCodBarras.SetFocus
                rec.Close
                Exit Function
            End If
            If rec.State = 1 Then rec.Close
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboLinea_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboLinea_LostFocus()
    If mLinea = "" Or (mLinea <> "" And CInt(Chk0(mLinea)) <> cboLinea.ItemData(cboLinea.ListIndex)) Then
        Call CargoComboRubros(mLinea)
        mLinea = ""
    End If
End Sub

Private Sub CargoComboRubros(linea As String)
    If linea = "" Or (linea <> "" And CInt(Chk0(linea)) <> cboLinea.ItemData(cboLinea.ListIndex)) Then
        Set Rec1 = New ADODB.Recordset
        cboRubro.Clear
        SQL = "SELECT RUB_CODIGO,RUB_DESCRI"
        SQL = SQL & " FROM RUBROS"
        SQL = SQL & " WHERE LNA_CODIGO=" & cboLinea.ItemData(cboLinea.ListIndex)
        SQL = SQL & " ORDER BY RUB_DESCRI"
        
        Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If (Rec1.BOF And Rec1.EOF) = 0 Then
           Do While Rec1.EOF = False
              cboRubro.AddItem Trim(Rec1!RUB_DESCRI)
              cboRubro.ItemData(cboRubro.NewIndex) = Rec1!RUB_CODIGO
              Rec1.MoveNext
           Loop
           cboRubro.ListIndex = cboRubro.ListIndex + 1
        Else
           MsgBox "No hay cargado Tipos para ese Rubro.", vbOKOnly + vbCritical, TIT_MSGBOX
        End If
        Rec1.Close
    End If
End Sub

Private Sub cboRubro_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub chkPtoEstado_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'Nuevo
                'Insert en Productos
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "  (PTO_CODIGO, PTO_DESCRI, LNA_CODIGO, RUB_CODIGO, MAR_CODIGO,"
                cSQL = cSQL & " PTO_STKMIN, PTO_ESTADO, PTO_PRECTO, PTO_PREVTA, PTO_CODBARRAS)"
                cSQL = cSQL & "VALUES ("
                cSQL = cSQL & XN(txtID.Text) & ", " & XS(txtDescri.Text) & ", "
                cSQL = cSQL & cboLinea.ItemData(cboLinea.ListIndex) & ", "
                cSQL = cSQL & cboRubro.ItemData(cboRubro.ListIndex) & ", "
                cSQL = cSQL & XN(txtCodMarca.Text) & ", "
                cSQL = cSQL & XN(txtStockMinimo.Text) & ", "
                If chkPtoEstado.Value = Checked Then
                    cSQL = cSQL & "'S',"
                Else
                    cSQL = cSQL & "'N',"
                End If
                cSQL = cSQL & XN(TxtPrecioCto.Text) & ", "
                cSQL = cSQL & XN(TxtPrecioVta.Text) & ", "
                cSQL = cSQL & XS(txtCodBarras.Text) & ")"
                DBConn.Execute cSQL
                
                'Insert en Stock
                cSQL = "INSERT INTO STOCK"
                cSQL = cSQL & "  (STK_CODIGO, PTO_CODIGO, DST_STKFIS) "
                cSQL = cSQL & "VALUES (" & XN(Sucursal) & " , " & XN(txtID.Text) & " ,0)"
                DBConn.Execute cSQL
                
                'Insert en Lista de Precios
                'Busco el Max Nro de Lista de Precios
                'Inserto el Producto con el Max Nro de Lista de Precio
'                cSQL = "SELECT MAX(LIS_CODIGO)AS MaxNroLista " & _
'                       "  FROM LISTA_PRECIO"
'                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
'                If (rec.BOF And rec.EOF) = 0 Then
'                   Do While rec.EOF = False
'                      cSQL = "INSERT INTO DETALLE_LISTA_PRECIO"
'                      cSQL = cSQL & "(LIS_CODIGO, PTO_CODIGO, LIS_PRECIO, LIS_COSTO) "
'                      cSQL = cSQL & " VALUES (" & XN(rec!MaxNroLista) & ", "
'                      cSQL = cSQL & XN(txtID.Text) & ","
'                      cSQL = cSQL & XN(TxtPrecioVta.Text) & ", "
'                      cSQL = cSQL & XN(TxtPrecioCto.Text) & ") "
'                      DBConn.Execute cSQL
'                      rec.MoveNext
'                   Loop
'                   If cboLinea.ListCount > 0 Then cboLinea.ListIndex = 0
'                End If
'                rec.Close
                
            Case 2 'Editar
                'UPDATE en Productos
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & " PTO_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & " ,LNA_CODIGO = " & cboLinea.ItemData(cboLinea.ListIndex)
                cSQL = cSQL & " ,RUB_CODIGO = " & cboRubro.ItemData(cboRubro.ListIndex)
                cSQL = cSQL & " ,MAR_CODIGO = " & XN(txtCodMarca.Text)
                cSQL = cSQL & " ,PTO_STKMIN = " & XN(txtStockMinimo.Text)
                If chkPtoEstado.Value = Checked Then
                    cSQL = cSQL & " ,PTO_ESTADO = 'S'"
                Else
                    cSQL = cSQL & " ,PTO_ESTADO = 'N'"
                End If
                cSQL = cSQL & " ,PTO_PRECTO = " & XN(TxtPrecioCto.Text)
                cSQL = cSQL & " ,PTO_PREVTA = " & XN(TxtPrecioVta.Text)
                cSQL = cSQL & " ,PTO_CODBARRAS = " & XS(txtCodBarras.Text)
                cSQL = cSQL & " WHERE PTO_CODIGO  = " & XN(txtID.Text)
                DBConn.Execute cSQL
                
                'UPDATE en Stock
                'NO hay que actualizaar NADA en Stock
                
                'UPDATE en Lista de Precios
'                cSQL = "UPDATE DETALLE_LISTA_PRECIO"
'                cSQL = cSQL & "   SET LIS_PRECIO = " & XN(TxtPrecioVta.Text) & _
'                              ",      LIS_COSTO  = " & XN(TxtPrecioCto.Text) & _
'                              " WHERE PTO_CODIGO = " & XN(txtID.Text)
'                DBConn.Execute cSQL
                
            Case 4 'eliminar
                
                'DELETE en Lista de Precios
                cSQL = "DELETE FROM DETALLE_LISTA_PRECIO " & _
                       " WHERE PTO_CODIGO  = " & XN(txtID.Text)
                DBConn.Execute cSQL
                
                'DELETE en Stock
                cSQL = "DELETE FROM STOCK " & _
                       " WHERE PTO_CODIGO  = " & XN(txtID.Text) & _
                       " AND STK_CODIGO    = " & XN(Sucursal)
                DBConn.Execute cSQL
                
                'DELETE en Productos
                cSQL = "DELETE FROM " & cTabla & _
                       " WHERE PTO_CODIGO  = " & XN(txtID.Text)
                DBConn.Execute cSQL
                                           
        End Select
        
        
        DBConn.CommitTrans
        'On Error GoTo 0
        
        'actualizo la lista base
        ActualizarListaBase vMode
        
        Screen.MousePointer = vbDefault
        Unload Me
    End If
    Exit Sub
    
ErrorTran:
    
    DBConn.RollbackTrans
    Screen.MousePointer = vbDefault
    
    'manejo el error
    'ManejoDeErrores DBConn.ErrorNative
    MsgBox Err.Description, vbCritical
    
End Sub


Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 12)
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'hizo click en una columna no correcta
    If vMode = 2 And vFieldID = "0" Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    
    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim cSQL As String
    Dim hSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    
    'txtID.MaxLength = 4
    'txtDescri.MaxLength = 30
    'cargo el combo de PAIS
    cboLinea.Clear
    cSQL = "SELECT LNA_CODIGO, LNA_DESCRI FROM LINEAS ORDER BY LNA_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboLinea.AddItem Trim(rec!LNA_DESCRI)
          cboLinea.ItemData(cboLinea.NewIndex) = rec!LNA_CODIGO
          rec.MoveNext
       Loop
       If cboLinea.ListCount > 0 Then cboLinea.ListIndex = 0
    End If
    rec.Close
    mLinea = ""
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE PTO_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!PTO_CODIGO
                txtDescri.Text = rec!PTO_DESCRI
                mLinea = rec!LNA_CODIGO
                Call BuscaCodigoProxItemData(CInt(rec!LNA_CODIGO), cboLinea)
                'cboLinea_LostFocus
                Call CargoComboRubros("")
                Call BuscaCodigoProxItemData(CInt(rec!RUB_CODIGO), cboRubro)
                txtCodMarca.Text = ChkNull(rec!MAR_CODIGO)
                txtCodMarca_LostFocus
                txtStockMinimo.Text = ChkNull(rec!PTO_STKMIN)
                TxtPrecioCto.Text = Valido_Importe(Chk0(rec!PTO_PRECTO))
                TxtPrecioVta.Text = Valido_Importe(Chk0(rec!PTO_PREVTA))
                txtCodBarras.Text = ChkNull(rec!PTO_CODBARRAS)
                If ChkNull(rec!PTO_ESTADO) = "N" Or ChkNull(rec!PTO_ESTADO) = "" Then
                    chkPtoEstado.Value = Unchecked
                Else
                    chkPtoEstado.Value = Checked
                End If
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub

Private Sub txtCodBarras_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCodBarras_GotFocus()
    SelecTexto txtCodBarras
End Sub

Private Sub txtCodBarras_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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

Private Sub txtdescri_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtdescri_GotFocus()
    seltxt
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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

Private Sub txtID_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtID_GotFocus()
    seltxt
End Sub

Private Sub txtID_LostFocus()

    Dim cSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    If vMode = 1 Then ' si se esta usando en modo de nuevo registro
        If txtID.Text = "" Then
            If cSugerirID = True Then
                cSQL = "SELECT MAX(" & cCampoID & ") FROM " & cTabla
                'cSQL = cSQL & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                'cSQL = cSQL & " AND PRO_CODIGO = " & cboProvincia.ItemData(cboProvincia.ListIndex)
                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (rec.BOF And rec.EOF) = 0 Then
                    If rec.Fields(0) > 0 Then
                        txtID.Text = rec.Fields(0) + 1
                    Else
                        txtID.Text = 1
                    End If
                End If
            End If
        Else
            'verifico que no sea clave repetida
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & XN(txtID.Text)
            'cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                If rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "Código de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El código ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub TxtPrecioCto_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub TxtPrecioCto_GotFocus()
    SelecTexto TxtPrecioCto
End Sub

Private Sub TxtPrecioCto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtPrecioCto, KeyAscii)
End Sub

Private Sub TxtPrecioCto_LostFocus()
    If TxtPrecioCto.Text <> "" Then
        TxtPrecioCto.Text = Valido_Importe(TxtPrecioCto.Text)
    End If
End Sub

Private Sub TxtPrecioVta_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub TxtPrecioVta_GotFocus()
    SelecTexto TxtPrecioVta
End Sub

Private Sub TxtPrecioVta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtPrecioVta, KeyAscii)
End Sub

Private Sub TxtPrecioVta_LostFocus()
    If TxtPrecioVta.Text <> "" Then
        TxtPrecioVta.Text = Valido_Importe(TxtPrecioVta.Text)
    End If
End Sub

Private Sub txtStockMinimo_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtStockMinimo_GotFocus()
    SelecTexto txtStockMinimo
End Sub

Private Sub txtStockMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
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

