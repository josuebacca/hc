VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ABMProveedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Proveedor..."
   ClientHeight    =   5205
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMProveedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   840
      TabIndex        =   29
      Top             =   4755
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ComboBox cboTipoProveedor 
      Height          =   315
      ItemData        =   "ABMProveedor.frx":000C
      Left            =   1185
      List            =   "ABMProveedor.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   3375
   End
   Begin VB.TextBox txtIngresosBrutos 
      Height          =   315
      Left            =   3540
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1680
      Width           =   1005
   End
   Begin MSMask.MaskEdBox txtCuit 
      Height          =   315
      Left            =   1185
      TabIndex        =   4
      Top             =   1680
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboIva 
      Height          =   315
      ItemData        =   "ABMProveedor.frx":0010
      Left            =   1185
      List            =   "ABMProveedor.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1335
      Width           =   3375
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3180
      Width           =   3375
   End
   Begin VB.TextBox txtMail 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   12
      Top             =   4290
      Width           =   3375
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   1185
      MaxLength       =   30
      TabIndex        =   11
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   1185
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3630
      Width           =   3375
   End
   Begin VB.ComboBox cboLocalidad 
      Height          =   315
      ItemData        =   "ABMProveedor.frx":0014
      Left            =   1185
      List            =   "ABMProveedor.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2835
      Width           =   3375
   End
   Begin VB.ComboBox cboProvincia 
      Height          =   315
      ItemData        =   "ABMProveedor.frx":0018
      Left            =   1185
      List            =   "ABMProveedor.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2490
      Width           =   3375
   End
   Begin VB.ComboBox cboPais 
      Height          =   315
      ItemData        =   "ABMProveedor.frx":001C
      Left            =   1185
      List            =   "ABMProveedor.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2145
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMProveedor.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4785
      Width           =   330
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   2
      Top             =   885
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1185
      TabIndex        =   1
      Top             =   540
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3270
      TabIndex        =   14
      Top             =   4785
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1920
      TabIndex        =   13
      Top             =   4785
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Prov:"
      Height          =   195
      Index           =   12
      Left            =   135
      TabIndex        =   28
      Top             =   255
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ing. Brutos:"
      Height          =   195
      Index           =   11
      Left            =   2625
      TabIndex        =   27
      Top             =   1740
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C.U.I.T.:"
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   26
      Top             =   1740
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cond. I.V.A.:"
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   25
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   24
      Top             =   3225
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "e-mail:"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   23
      Top             =   4335
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   22
      Top             =   4005
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   21
      Top             =   3675
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localidad:"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   20
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Provincia:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   19
      Top             =   2535
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "País:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   18
      Top             =   2190
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   930
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   15
      Top             =   570
      Width           =   270
   End
End
Attribute VB_Name = "ABMProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Dim vFieldID As String
Dim vFieldID1 As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String
Dim Pais As String
Dim Provincia As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "PROVEEDOR"
Const cCampoID = "PROV_CODIGO"
Const cDesRegistro = "Proveedor"

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
            AcCtrl txtNombre
            AcCtrl cboIva
            AcCtrl txtCuit
            AcCtrl txtIngresosBrutos
            AcCtrl cboPais
            AcCtrl cboProvincia
            AcCtrl cboLocalidad
            AcCtrl txtDomicilio
            AcCtrl txtTelefono
            AcCtrl txtFax
            AcCtrl txtMail
        Case 3, 4
            DesacCtrl txtNombre
            DesacCtrl cboIva
            DesacCtrl txtCuit
            DesacCtrl txtIngresosBrutos
            DesacCtrl cboPais
            DesacCtrl cboProvincia
            DesacCtrl cboLocalidad
            DesacCtrl txtDomicilio
            DesacCtrl txtTelefono
            DesacCtrl txtFax
            DesacCtrl txtMail
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo " & cDesRegistro
            DesacCtrl txtID
            AcCtrl cboTipoProveedor
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro
            DesacCtrl txtID
            DesacCtrl cboTipoProveedor
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del " & cDesRegistro
            DesacCtrl txtID
            DesacCtrl cboTipoProveedor
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro
            DesacCtrl txtID
            DesacCtrl cboTipoProveedor
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
            vFieldID1 = vListView.SelectedItem.SubItems(3) 'CODIGO DE TIPO DE PROVEEDOR
        Else
            vFieldID = 0
            vFieldID1 = 0
        End If
    Else
        vFieldID = 0
        vFieldID1 = 0
    End If

End Function


Function Validar(pMode As Integer) As Boolean

    If pMode <> 4 Then
        Validar = False
        If txtID.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Identificación del  " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtNombre.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Nombre del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtNombre.SetFocus
            Exit Function
        
        ElseIf cboPais.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Paí del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboPais.SetFocus
            Exit Function
            
        ElseIf cboProvincia.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Provincia del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboPais.SetFocus
            Exit Function
        
        ElseIf cboLocalidad.ListIndex = -1 Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Localidad del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            cboProvincia.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboIva_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboLocalidad_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboPais_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboPais_LostFocus()
    If vMode = 2 And Pais = cboPais.Text Then
        Exit Sub
    End If
    Set Rec1 = New ADODB.Recordset
    cboProvincia.Clear
    SQL = "SELECT PRO_CODIGO,PRO_DESCRI"
    SQL = SQL & " FROM PROVINCIA "
    SQL = SQL & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    SQL = SQL & " ORDER BY PRO_DESCRI"
    
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
       Do While Rec1.EOF = False
          cboProvincia.AddItem Trim(Rec1!PRO_DESCRI)
          cboProvincia.ItemData(cboProvincia.NewIndex) = Rec1!PRO_CODIGO
          Rec1.MoveNext
       Loop
       cboProvincia.ListIndex = cboProvincia.ListIndex + 1
    Else
       MsgBox "No hay cargado Provincia para ese País.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    Rec1.Close
End Sub

Private Sub cboProvincia_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cboProvincia_LostFocus()
    If vMode = 2 And Provincia = cboProvincia.Text Then
        Exit Sub
    End If
    Set Rec1 = New ADODB.Recordset
    cboLocalidad.Clear
    SQL = "SELECT LOC_CODIGO,LOC_DESCRI FROM LOCALIDAD"
    SQL = SQL & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    SQL = SQL & " AND PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
    SQL = SQL & " ORDER BY LOC_DESCRI "
    
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If (Rec1.BOF And Rec1.EOF) = 0 Then
       Do While Rec1.EOF = False
          cboLocalidad.AddItem Trim(Rec1!LOC_DESCRI)
          cboLocalidad.ItemData(cboLocalidad.NewIndex) = Rec1!LOC_CODIGO
          Rec1.MoveNext
       Loop
       cboLocalidad.ListIndex = cboLocalidad.ListIndex + 1
    Else
       MsgBox "No hay cargada Localidad para esta Provincia.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    Rec1.Close
End Sub

Private Sub cboTipoProveedor_LostFocus()
    AcCtrl txtID
    txtID_LostFocus
    DesacCtrl txtID
End Sub

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (PROV_CODIGO, PROV_RAZSOC, PROV_DOMICI, PROV_CUIT,"
                cSQL = cSQL & " PROV_INGBRU, TPR_CODIGO, IVA_CODIGO,"
                cSQL = cSQL & " PROV_TELEFONO, PROV_MAIL, PROV_FAX,"
                cSQL = cSQL & " LOC_CODIGO, PRO_CODIGO, PAI_CODIGO) "
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text) & ", " & XS(txtCuit.Text) & ", "
                cSQL = cSQL & XS(txtIngresosBrutos.Text) & ", "
                cSQL = cSQL & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex) & ", "
                cSQL = cSQL & cboIva.ItemData(cboIva.ListIndex) & ", "
                cSQL = cSQL & XS(txtTelefono.Text) & ", "
                cSQL = cSQL & XS(txtMail.Text) & ", " & XS(txtFax.Text) & ", "
                cSQL = cSQL & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                cSQL = cSQL & cboProvincia.ItemData(cboProvincia.ListIndex) & ", "
                cSQL = cSQL & cboPais.ItemData(cboPais.ListIndex) & ")"
                
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  PROV_RAZSOC=" & XS(txtNombre.Text)
                cSQL = cSQL & " ,PROV_DOMICI=" & XS(txtDomicilio.Text)
                cSQL = cSQL & " ,PROV_CUIT=" & XS(txtCuit.Text)
                cSQL = cSQL & " ,PROV_INGBRU=" & XS(txtIngresosBrutos.Text)
                cSQL = cSQL & " ,IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
                cSQL = cSQL & " ,PROV_TELEFONO=" & XS(txtTelefono.Text)
                cSQL = cSQL & " ,PROV_MAIL=" & XS(txtMail.Text)
                cSQL = cSQL & " ,PROV_FAX=" & XS(txtFax.Text)
                cSQL = cSQL & " ,LOC_CODIGO=" & cboLocalidad.ItemData(cboLocalidad.ListIndex)
                cSQL = cSQL & " ,PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
                cSQL = cSQL & " ,PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " WHERE PROV_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
            
            Case 4 'eliminar
                cSQL = "DELETE FROM " & cTabla & " WHERE PROV_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND TPR_CODIGO=" & cboTipoProveedor.ItemData(cboTipoProveedor.ListIndex)
                
        End Select
        
        DBConn.Execute cSQL
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

Private Sub Command1_Click()
    SQL = "SELECT * FROM XX_PROVEEDORES"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            SQL = "INSERT INTO PROVEEDOR (TPR_CODIGO,PROV_CODIGO,PROV_RAZSOC,"
            SQL = SQL & " PROV_DOMICI,PROV_CUIT,PROV_INGBRU,PROV_TELEFONO,PROV_MAIL,"
            SQL = SQL & " IVA_CODIGO,PAI_CODIGO,PRO_CODIGO,LOC_CODIGO) VALUES (1,"
            SQL = SQL & XN(rec!ID) & ","
            SQL = SQL & XS(rec!razon) & ","
            SQL = SQL & XS(ChkNull(rec!domicilio)) & ","
            SQL = SQL & XS(ChkNull(rec!cuit)) & ","
            SQL = SQL & XS(ChkNull(rec!ing_bru)) & ","
            SQL = SQL & XS(ChkNull(rec!tel)) & ","
            SQL = SQL & XS(ChkNull(rec!email)) & ","
            SQL = SQL & "1,1,1,13)"
            DBConn.Execute SQL
            rec.MoveNext
        Loop
    End If
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
    'CARGO COMBO CONDICIN IVA
    Call CargoComboBox(cboIva, "CONDICION_IVA", "IVA_CODIGO", "IVA_DESCRI")
    If cboIva.ListCount > 0 Then
        cboIva.ListIndex = 0
    End If
    
    'CARGO COMBO CANALES
    Call CargoComboBox(cboTipoProveedor, "TIPO_PROVEEDOR", "TPR_CODIGO", "TPR_DESCRI")
    If cboTipoProveedor.ListCount > 0 Then
        cboTipoProveedor.ListIndex = 0
    End If
    
    'cargo el combo de PAIS
    cboPais.Clear
    cSQL = "SELECT * FROM PAIS ORDER BY PAI_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboPais.AddItem Trim(rec!PAI_DESCRI)
          cboPais.ItemData(cboPais.NewIndex) = rec!PAI_CODIGO
          rec.MoveNext
       Loop
       cboPais.ListIndex = cboPais.ListIndex + 1
    End If
    rec.Close
    
    Pais = ""
    Provincia = ""
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE PROV_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            cSQL = cSQL & " AND TPR_CODIGO=" & vFieldID1
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!PROV_CODIGO
                txtNombre.Text = rec!PROV_RAZSOC
                
                Call BuscaCodigoProxItemData(rec!IVA_CODIGO, cboIva)
                txtCuit.Text = ChkNull(rec!PROV_CUIT)
                txtIngresosBrutos.Text = ChkNull(rec!PROV_INGBRU)
                Call BuscaCodigoProxItemData(rec!TPR_CODIGO, cboTipoProveedor)
                
                Call BuscaCodigoProxItemData(CInt(rec!PAI_CODIGO), cboPais)
                cboPais_LostFocus
                Pais = cboPais.Text
                
                Call BuscaCodigoProxItemData(CInt(rec!PRO_CODIGO), cboProvincia)
                cboProvincia_LostFocus
                Provincia = cboProvincia.Text
                
                Call BuscaCodigoProxItemData(CInt(rec!LOC_CODIGO), cboLocalidad)
                txtDomicilio.Text = ChkNull(rec!PROV_DOMICI)
                txtTelefono.Text = ChkNull(rec!PROV_TELEFONO)
                txtFax.Text = ChkNull(rec!PROV_FAX)
                txtMail.Text = ChkNull(rec!PROV_MAIL)
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub

Private Sub txtCuit_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCuit_GotFocus()
    SelecTexto txtCuit
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCuit_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtCuit.ClipText)) = 12 Then
      txtCuit.SelStart = 12
  End If
End Sub

Private Sub txtCuit_LostFocus()
    If txtCuit.Text <> "" Then
        'rutina de validación de CUIT
        If Not ValidoCuit(txtCuit) Then
            txtCuit.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDomicilio_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDomicilio_GotFocus()
    SelecTexto txtDomicilio
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtFax_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtFax_GotFocus()
    SelecTexto txtFax
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtIngresosBrutos_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtIngresosBrutos_GotFocus()
    SelecTexto txtIngresosBrutos
End Sub

Private Sub txtIngresosBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtMail_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtMail_GotFocus()
    SelecTexto txtMail
End Sub

Private Sub txtNombre_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtNombre_GotFocus()
    seltxt
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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

Private Sub txtTelefono_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtTelefono_GotFocus()
    SelecTexto txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
