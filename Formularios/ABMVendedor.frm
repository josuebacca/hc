VERSION 5.00
Begin VB.Form ABMVendedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Personal..."
   ClientHeight    =   5445
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtConsul 
      Height          =   315
      Left            =   1080
      TabIndex        =   25
      Top             =   1320
      Width           =   720
   End
   Begin VB.ComboBox cboprofesion 
      Height          =   315
      ItemData        =   "ABMVendedor.frx":000C
      Left            =   1050
      List            =   "ABMVendedor.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.CheckBox chkVenEstado 
      Caption         =   "Dar de Baja"
      Height          =   285
      Left            =   1065
      TabIndex        =   10
      Top             =   4680
      Width           =   1140
   End
   Begin VB.TextBox txtDomicilio 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2715
      Width           =   3375
   End
   Begin VB.TextBox txtMail 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4065
      Width           =   3375
   End
   Begin VB.TextBox txtFax 
      Height          =   315
      Left            =   1065
      MaxLength       =   30
      TabIndex        =   8
      Top             =   3615
      Width           =   3375
   End
   Begin VB.TextBox txtTelefono 
      Height          =   315
      Left            =   1065
      MaxLength       =   30
      TabIndex        =   7
      Top             =   3165
      Width           =   3375
   End
   Begin VB.ComboBox cboLocalidad 
      Height          =   315
      ItemData        =   "ABMVendedor.frx":0010
      Left            =   1065
      List            =   "ABMVendedor.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2370
      Width           =   3375
   End
   Begin VB.ComboBox cboProvincia 
      Height          =   315
      ItemData        =   "ABMVendedor.frx":0014
      Left            =   1065
      List            =   "ABMVendedor.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2025
      Width           =   3375
   End
   Begin VB.ComboBox cboPais 
      Height          =   315
      ItemData        =   "ABMVendedor.frx":0018
      Left            =   1065
      List            =   "ABMVendedor.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMVendedor.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4935
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1065
      TabIndex        =   0
      Top             =   285
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   12
      Top             =   4935
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   11
      Top             =   4935
      Width           =   1300
   End
   Begin VB.Label Label2 
      Caption         =   "Consultorio:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ocupacion:"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   22
      Top             =   2760
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "e-mail:"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   21
      Top             =   4110
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   20
      Top             =   3660
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   19
      Top             =   3210
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Localidad:"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   18
      Top             =   2415
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Provincia:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   17
      Top             =   2070
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "País:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   16
      Top             =   1725
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   14
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   13
      Top             =   315
      Width           =   270
   End
End
Attribute VB_Name = "ABMVendedor"
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
Dim Pais As String
Dim Provincia As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "VENDEDOR"
Const cCampoID = "VEN_CODIGO"
Const cDesRegistro = "Personal"

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
            AcCtrl cboPais
            AcCtrl cboProvincia
            AcCtrl cboLocalidad
            AcCtrl cboprofesion
            AcCtrl txtDomicilio
            AcCtrl txtTelefono
            AcCtrl txtFax
            AcCtrl txtMail
            AcCtrl chkVenEstado
        Case 3, 4
            DesacCtrl txtNombre
            DesacCtrl cboPais
            DesacCtrl cboProvincia
            DesacCtrl cboLocalidad
            DesacCtrl cboprofesion
            DesacCtrl txtDomicilio
            DesacCtrl txtTelefono
            DesacCtrl txtFax
            DesacCtrl txtMail
            DesacCtrl chkVenEstado
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo " & cDesRegistro
            txtID_LostFocus
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del " & cDesRegistro
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro
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
                             "Ingrese el País del " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
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
    sql = "SELECT PRO_CODIGO,PRO_DESCRI"
    sql = sql & " FROM PROVINCIA "
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    sql = sql & " ORDER BY PRO_CODIGO"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
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
    sql = "SELECT LOC_CODIGO,LOC_DESCRI FROM LOCALIDAD"
    sql = sql & " WHERE PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
    If cboProvincia.ListIndex <> -1 Then
        sql = sql & " AND PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
    End If
    sql = sql & " ORDER BY LOC_CODIGO "
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
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

Private Sub chkVenEstado_Click()
    cmdAceptar.Enabled = True
End Sub

Private Sub cmdAceptar_Click()
    'sql = "SELECT VEN_CONSULTORIO FROM VENDEDOR"
    'sql = sql & " WHERE VEN_CODIGO = "
    'sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex)
    'rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'txtConsultorio.Text = rec!VEN_CONSULTORIO
    'txtProfesion.Text = rec!PR_CODIGO
    'rec.Close
    Dim cSQL As String
    
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (VEN_CODIGO, VEN_NOMBRE, VEN_DOMICI, VEN_TELEFONO,"
                cSQL = cSQL & " VEN_MAIL, VEN_FAX, LOC_CODIGO, PRO_CODIGO, PAI_CODIGO,PR_CODIGO,VEN_ESTADO,VEN_CONSULTORIO) "
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text) & ", " & XS(txtTelefono.Text) & ", "
                cSQL = cSQL & XS(txtMail.Text) & ", " & XS(txtFax.Text) & ", "
                cSQL = cSQL & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                cSQL = cSQL & cboProvincia.ItemData(cboProvincia.ListIndex) & ", "
                cSQL = cSQL & cboPais.ItemData(cboPais.ListIndex) & ","
                cSQL = cSQL & cboprofesion.ItemData(cboprofesion.ListIndex) & ","
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & "'S')"
                Else
                    cSQL = cSQL & "'N')"
                End If
            cSQL = cSQL & XN(txtConsul.Text)
                
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  VEN_NOMBRE=" & XS(txtNombre.Text)
                cSQL = cSQL & " ,VEN_DOMICI=" & XS(txtDomicilio.Text)
                cSQL = cSQL & " ,VEN_TELEFONO=" & XS(txtTelefono.Text)
                cSQL = cSQL & " ,VEN_MAIL=" & XS(txtMail.Text)
                cSQL = cSQL & " ,VEN_FAX=" & XS(txtFax.Text)
                cSQL = cSQL & " ,LOC_CODIGO=" & cboLocalidad.ItemData(cboLocalidad.ListIndex)
                cSQL = cSQL & " ,PRO_CODIGO=" & cboProvincia.ItemData(cboProvincia.ListIndex)
                cSQL = cSQL & " ,PAI_CODIGO=" & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " ,PR_CODIGO=" & cboprofesion.ItemData(cboprofesion.ListIndex)
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & " ,VEN_ESTADO = 'S'"
                Else
                    cSQL = cSQL & " ,VEN_ESTADO = 'N'"
                End If
                cSQL = cSQL & " ,VEN_CONSULTORIO=" & XN(txtConsul.Text)
                cSQL = cSQL & " WHERE VEN_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE VEN_CODIGO  = " & XN(txtID.Text)
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
    
    'cargo el combo de PROFESION
    CargoCboProfesion
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
            cSQL = "SELECT * FROM " & cTabla & "  WHERE VEN_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                txtID.Text = rec!VEN_CODIGO
                txtNombre.Text = rec!VEN_NOMBRE
                'si encontró el registro muestro los datos
                Call BuscaCodigoProxItemData(CInt(rec!PAI_CODIGO), cboPais)
                cboPais_LostFocus
                Pais = cboPais.Text
                
                Call BuscaCodigoProxItemData(CInt(rec!PRO_CODIGO), cboProvincia)
                cboProvincia_LostFocus
                Provincia = cboProvincia.Text
                
                
                Call BuscaCodigoProxItemData(CInt(rec!LOC_CODIGO), cboLocalidad)
                
                Call BuscaCodigoProxItemData(CInt(rec!PR_CODIGO), cboprofesion)
                
                txtDomicilio.Text = ChkNull(rec!VEN_DOMICI)
                txtTelefono.Text = ChkNull(rec!VEN_TELEFONO)
                txtFax.Text = ChkNull(rec!VEN_FAX)
                txtMail.Text = ChkNull(rec!VEN_MAIL)
                
                If ChkNull(rec!VEN_ESTADO) = "N" Or ChkNull(rec!VEN_ESTADO) = "" Then
                    chkVenEstado.Value = Unchecked
                Else
                    chkVenEstado.Value = Checked
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
Private Sub CargoCboProfesion()
    cboprofesion.Clear
    cSQL = "SELECT * FROM PROFESION ORDER BY PR_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboprofesion.AddItem Trim(rec!PR_DESCRI)
          cboprofesion.ItemData(cboprofesion.NewIndex) = rec!PR_CODIGO
          rec.MoveNext
       Loop
       cboprofesion.ListIndex = cboprofesion.ListIndex + 1
    End If
    rec.Close
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
