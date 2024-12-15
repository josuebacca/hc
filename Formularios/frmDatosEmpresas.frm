VERSION 5.00
Begin VB.Form frmDatosEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de Empresas..."
   ClientHeight    =   3555
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDatosEmpresas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   90
      Picture         =   "frmDatosEmpresas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3135
      Width           =   330
   End
   Begin VB.TextBox txtEmp_Descri 
      Height          =   285
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   1
      Top             =   450
      Width           =   4920
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1305
      MaxLength       =   10
      TabIndex        =   0
      Top             =   75
      Width           =   1005
   End
   Begin VB.TextBox txtEmp_Cuit 
      Height          =   285
      Left            =   1305
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1980
      Width           =   2280
   End
   Begin VB.TextBox txtEmp_IB 
      Height          =   285
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2355
      Width           =   2280
   End
   Begin VB.TextBox txtEMP_MUNIC 
      Height          =   285
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2715
      Width           =   2280
   End
   Begin VB.TextBox txtEmp_Direccion 
      Height          =   285
      Left            =   1305
      MaxLength       =   100
      TabIndex        =   2
      Top             =   810
      Width           =   4920
   End
   Begin VB.ComboBox cboTipo_Iva 
      Height          =   315
      ItemData        =   "frmDatosEmpresas.frx":0156
      Left            =   1305
      List            =   "frmDatosEmpresas.frx":0158
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1575
      Width           =   2835
   End
   Begin VB.TextBox txtTele_Enti 
      Height          =   285
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1170
      Width           =   2835
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   4935
      TabIndex        =   9
      Top             =   3135
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3555
      TabIndex        =   8
      Top             =   3135
      Width           =   1300
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razón Social"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   17
      Top             =   465
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   150
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   825
      Width           =   675
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INGR. BRUTOS"
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   14
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Munic."
      Height          =   195
      Index           =   7
      Left            =   90
      TabIndex        =   13
      Top             =   2730
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T."
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   12
      Top             =   2010
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Iva"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   1620
      Width           =   585
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   1200
      Width           =   630
   End
End
Attribute VB_Name = "frmDatosEmpresa"
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


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "empresa"
Const cCampoID = "emp_id"
Const cDesRegistro = "Empresa"
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
    If err.Number = 35613 Then
        Call frmPrincipal.mnuContextABM_Click(4)
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
            AcCtrl txtEmp_Descri
            AcCtrl txtEmp_Direccion
            AcCtrl txtTele_Enti
            AcCtrl cboTipo_Iva
            AcCtrl txtEmp_Cuit
            AcCtrl txtEmp_IB
            AcCtrl txtEMP_MUNIC
        Case 3, 4
            DesacCtrl txtEmp_Descri
            DesacCtrl txtEmp_Direccion
            DesacCtrl txtTele_Enti
            DesacCtrl cboTipo_Iva
            DesacCtrl txtEmp_Cuit
            DesacCtrl txtEmp_IB
            DesacCtrl txtEMP_MUNIC
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nueva Empresa..."
            AcCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Empresa..."
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de la Empresa..."
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Empresa..."
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
                             "Ingrese el código de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtEmp_Descri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la descripción de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtEmp_Descri.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboTipo_Iva_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
    
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (emp_id, emp_descri, emp_direccion, emp_cuit, emp_ib, emp_munic, emp_telefono, tiv_id) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & txtID.Text & ", " & XS(txtEmp_Descri.Text) & ", " & XS(txtEmp_Direccion.Text) & ", " & XS(txtEmp_Cuit.Text) & ", " & XS(txtEmp_IB.Text) & ", " & XS(txtEMP_MUNIC.Text) & ", " & XS(txtTele_Enti.Text) & ", " & XN(Trim(Right(Trim(cboTipo_Iva.List(cboTipo_Iva.ListIndex)), 2))) & ")"
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     emp_descri=" & XS(txtEmp_Descri.Text) & ", emp_direccion=" & XS(txtEmp_Direccion.Text) & ", emp_cuit=" & XS(txtEmp_Cuit.Text) & ", emp_ib=" & XS(txtEmp_IB.Text) & ", emp_munic=" & XS(txtEMP_MUNIC.Text) & ", emp_telefono = " & XS(txtTele_Enti.Text) & ", tiv_id = " & XN(Trim(Right(Trim(cboTipo_Iva.List(cboTipo_Iva.ListIndex)), 2)))
                cSQL = cSQL & "     WHERE emp_id =" & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE emp_id  = " & txtID.Text
            
        End Select
        
        DBConn.Execute cSQL
        DBConn.CommitTrans
        On Error GoTo 0
        
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
    ManejoDeErrores DBConn.ErrorNative
    
End Sub


Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 5)
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
    
    If KeyAscii = vbKeyEscape Then Unload Me
    
End Sub

Private Sub Form_Load()

    Dim cSQL As String
    Dim hSQL As String
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    'cargo el combo de tipo de iva
    cboTipo_Iva.Clear
    cSQL = "SELECT * FROM tipoiva order by tiv_descri"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboTipo_Iva.AddItem Trim(rec!tiv_descri) & Space(100) & Trim(Str(rec!tiv_id))
          rec.MoveNext
       Loop
       cboTipo_Iva.ListIndex = cboTipo_Iva.ListIndex + 1
    End If
    rec.Close
    
    
    'Me.Top = vFormLlama.Top + 1500
    'Me.Left = vFormLlama.Left + 1000
    
    txtID.MaxLength = 4
    'txtEmp_Descri.MaxLength = 30
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE emp_id = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!emp_id
                txtEmp_Descri.Text = rec!emp_descri
                txtEmp_Direccion.Text = ChkNull(rec!emp_direccion)
                txtEmp_Cuit.Text = ChkNull(rec!emp_cuit)
                txtEmp_IB.Text = ChkNull(rec!emp_ib)
                txtEMP_MUNIC.Text = ChkNull(rec!emp_munic)
                txtTele_Enti.Text = ChkNull(rec!emp_telefono)
                'pone el tipo de iva
                cboTipo_Iva.ListIndex = 0
                Do While Trim(rec!tiv_id) <> Trim(Right(cboTipo_Iva.List(cboTipo_Iva.ListIndex), 2))
                    cboTipo_Iva.ListIndex = cboTipo_Iva.ListIndex + 1
                Loop
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
    
End Sub


Private Sub txtEmp_Cuit_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtEmp_Cuit_GotFocus()
    seltxt
End Sub

Private Sub txtEmp_Cuit_LostFocus()
    'If Me.ActiveControl.Name = "cmdCerrar" Then
    '    Exit Sub
    'Else
        If txtEmp_Cuit.Text <> "" Then
            If Len(txtEmp_Cuit.Text) > 11 Then
                txtEmp_Cuit.Text = Mid(txtEmp_Cuit.Text, 1, 2) & Mid(txtEmp_Cuit.Text, 4, 8) & Mid(txtEmp_Cuit.Text, 13, 1)
            End If
            If Not ValidoCuit(txtEmp_Cuit.Text) Then
                txtEmp_Cuit.SetFocus
            Else
                txtEmp_Cuit.Text = Mid(txtEmp_Cuit.Text, 1, 2) & "-" & Mid(txtEmp_Cuit.Text, 3, 8) & "-" & Mid(txtEmp_Cuit.Text, 11, 1)
            End If
            
            If vMode = 1 Then
                'valido que no exista otra entidad con el mismo nro. de cuit
                Dim rec As ADODB.Recordset
                Set rec = New ADODB.Recordset
                cSQL = "SELECT * FROM empresa WHERE emp_cuit = " & XS(txtEmp_Cuit.Text)
                rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
                If (rec.BOF And rec.EOF) = False Then
                    MsgBox "Ya existe otra Empresa con este Nro. de CUIT: " & Trim(rec!emp_id) & " - " & Trim(rec!emp_descri) & "." & Chr(13) & "Verifique y vuelva a intentarlo.", vbCritical, TIT_MSGBOX
                    txtEmp_Cuit.SetFocus
                End If
                rec.Close
            End If
            
        End If
        
    'End If
End Sub

Private Sub txtemp_descri_Change()

    cmdAceptar.Enabled = True

End Sub

Private Sub txtEmp_Descri_GotFocus()
    seltxt
End Sub

Private Sub txtEmp_Direccion_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtEmp_Direccion_GotFocus()
    seltxt
End Sub

Private Sub txtEmp_IB_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtEmp_IB_GotFocus()
    seltxt
End Sub

Private Sub txtEMP_MUNIC_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtEMP_MUNIC_GotFocus()
    seltxt
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
            cSQL = "SELECT COUNT(*) FROM " & cTabla & " WHERE " & cCampoID & " = " & txtID.Text
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


Private Sub txtTele_Enti_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtTele_Enti_GotFocus()
    seltxt
End Sub
