VERSION 5.00
Begin VB.Form ABMComisionVendedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de la Comisión..."
   ClientHeight    =   2340
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMComisionVendedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPorCob 
      Height          =   315
      Left            =   3705
      TabIndex        =   4
      Top             =   1350
      Width           =   930
   End
   Begin VB.ComboBox cboRepresentada 
      Height          =   315
      ItemData        =   "ABMComisionVendedores.frx":000C
      Left            =   1260
      List            =   "ABMComisionVendedores.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   645
      Width           =   3375
   End
   Begin VB.ComboBox CboVendedor 
      Height          =   315
      ItemData        =   "ABMComisionVendedores.frx":0010
      Left            =   1260
      List            =   "ABMComisionVendedores.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   285
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   255
      Picture         =   "ABMComisionVendedores.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1890
      Width           =   330
   End
   Begin VB.TextBox txtPorVta 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   1350
      Width           =   930
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   1005
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3360
      TabIndex        =   6
      Top             =   1890
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2010
      TabIndex        =   5
      Top             =   1890
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "% Cobranza:"
      Height          =   195
      Index           =   4
      Left            =   2625
      TabIndex        =   12
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Representada:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   11
      Top             =   690
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   330
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "% Venta:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   1395
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   1035
      Width           =   270
   End
End
Attribute VB_Name = "ABMComisionVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuración de la ventana de datos
Dim vFieldID As String
Dim vFieldID1 As String
Dim vFieldID2 As String
Dim vStringSQL As String
Dim vFormLlama As Form
Dim vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "COMISION"
Const cCampoID = "COM_CODIGO"
Const cDesRegistro = "Comisión"

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
        Case 1
            AcCtrl CboVendedor
            AcCtrl cboRepresentada
        Case 2, 3, 4
            DesacCtrl CboVendedor
            DesacCtrl cboRepresentada
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nueva " & cDesRegistro & "..."
            AcCtrl txtID
            AcCtrl txtPorCob
            AcCtrl txtPorVta
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando " & cDesRegistro & "..."
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de la " & cDesRegistro & "..."
            DesacCtrl txtID
            DesacCtrl txtPorCob
            DesacCtrl txtPorVta
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando " & cDesRegistro & "..."
            DesacCtrl txtID
            DesacCtrl txtPorCob
            DesacCtrl txtPorVta
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
            vFieldID1 = vListView.SelectedItem.SubItems(1) 'VENDEDOR
            vFieldID2 = vListView.SelectedItem.SubItems(3) 'REPRESENTADA
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
                             "Ingrese la Identificación de la " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
    
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & " (VEN_CODIGO, REP_CODIGO , COM_CODIGO,"
                cSQL = cSQL & " COM_PORVTA, COM_PORCOB)"
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & " (" & CboVendedor.ItemData(CboVendedor.ListIndex) & ", "
                cSQL = cSQL & cboRepresentada.ItemData(cboRepresentada.ListIndex) & ", "
                cSQL = cSQL & XN(txtID.Text) & ", " & XN(txtPorVta.Text) & ", "
                cSQL = cSQL & XN(txtPorCob.Text) & ")"
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     COM_PORVTA = " & XN(txtPorVta.Text)
                cSQL = cSQL & "    ,COM_PORCOB = " & XN(txtPorCob.Text)
                cSQL = cSQL & " WHERE COM_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND VEN_CODIGO = " & CboVendedor.ItemData(CboVendedor.ListIndex)
                cSQL = cSQL & " AND REP_CODIGO = " & cboRepresentada.ItemData(cboRepresentada.ListIndex)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE COM_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND VEN_CODIGO = " & CboVendedor.ItemData(CboVendedor.ListIndex)
                cSQL = cSQL & " AND REP_CODIGO = " & cboRepresentada.ItemData(cboRepresentada.ListIndex)
            
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
    ManejoDeErrores DBConn.ErrorNative
    
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
    'CARGO COMBO VENDEDOR
    Call CargoComboBox(CboVendedor, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE")
    If CboVendedor.ListCount > 0 Then
        CboVendedor.ListIndex = 0
    End If
    
    'CARGO COMBO REPRESENTADA
    Call CargoComboBox(cboRepresentada, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    If cboRepresentada.ListCount > 0 Then
        cboRepresentada.ListIndex = 0
    End If
    
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE COM_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            cSQL = cSQL & " AND VEN_CODIGO = " & Mid(vFieldID1, 1, 10)
            cSQL = cSQL & " AND REP_CODIGO = " & Mid(vFieldID2, 1, 10)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                Call BuscaCodigoProxItemData(CInt(rec!VEN_CODIGO), CboVendedor)
                Call BuscaCodigoProxItemData(CInt(rec!REP_CODIGO), cboRepresentada)
                txtID.Text = rec!COM_CODIGO
                txtPorCob.Text = Format(rec!COM_PORCOB, "0.00")
                txtPorVta.Text = Format(rec!COM_PORVTA, "0.00")
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
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

Private Sub txtPorCob_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtPorCob_GotFocus()
    SelecTexto txtPorCob
End Sub

Private Sub txtPorCob_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorCob, KeyAscii)
End Sub

Private Sub txtPorCob_LostFocus()
    If txtPorCob.Text <> "" Then
        If ValidarPorcentaje(txtPorCob) = False Then
        txtPorCob.Text = "0,00"
        End If
    Else
        txtPorCob.Text = "0,00"
    End If
End Sub

Private Sub txtPorVta_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtPorVta_GotFocus()
    SelecTexto txtPorVta
End Sub

Private Sub txtPorVta_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorVta, KeyAscii)
End Sub

Private Sub txtPorVta_LostFocus()
    If txtPorVta.Text <> "" Then
        If ValidarPorcentaje(txtPorVta) = False Then
            txtPorVta.Text = "0,00"
        End If
    Else
        txtPorVta.Text = "0,00"
    End If
End Sub
