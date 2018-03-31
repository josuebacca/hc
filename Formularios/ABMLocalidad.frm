VERSION 5.00
Begin VB.Form ABMLocalidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de la Localidad..."
   ClientHeight    =   2205
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
   Icon            =   "ABMLocalidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboProvincia 
      Height          =   315
      ItemData        =   "ABMLocalidad.frx":000C
      Left            =   1065
      List            =   "ABMLocalidad.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   645
      Width           =   3375
   End
   Begin VB.ComboBox cboPais 
      Height          =   315
      ItemData        =   "ABMLocalidad.frx":0010
      Left            =   1065
      List            =   "ABMLocalidad.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   285
      Width           =   3375
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMLocalidad.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1770
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtDescri 
      Height          =   300
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1350
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   1065
      TabIndex        =   2
      Top             =   1005
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   5
      Top             =   1770
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   1770
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Provincia:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   10
      Top             =   690
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "País:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   330
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   1395
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   1035
      Width           =   270
   End
End
Attribute VB_Name = "ABMLocalidad"
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
Const cTabla = "LOCALIDAD"
Const cCampoID = "LOC_CODIGO"
Const cDesRegistro = "Localidad"

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
            AcCtrl cboPais
            AcCtrl cboProvincia
        Case 2, 3, 4
            DesacCtrl cboPais
            DesacCtrl cboProvincia
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nueva Localidad.."
            AcCtrl txtDescri
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Localidad..."
            DesacCtrl txtID
            'DesacCtrl txtDescri
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de la Localidad..."
            DesacCtrl txtID
            DesacCtrl txtDescri
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Localidad..."
            DesacCtrl txtID
            DesacCtrl txtDescri
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
            vFieldID1 = vListView.SelectedItem.SubItems(3) 'PROVINCIA
            vFieldID2 = vListView.SelectedItem.SubItems(5) 'PAIS
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
                             "Ingrese la Identificación de la Localidad antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la descripción de la Localidad antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cboPais_LostFocus()
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

Private Sub cboProvincia_LostFocus()
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
                cSQL = cSQL & "     (PAI_CODIGO, PRO_CODIGO ,LOC_CODIGO, LOC_DESCRI) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & cboPais.ItemData(cboPais.ListIndex) & ", "
                cSQL = cSQL & cboProvincia.ItemData(cboProvincia.ListIndex) & ", "
                cSQL = cSQL & XN(txtID.Text) & ", " & XS(txtDescri.Text) & ") "
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     LOC_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & " WHERE LOC_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " AND PRO_CODIGO = " & cboProvincia.ItemData(cboProvincia.ListIndex)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE LOC_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " AND PRO_CODIGO = " & cboProvincia.ItemData(cboProvincia.ListIndex)
            
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
    
    'txtID.MaxLength = 4
    'txtDescri.MaxLength = 30
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
    
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE LOC_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            cSQL = cSQL & " AND PRO_CODIGO = " & Mid(vFieldID1, 1, 10)
            cSQL = cSQL & " AND PAI_CODIGO = " & Mid(vFieldID2, 1, 10)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                Call BuscaCodigoProxItemData(CInt(rec!PAI_CODIGO), cboPais)
                cboPais_LostFocus
                Call BuscaCodigoProxItemData(CInt(rec!PRO_CODIGO), cboProvincia)
                txtID.Text = rec!LOC_CODIGO
                txtDescri.Text = rec!LOC_DESCRI
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
                cSQL = cSQL & " WHERE PAI_CODIGO = " & cboPais.ItemData(cboPais.ListIndex)
                cSQL = cSQL & " AND PRO_CODIGO = " & cboProvincia.ItemData(cboProvincia.ListIndex)
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
