VERSION 5.00
Begin VB.Form ABMTarjetaPlan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Plan..."
   ClientHeight    =   2220
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
   Icon            =   "ABMTarjetaPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCuotas 
      Height          =   300
      Left            =   1065
      TabIndex        =   3
      Top             =   1350
      Width           =   720
   End
   Begin VB.ComboBox cboTarjeta 
      Height          =   315
      ItemData        =   "ABMTarjetaPlan.frx":000C
      Left            =   1065
      List            =   "ABMTarjetaPlan.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   285
      Width           =   3375
   End
   Begin VB.TextBox txtDescri 
      Height          =   300
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   2
      Top             =   990
      Width           =   3375
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   645
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   5
      Top             =   1800
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuotas:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Top             =   1380
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tarjeta:"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   330
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripci�n:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   1035
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   675
      Width           =   270
   End
End
Attribute VB_Name = "ABMTarjetaPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'parametros para la configuraci�n de la ventana de datos
Dim vFieldID As String
Dim vFieldID1 As String
Dim vStringSQL As String
Dim vFormLlama As Form
Public vMode As Integer
Dim vListView As ListView
Dim vDesFieldID As String


'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "TARJETA_PLAN"
Const cCampoID = "PLA_CODIGO"
Const cDesRegistro = "Tarjeta Plan"

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
        
            'recorro la coleci�n de campos a actualizar
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
    'Parametro: pMode indica el modo en que se utilizar� este form
    '  pMode  =             1> Indica nuevo registro
    '                       2> Editar registro existente
    '                       3> Mostrar dato del registro existente
    '                       4> Eliminar registro existente
    
    Select Case pMode
        Case 1
            AcCtrl cboTarjeta
        Case 2, 3, 4
            DesacCtrl cboTarjeta
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo Plan.."
            AcCtrl txtID
            AcCtrl txtDescri
            AcCtrl txtCuotas
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Plan..."
            DesacCtrl txtID
            'DesacCtrl txtDescri
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del Plan..."
            DesacCtrl txtID
            DesacCtrl txtDescri
            DesacCtrl txtCuotas
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Plan..."
            DesacCtrl txtID
            DesacCtrl txtDescri
            DesacCtrl txtCuotas
    End Select
    
End Function

Public Function SetWindow(pWindow As Form, pSQL As String, pMode As Integer, pListview As ListView, pDesID As String)
    
    Set vFormLlama = pWindow 'Objeto ventana que que llama a la ventana de datos
    vStringSQL = pSQL 'string utilizado para argar la lista base
    vMode = pMode  'modo en que se utilizar� la ventana de datos
    Set vListView = pListview 'objeto listview que se est� editando
    vDesFieldID = pDesID 'nombre del campo identificador
    
    'valor del campo identificador de registro seleccionado (0 si es un reg. nuevo)
    If vMode <> 1 Then
        If vListView.SelectedItem.Selected = True Then
            vFieldID = vListView.SelectedItem.Key
            vFieldID1 = vListView.SelectedItem.SubItems(3)
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
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese la Identificaci�n del Plan antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese la descripci�n del Plan antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
            
        ElseIf txtCuotas.Text = "" Then
            Beep
            MsgBox "Falta informaci�n." & Chr(13) & _
                             "Ingrese las cuotas del Plan antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtCuotas.SetFocus
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
                cSQL = cSQL & "     (TAR_CODIGO, PLA_CODIGO ,PLA_DESCRI, PLA_CUOTAS) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & cboTarjeta.ItemData(cboTarjeta.ListIndex) & ", "
                cSQL = cSQL & XN(txtID.Text) & ", " & XS(txtDescri.Text) & ", " & XN(txtCuotas.Text) & ") "
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  PLA_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & "  ,PLA_CUOTAS = " & XN(txtCuotas.Text)
                cSQL = cSQL & " WHERE PLA_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND TAR_CODIGO = " & cboTarjeta.ItemData(cboTarjeta.ListIndex)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE PLA_CODIGO  = " & XN(txtID.Text)
                cSQL = cSQL & " AND TAR_CODIGO = " & cboTarjeta.ItemData(cboTarjeta.ListIndex)
            
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
    'cargo el combo de BANCOS
    cboTarjeta.Clear
    cSQL = "SELECT * FROM TARJETA order by TAR_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboTarjeta.AddItem Trim(rec!TAR_DESCRI)
          cboTarjeta.ItemData(cboTarjeta.NewIndex) = rec!TAR_CODIGO
          rec.MoveNext
       Loop
       cboTarjeta.ListIndex = cboTarjeta.ListIndex + 1
    End If
    rec.Close
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE PLA_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            cSQL = cSQL & " AND TAR_CODIGO = " & Mid(vFieldID1, 1, 10)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontr� el registro muestro los datos
                Call BuscaCodigoProxItemData(CInt(rec!TAR_CODIGO), cboTarjeta)
                txtID.Text = rec!PLA_CODIGO
                txtDescri.Text = Trim(rec!PLA_DESCRI)
                txtCuotas.Text = Trim(rec!PLA_CUOTAS)
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub

Private Sub txtCuotas_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtCuotas_GotFocus()
    seltxt
End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
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
                'cSQL = cSQL & " WHERE TAR_CODIGO = " & cboTarjeta.ItemData(cboTarjeta.ListIndex)
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
            'cSQL = cSQL & " AND TAR_CODIGO = " & cboTarjeta.ItemData(cboTarjeta.ListIndex)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                If rec.Fields(0) > 0 Then
                    Beep
                    MsgBox "C�digo de " & cDesRegistro & " repetido." & Chr(13) & _
                                     "El c�digo ingresado Pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
End Sub
