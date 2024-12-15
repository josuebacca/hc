VERSION 5.00
Begin VB.Form ABMTipoRevelado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Tipo Revelado..."
   ClientHeight    =   2025
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
   Icon            =   "ABMTipoRevelado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   210
      Picture         =   "ABMTipoRevelado.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1605
      Width           =   330
   End
   Begin VB.TextBox txtDescri 
      Height          =   300
      Left            =   210
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1200
      Width           =   4275
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   210
      TabIndex        =   0
      Top             =   465
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3120
      TabIndex        =   3
      Top             =   1605
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1740
      TabIndex        =   2
      Top             =   1605
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   975
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   270
   End
End
Attribute VB_Name = "ABMTipoRevelado"
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
Const cTabla = "TIPO_REVELADO"
Const cCampoID = "TRE_CODIGO"
Const cDesRegistro = "Tipo Revelado"

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
            AcCtrl txtDescri
        Case 3, 4
            DesacCtrl txtDescri
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo Tipo Revelado..."
            txtID_LostFocus
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Tipo Revelado..."
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del Tipo Revelado..."
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Tipo Revelado..."
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
                             "Ingrese la Identificación del Tipo Revelado antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la descripción del Tipo Revelado antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
        End If
    End If
    
    Validar = True
    
End Function

Private Sub cmdAceptar_Click()

    Dim cSQL As String
    
    If Validar(vMode) = True Then
        
        On Error GoTo ErrorTran
        
        Screen.MousePointer = vbHourglass
    
        DBConn.BeginTrans
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (TRE_CODIGO, TRE_DESCRI) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtDescri.Text) & ") "
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     TRE_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & " WHERE TRE_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE TRE_CODIGO  = " & XN(txtID.Text)
            
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
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE TRE_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!TRE_CODIGO
                txtDescri.Text = rec!TRE_DESCRI
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
