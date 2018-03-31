VERSION 5.00
Begin VB.Form ABMMedicamentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Medicamento..."
   ClientHeight    =   4230
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
   Icon            =   "ABMMedicamentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optAdultos 
      Caption         =   "Para Adultos"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.OptionButton optNiños 
      Caption         =   "Para Niños"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   3120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.ComboBox cboGrupo 
      Height          =   315
      ItemData        =   "ABMMedicamentos.frx":000C
      Left            =   1170
      List            =   "ABMMedicamentos.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CheckBox chkVenEstado 
      Caption         =   "Dar de Baja"
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   3600
      Width           =   1140
   End
   Begin VB.TextBox txtPresentacion 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1035
      Width           =   3255
   End
   Begin VB.TextBox txtDosificacion 
      Height          =   1035
      Left            =   1185
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1485
      Width           =   3255
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMMedicamentos.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3855
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   3255
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1185
      TabIndex        =   0
      Top             =   285
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   9
      Top             =   3855
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   8
      Top             =   3855
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grupo:"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   2685
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Presentacion:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   14
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dosificación:"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   13
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   315
      Width           =   270
   End
End
Attribute VB_Name = "ABMMedicamentos"
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
Const cTabla = "MEDICAMENTOS"
Const cCampoID = "MED_CODIGO"
Const cDesRegistro = "Medicamento"

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
            AcCtrl txtPresentacion
            AcCtrl txtDosificacion
            AcCtrl chkVenEstado
            AcCtrl optAdultos
            AcCtrl optNiños
        Case 3, 4
            DesacCtrl txtNombre
            DesacCtrl txtPresentacion
            DesacCtrl txtDosificacion
            DesacCtrl chkVenEstado
            DesacCtrl optAdultos
            DesacCtrl optNiños
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
        End If
    End If
    
    Validar = True
    
End Function

Private Sub chkVenEstado_Click()
    cmdAceptar.Enabled = True
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
                cSQL = cSQL & "     (MED_CODIGO, MED_NOMBRE, MED_PRESENTACION, MED_DOSIFICACION,"
                cSQL = cSQL & " GRU_CODIGO, MED_EDAD,MED_ESTADO) "
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XS(txtPresentacion.Text) & ", " & XS(txtDosificacion.Text) & ", "
                cSQL = cSQL & cboGrupo.ItemData(cboGrupo.ListIndex) & ","
                If optAdultos.Value = True Then
                    cSQL = cSQL & "'A',"
                Else
                    cSQL = cSQL & "'N',"
                End If
                
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & "'S')"
                Else
                    cSQL = cSQL & "'N')"
                End If
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "  MED_NOMBRE=" & XS(txtNombre.Text)
                cSQL = cSQL & " ,MED_PRESENTACION=" & XS(txtPresentacion.Text)
                cSQL = cSQL & " ,MED_DOSIFICACION=" & XS(txtDosificacion.Text)
                cSQL = cSQL & " ,GRU_CODIGO=" & cboGrupo.ItemData(cboGrupo.ListIndex)
                If optAdultos.Value = True Then
                    cSQL = cSQL & " ,MED_EDAD = 'A'"
                Else
                    cSQL = cSQL & " ,MED_EDAD = 'N'"
                End If
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & " ,MED_ESTADO = 'S'"
                Else
                    cSQL = cSQL & " ,MED_ESTADO = 'N'"
                End If
                cSQL = cSQL & " WHERE MED_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE MED_CODIGO  = " & XN(txtID.Text)
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
    CargoCboGrupo
    
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE MED_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                txtID.Text = rec!MED_CODIGO
                txtNombre.Text = rec!MED_NOMBRE
                'si encontró el registro muestro los datos
                
                txtPresentacion.Text = ChkNull(rec!MED_PRESENTACION)
                txtDosificacion.Text = ChkNull(rec!MED_DOSIFICACION)
                
                Call BuscaCodigoProxItemData(ChkNull(rec!GRU_CODIGO), cboGrupo)
                
                If ChkNull(rec!MED_EDAD) = "N" Then optNiños.Value = True
                If ChkNull(rec!MED_EDAD) = "A" Then optAdultos.Value = True
                
                If ChkNull(rec!MED_ESTADO) = "N" Or ChkNull(rec!MED_ESTADO) = "" Then
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
Private Sub CargoCboGrupo()
    cboGrupo.Clear
    cSQL = "SELECT * FROM GRUPOS ORDER BY GRU_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboGrupo.AddItem Trim(rec!GRU_DESCRI)
          cboGrupo.ItemData(cboGrupo.NewIndex) = rec!GRU_CODIGO
          rec.MoveNext
       Loop
       cboGrupo.ListIndex = cboGrupo.ListIndex + 1
    End If
    rec.Close
End Sub



Private Sub txtPresentacion_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtPresentacion_GotFocus()
    SelecTexto txtPresentacion
End Sub

Private Sub txtPresentacion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDosificacion_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtDosificacion_GotFocus()
    SelecTexto txtDosificacion
End Sub

Private Sub txtDosificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
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

