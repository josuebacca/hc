VERSION 5.00
Begin VB.Form ABMProtocolos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Protocolo..."
   ClientHeight    =   8430
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMProtocolos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsig 
      Caption         =   ">"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "<"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton cmdult 
      Caption         =   ">>"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   7800
      Width           =   495
   End
   Begin VB.CommandButton cmdpri 
      Caption         =   "<<"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox txtAbrevia 
      Height          =   5820
      Index           =   5
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   9555
   End
   Begin VB.TextBox txtAbrevia 
      Height          =   5820
      Index           =   4
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   9555
   End
   Begin VB.TextBox txtAbrevia 
      Height          =   5820
      Index           =   3
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   9555
   End
   Begin VB.TextBox txtAbrevia 
      Height          =   5820
      Index           =   2
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   9555
   End
   Begin VB.TextBox txtAbrevia 
      Height          =   5820
      Index           =   1
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   9555
   End
   Begin VB.TextBox txtAbrevia 
      Height          =   5820
      Index           =   0
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   9555
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   6240
      Picture         =   "ABMProtocolos.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   330
   End
   Begin VB.TextBox txtDescri 
      Height          =   300
      Left            =   210
      TabIndex        =   1
      Top             =   1035
      Width           =   9435
   End
   Begin VB.TextBox txtID 
      Height          =   300
      Left            =   210
      TabIndex        =   0
      Top             =   390
      Width           =   720
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   8160
      TabIndex        =   4
      Top             =   8025
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   6825
      TabIndex        =   3
      Top             =   8025
      Width           =   1300
   End
   Begin VB.Label lblnroja 
      AutoSize        =   -1  'True
      Caption         =   "Label16"
      Height          =   195
      Left            =   4680
      TabIndex        =   18
      Top             =   7560
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contenido:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1455
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   810
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Id.:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   165
      Width           =   270
   End
End
Attribute VB_Name = "ABMProtocolos"
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
Const cTabla = "TIPO_IMAGEN"
Const cCampoID = "TIP_CODIGO"
Const cDesRegistro = "Protocolo"

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
            'AcCtrl txtAbrevia
        Case 3, 4
            DesacCtrl txtDescri
            'DesacCtrl txtAbrevia
    End Select
    
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nuevo Protocolo..."
            txtID_LostFocus
            DesacCtrl txtID
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Protocolo..."
            DesacCtrl txtID
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos del Protocolo..."
            DesacCtrl txtID
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Protocolo..."
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
                             "Ingrese la Identificación del Protocolo antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtDescri.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el nombre del Protocolo antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtDescri.SetFocus
            Exit Function
        ElseIf txtAbrevia(0).Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el contenido del Protocolo antes de aceptar.", vbCritical + vbOKOnly, App.Title
            'txtAbrevia(0).SetFocus
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
                cSQL = cSQL & "     (TIP_CODIGO, TIP_NOMBRE, TIP_CONTEN, "
                cSQL = cSQL & "     ,TIP_CONTEN1, TIP_CONTEN2, TIP_CONTEN3, "
                cSQL = cSQL & "     ,TIP_CONTEN4, TIP_CONTEN5) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtDescri.Text) & ", "
                cSQL = cSQL & XS(txtAbrevia(0).Text) & ", " & XS(txtAbrevia(1).Text) & ", "
                cSQL = cSQL & XS(txtAbrevia(2).Text) & ", " & XS(txtAbrevia(3).Text) & ", "
                cSQL = cSQL & XS(txtAbrevia(4).Text) & ", " & XS(txtAbrevia(5).Text) & ") "
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     TIP_NOMBRE = " & XS(txtDescri.Text)
                cSQL = cSQL & "    ,TIP_CONTEN = " & XS(txtAbrevia(0).Text)
                cSQL = cSQL & "    ,TIP_CONTEN1 = " & XS(txtAbrevia(1).Text)
                cSQL = cSQL & "    ,TIP_CONTEN2 = " & XS(txtAbrevia(2).Text)
                cSQL = cSQL & "    ,TIP_CONTEN3 = " & XS(txtAbrevia(3).Text)
                cSQL = cSQL & "    ,TIP_CONTEN4 = " & XS(txtAbrevia(4).Text)
                cSQL = cSQL & "    ,TIP_CONTEN5 = " & XS(txtAbrevia(5).Text)
                cSQL = cSQL & " WHERE TIP_CODIGO  = " & XN(txtID.Text)
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla & " WHERE TIP_CODIGO  = " & XN(txtID.Text)
            
        End Select
        
        DBConn.Execute cSQL
        DBConn.CommitTrans
        
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

Private Sub cmdprev_Click()
    If hojaactual > 0 Then
        lblnroja.Caption = "Hoja " & hojaactual
         muestro_ImgDescri hojaactual
    End If
End Sub

Private Sub cmdpri_Click()
    lblnroja.Caption = "Hoja 1"
    muestro_ImgDescri 1
    
End Sub

Private Sub cmdsig_Click()
 If hojaactual < 5 Then
    lblnroja.Caption = "Hoja " & hojaactual + 2
    muestro_ImgDescri hojaactual + 2
 End If
End Sub

Private Sub cmdult_Click()
    lblnroja.Caption = "Hoja 6"
    muestro_ImgDescri 6
End Sub

Private Sub Form_Activate()
    'hizo click en una columna no correcta
    If vMode = 2 And vFieldID = "0" Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
    
    
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
            cSQL = "SELECT * FROM " & cTabla & "  WHERE TIP_CODIGO = " & Mid(vFieldID, 2, Len(vFieldID) - 2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!TIP_CODIGO
                txtDescri.Text = rec!TIP_NOMBRE
                txtAbrevia(0).Text = ChkNull(rec!TIP_CONTEN)
                txtAbrevia(1).Text = ChkNull(rec!TIP_CONTEN1)
                txtAbrevia(2).Text = ChkNull(rec!TIP_CONTEN2)
                txtAbrevia(3).Text = ChkNull(rec!TIP_CONTEN3)
                txtAbrevia(4).Text = ChkNull(rec!TIP_CONTEN4)
                txtAbrevia(5).Text = ChkNull(rec!TIP_CONTEN5)
                muestro_ImgDescri 1
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    lblnroja.Caption = "Hoja 1"
    muestro_ImgDescri 1
    
    'establesco funcionalidad del form de datos
    SetMode vMode
End Sub



Private Sub txtAbrevia_Change(Index As Integer)
    cmdAceptar.Enabled = True
End Sub




Private Sub txtAbrevia_GotFocus(Index As Integer)
    seltxt txtAbrevia(Index)
End Sub

Private Sub txtAbrevia_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtdescri_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtdescri_GotFocus()
    seltxt
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
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
                                     "El código ingresado pertenece a otro registro de " & cDesRegistro & ".", vbCritical + vbOKOnly, App.Title
                    txtID.Text = ""
                    txtID.SetFocus
                End If
            End If
        End If
    End If
End Sub
Private Function muestro_ImgDescri(hoja As Integer)
    'hoja va de 1 a 6
    'elemento de matriz va de 0 a 5
    Dim i As Integer
        For i = 0 To 5
            If (hoja - 1) = i Then
                txtAbrevia(i).Visible = True
            Else
                txtAbrevia(i).Visible = False
            End If
        Next
End Function
Private Function hojaactual() As Integer
    Dim i As Integer
    For i = 0 To 5
        If txtAbrevia(i).Visible = True Then
            hojaactual = i
        End If
    Next

End Function

