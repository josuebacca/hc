VERSION 5.00
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form ABMCuentaBancaria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Cuenta Bancaria..."
   ClientHeight    =   2700
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
   Icon            =   "ABMCuentaBancaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin FechaCtl.Fecha txtfechaApertura 
      Height          =   345
      Left            =   1335
      TabIndex        =   3
      Top             =   1170
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.TextBox txtdescri 
      Height          =   315
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1890
      Width           =   4920
   End
   Begin VB.TextBox txtSaldoActual 
      Height          =   315
      Left            =   3930
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1530
      Width           =   1245
   End
   Begin VB.TextBox txtSaldoInicial 
      Height          =   315
      Left            =   3930
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1170
      Width           =   1245
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      ItemData        =   "ABMCuentaBancaria.frx":000C
      Left            =   1335
      List            =   "ABMCuentaBancaria.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3855
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   90
      Picture         =   "ABMCuentaBancaria.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2295
      Width           =   330
   End
   Begin VB.TextBox txtID 
      Height          =   315
      Left            =   1335
      MaxLength       =   10
      TabIndex        =   1
      Top             =   450
      Width           =   1245
   End
   Begin VB.ComboBox cboTipoCuenta 
      Height          =   315
      ItemData        =   "ABMCuentaBancaria.frx":015A
      Left            =   1335
      List            =   "ABMCuentaBancaria.frx":015C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   3855
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   4935
      TabIndex        =   9
      Top             =   2295
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3555
      TabIndex        =   8
      Top             =   2295
      Width           =   1300
   End
   Begin FechaCtl.Fecha txtFechaCierre 
      Height          =   345
      Left            =   1335
      TabIndex        =   5
      Top             =   1530
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   195
      Index           =   5
      Left            =   75
      TabIndex        =   18
      Top             =   1920
      Width           =   870
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Actual:"
      Height          =   195
      Index           =   4
      Left            =   2850
      TabIndex        =   17
      Top             =   1575
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Cierre:"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   16
      Top             =   1575
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Inicial:"
      Height          =   195
      Index           =   2
      Left            =   2850
      TabIndex        =   15
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Cuenta:"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   13
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Apertura:"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   11
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cuenta:"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   10
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "ABMCuentaBancaria"
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

'constantes para funcionalidad de uso del formulario
Const cSugerirID = True 'si es True si sugiere un identificador cuando deja el campo en blanco
Const cTabla = "CTA_BANCARIA"
Const cCampoID = "BAN_CODINT"
Const cDesRegistro = "Cuenta Bancaria"

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
            AcCtrl cboTipoCuenta
            'AcCtrl txtfechaApertura
            'AcCtrl txtSaldoActual
            'AcCtrl txtSaldoInicial
            'AcCtrl txtFechaCierre
            'AcCtrl txtdescri
        Case 3, 4
            DesacCtrl cboTipoCuenta
            'DesacCtrl txtfechaApertura
            'DesacCtrl txtSaldoActual
            'DesacCtrl txtSaldoInicial
            'DesacCtrl txtFechaCierre
            'DesacCtrl txtdescri
    End Select
    
    Select Case pMode
        Case 1
            cmdAceptar.Enabled = False
            Me.Caption = "Nueva Cuenta Bancaria..."
            AcCtrl txtID
            AcCtrl cboBanco
        Case 2
            cmdAceptar.Enabled = False
            Me.Caption = "Editando Cuenta Bancaria..."
            DesacCtrl txtID
            DesacCtrl cboBanco
        Case 3
            cmdAceptar.Visible = False
            Me.Caption = "Datos de la Cuenta Bancaria..."
            DesacCtrl txtID
            DesacCtrl cboBanco
        Case 4
            cmdAceptar.Enabled = True
            Me.Caption = "Eliminando Empresa..."
            DesacCtrl txtID
            DesacCtrl cboBanco
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
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese el Número de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtID.SetFocus
            Exit Function
        ElseIf txtfechaApertura.Text = "" Then
            Beep
            MsgBox "Falta información." & Chr(13) & _
                             "Ingrese la Fecha de Apertura de " & cDesRegistro & " antes de aceptar.", vbCritical + vbOKOnly, App.Title
            txtfechaApertura.SetFocus
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
                cSQL = cSQL & "     (CTA_NROCTA, BAN_CODINT, CTA_FECAPE, CTA_SALINI, CTA_SALACT, CTA_FECCIE, TCU_CODIGO, CTA_DESCRI) "
                cSQL = cSQL & "VALUES "
                cSQL = cSQL & "     (" & XS(txtID.Text) & ", " & XN(cboBanco.ItemData(cboBanco.ListIndex)) & ", "
                cSQL = cSQL & XD(txtfechaApertura.Text) & ", " & XN(txtSaldoInicial.Text) & ", "
                cSQL = cSQL & XN(txtSaldoActual.Text) & ", " & XD(txtFechaCierre.Text) & ", "
                cSQL = cSQL & XN(cboTipoCuenta.ItemData(cboTipoCuenta.ListIndex)) & ", " & XS(txtDescri.Text) & ")"
            
            Case 2 'editar
                
                cSQL = "UPDATE " & cTabla & " SET "
                cSQL = cSQL & "     CTA_FECAPE=" & XD(txtfechaApertura.Text) & ", CTA_SALINI=" & XN(txtSaldoInicial.Text)
                cSQL = cSQL & ", CTA_SALACT=" & XN(txtSaldoActual.Text) & ", CTA_FECCIE=" & XD(txtFechaCierre.Text)
                cSQL = cSQL & ", TCU_CODIGO=" & XN(cboTipoCuenta.ItemData(cboTipoCuenta.ListIndex)) & ", CTA_DESCRI = " & XS(txtDescri.Text)
                cSQL = cSQL & "  WHERE CTA_NROCTA =" & XN(txtID.Text)
                cSQL = cSQL & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
            
            Case 4 'eliminar
            
                cSQL = "DELETE FROM " & cTabla
                cSQL = cSQL & "  WHERE CTA_NROCTA =" & XN(txtID.Text)
                cSQL = cSQL & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
            
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
    
    'cargo el combo de tipo CUENTAS
    cboTipoCuenta.Clear
    cSQL = "SELECT * FROM TIPO_CUENTA order by TCU_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboTipoCuenta.AddItem Trim(rec!TCU_DESCRI)
          cboTipoCuenta.ItemData(cboTipoCuenta.NewIndex) = rec!TCU_CODIGO
          rec.MoveNext
       Loop
       cboTipoCuenta.ListIndex = cboTipoCuenta.ListIndex + 1
    End If
    rec.Close
    'cargo el combo de BANCOS
    cboBanco.Clear
    cSQL = "SELECT * FROM BANCO order by BAN_DESCRI"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
          cboBanco.AddItem Trim(rec!BAN_DESCRI)
          cboBanco.ItemData(cboBanco.NewIndex) = rec!BAN_CODINT
          rec.MoveNext
       Loop
       cboBanco.ListIndex = cboBanco.ListIndex + 1
    End If
    rec.Close
    
       
    If vMode <> 1 Then
        If vFieldID <> "0" Then
            cSQL = "SELECT * FROM " & cTabla & "  WHERE BAN_CODINT = " & Trim(Mid(vFieldID1, 1, 10))
            cSQL = cSQL & " AND CTA_NROCTA = " & Trim(Mid(vFieldID, 1, 10)) '& frmCListaBaseABM.lstvLista.ListItems(2)
            rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
            If (rec.BOF And rec.EOF) = 0 Then
                'si encontró el registro muestro los datos
                txtID.Text = rec!CTA_NROCTA
                Call BuscaCodigoProxItemData(CLng(rec!BAN_CODINT), cboBanco)
                Call BuscaCodigoProxItemData(CLng(rec!TCU_CODIGO), cboTipoCuenta)
                txtfechaApertura.Text = rec!CTA_FECAPE
                txtSaldoInicial.Text = ChkNull(rec!CTA_SALINI)
                txtSaldoActual.Text = ChkNull(rec!CTA_SALACT)
                txtFechaCierre.Text = ChkNull(rec!CTA_FECCIE)
                txtDescri.Text = ChkNull(rec!CTA_DESCRI)
            Else
                Beep
                MsgBox "Imposible encontrar el registro seleccionado.", vbCritical + vbOKOnly, App.Title
            End If
        End If
    End If
    
    'establesco funcionalidad del form de datos
    SetMode vMode
    
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

Private Sub txtSaldoActual_GotFocus()
    seltxt
End Sub

Private Sub txtSaldoActual_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSaldoActual, KeyAscii)
End Sub

Private Sub txtSaldoInicial_GotFocus()
    seltxt
End Sub

Private Sub txtSaldoInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSaldoInicial, KeyAscii)
End Sub
