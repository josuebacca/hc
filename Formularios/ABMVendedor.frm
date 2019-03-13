VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ABMVendedor 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6900
   ClientLeft      =   2700
   ClientTop       =   2625
   ClientWidth     =   7275
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkVenEstado 
      Caption         =   "Dar de Baja"
      Height          =   285
      Left            =   1065
      TabIndex        =   0
      Top             =   5760
      Width           =   1140
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   240
      Picture         =   "ABMVendedor.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6135
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3150
      TabIndex        =   2
      Top             =   6255
      Width           =   1300
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1800
      TabIndex        =   1
      Top             =   6255
      Width           =   1300
   End
   Begin TabDlg.SSTab tabVendedor 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6800
      _ExtentX        =   11986
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "ABMVendedor.frx":0156
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(12)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtID"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtConsul"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboprofesion"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtDomicilio"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTelefono"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboLocalidad"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cboProvincia"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboPais"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtNombre"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtcoseguro"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtPorcentCom"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtMail"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtFax"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Motivos"
      TabPicture(1)   =   "ABMVendedor.frx":0172
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDesasignar"
      Tab(1).Control(1)=   "cmdAsignar"
      Tab(1).Control(2)=   "grdMotivoAsignado"
      Tab(1).Control(3)=   "grdMotivo"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdDesasignar 
         Caption         =   ">"
         Height          =   400
         Left            =   -71880
         TabIndex        =   38
         Top             =   2400
         Width           =   400
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "<"
         Height          =   400
         Left            =   -71880
         TabIndex        =   37
         Top             =   1800
         Width           =   400
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   26
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox txtMail 
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   25
         Top             =   4290
         Width           =   3375
      End
      Begin VB.TextBox txtPorcentCom 
         Height          =   315
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   24
         Top             =   4770
         Width           =   735
      End
      Begin VB.TextBox txtcoseguro 
         Height          =   315
         Left            =   3450
         MaxLength       =   50
         TabIndex        =   23
         Top             =   4785
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   14
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox cboPais 
         Height          =   315
         ItemData        =   "ABMVendedor.frx":018E
         Left            =   1065
         List            =   "ABMVendedor.frx":0190
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1890
         Width           =   3375
      End
      Begin VB.ComboBox cboProvincia 
         Height          =   315
         ItemData        =   "ABMVendedor.frx":0192
         Left            =   1065
         List            =   "ABMVendedor.frx":0194
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2235
         Width           =   3375
      End
      Begin VB.ComboBox cboLocalidad 
         Height          =   315
         ItemData        =   "ABMVendedor.frx":0196
         Left            =   1065
         List            =   "ABMVendedor.frx":0198
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2610
         Width           =   3375
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   10
         Top             =   3375
         Width           =   3375
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2925
         Width           =   3375
      End
      Begin VB.ComboBox cboprofesion 
         Height          =   315
         ItemData        =   "ABMVendedor.frx":019A
         Left            =   1065
         List            =   "ABMVendedor.frx":019C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1170
         Width           =   3375
      End
      Begin VB.TextBox txtConsul 
         Height          =   315
         Left            =   1065
         TabIndex        =   7
         Top             =   1530
         Width           =   720
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   480
         Width           =   720
      End
      Begin MSFlexGridLib.MSFlexGrid grdMotivoAsignado 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   35
         Top             =   960
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid grdMotivo 
         Height          =   4455
         Left            =   -71280
         TabIndex        =   36
         Top             =   960
         Width           =   2800
         _ExtentX        =   4948
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Caption         =   "Motivos Disponibles"
         Height          =   375
         Left            =   -70680
         TabIndex        =   34
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Motivos Asignados"
         Height          =   495
         Left            =   -74280
         TabIndex        =   33
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   32
         Top             =   3885
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   31
         Top             =   4335
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Particular:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   4830
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   11
         Left            =   1800
         TabIndex        =   29
         Top             =   4830
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   12
         Left            =   4200
         TabIndex        =   28
         Top             =   4845
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coseguro:"
         Height          =   195
         Index           =   13
         Left            =   2520
         TabIndex        =   27
         Top             =   4845
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "País:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1935
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Localidad:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   2625
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   3420
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   2970
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ocupacion:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   16
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Consultorio:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id.:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   510
         Width           =   270
      End
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
        ElseIf txtPorcentCom.Text <> "" Then
            If txtPorcentCom.Text > 100 Then
                MsgBox "El porcentaje de comisión no puede ser mayor al 100 % ", vbOKOnly + vbCritical, TIT_MSGBOX
                Exit Function
            End If
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
        
        'Guardar motivos
        Select Case vMode
            Case 1 'nuevo
                If grdMotivoAsignado.Rows > 1 Then
                    GuardarMotivos
                End If
            Case 2 'editaar
                'hay q borrar por mas que este vacia la grilla
                BorrarMotivos (XN(txtID.Text))
                If grdMotivoAsignado.Rows > 1 Then
                    GuardarMotivos
                End If
            Case 4 'Eliminar
                If grdMotivoAsignado.Rows > 1 Then
                    BorrarMotivos (XN(txtID.Text))
                End If
        End Select
        Select Case vMode
            Case 1 'nuevo
            
                cSQL = "INSERT INTO " & cTabla
                cSQL = cSQL & "     (VEN_CODIGO, VEN_NOMBRE, VEN_DOMICI, VEN_TELEFONO,"
                cSQL = cSQL & " VEN_MAIL, VEN_FAX, LOC_CODIGO, PRO_CODIGO, PAI_CODIGO,PR_CODIGO,VEN_ESTADO,VEN_CONSULTORIO,VEN_PORCENTCOM,VEN_COSEGURO) "
                cSQL = cSQL & " VALUES "
                cSQL = cSQL & "     (" & XN(txtID.Text) & ", " & XS(txtNombre.Text) & ", "
                cSQL = cSQL & XS(txtDomicilio.Text) & ", " & XS(txtTelefono.Text) & ", "
                cSQL = cSQL & XS(txtMail.Text) & ", " & XS(txtFax.Text) & ", "
                cSQL = cSQL & cboLocalidad.ItemData(cboLocalidad.ListIndex) & ", "
                cSQL = cSQL & cboProvincia.ItemData(cboProvincia.ListIndex) & ", "
                cSQL = cSQL & cboPais.ItemData(cboPais.ListIndex) & ","
                cSQL = cSQL & cboprofesion.ItemData(cboprofesion.ListIndex) & ","
                If chkVenEstado.Value = Checked Then
                    cSQL = cSQL & "'S'" & ","
                Else
                    cSQL = cSQL & "'N'" & ","
                End If
            cSQL = cSQL & XN(txtConsul.Text) & ","
            cSQL = cSQL & XN(txtPorcentCom.Text) & ","
            cSQL = cSQL & XN(txtcoseguro.Text) & ")"
                
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
                cSQL = cSQL & " ,VEN_PORCENTCOM=" & XN(txtPorcentCom.Text)
                cSQL = cSQL & " ,VEN_COSEGURO=" & XN(txtcoseguro.Text)
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
Private Sub GuardarMotivos()
    Dim i As Integer
    For i = 2 To grdMotivoAsignado.Rows - 1
        cSQL = "INSERT INTO MOTIVO_VENDEDOR "
        cSQL = cSQL & "  (VEN_CODIGO,MOT_CODIGO)"
        cSQL = cSQL & " VALUES "
        cSQL = cSQL & "     (" & XN(txtID.Text) & " , "
        cSQL = cSQL & XN(grdMotivoAsignado.TextMatrix(i, 0)) & " )"
        DBConn.Execute cSQL
        
    Next
End Sub
Private Sub BorrarMotivos(vencod As Integer)
    cSQL = "DELETE FROM MOTIVO_VENDEDOR  WHERE VEN_CODIGO  = " & vencod
    DBConn.Execute cSQL
End Sub


Private Sub cmdAsignar_Click()
    cmdAceptar.Enabled = True
    Dim i As Integer
    i = 2
    Do While i <= grdMotivo.Rows - 1
        If grdMotivo.TextMatrix(i, 2) = "SI" Then
            'Limpio Campo Seleccionado
            grdMotivo.TextMatrix(i, 2) = "NO"
            grdMotivoAsignado.AddItem grdMotivo.TextMatrix(i, 0) & Chr(9) & _
                         grdMotivo.TextMatrix(i, 1) & Chr(9) & _
                         grdMotivo.TextMatrix(i, 2)

            grdMotivo.RemoveItem (i)
        Else
            i = i + 1
        End If
    Loop
    
End Sub

Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 12)
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdDesasignar_Click()
    cmdAceptar.Enabled = True
    Dim i As Integer
    i = 2
    Do While i <= grdMotivoAsignado.Rows - 1
        If grdMotivoAsignado.TextMatrix(i, 2) = "SI" Then
            'limpio campo seleccionado
            grdMotivoAsignado.TextMatrix(i, 2) = "NO"
            grdMotivo.AddItem grdMotivoAsignado.TextMatrix(i, 0) & Chr(9) & _
                         grdMotivoAsignado.TextMatrix(i, 1) & Chr(9) & _
                         grdMotivoAsignado.TextMatrix(i, 2)
            grdMotivoAsignado.RemoveItem (i)
        Else
            i = i + 1
        End If
    Loop
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
    configurogrilla
    CargoCboProfesion
    'CargoGrillaMotivo
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
                txtConsul.Text = ChkNull(rec!VEN_CONSULTORIO)
                txtDomicilio.Text = ChkNull(rec!VEN_DOMICI)
                txtTelefono.Text = ChkNull(rec!VEN_TELEFONO)
                txtFax.Text = ChkNull(rec!VEN_FAX)
                txtMail.Text = ChkNull(rec!VEN_MAIL)
                txtPorcentCom.Text = Format(Chk0(rec!VEN_PORCENTCOM), "#,##0.00")
                txtcoseguro.Text = Format(Chk0(rec!VEN_COSEGURO), "#,##0.00")
                
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
    cargarGrillasMotivoYAsignado
End Sub


Private Sub cargarGrillasMotivoYAsignado()
    'busco motivos asignados y genero coleccion con los numeros de motivo
    Dim motAsignados As New Collection
    Dim i As Integer
    'busco motivos asignados
    cSQL = "SELECT M.MOT_CODIGO FROM MOTIVO M,MOTIVO_VENDEDOR MV"
    cSQL = cSQL & " WHERE  M.MOT_CODIGO= MV.MOT_CODIGO"
    cSQL = cSQL & " AND  MV.VEN_CODIGO= " & txtID.Text
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    'lleno la collection con numeros d emotivos asign
    Do While rec.EOF = False
        motAsignados.Add (rec!MOT_CODIGO)
        rec.MoveNext
    Loop
    rec.Close
    If motAsignados.Count = 0 Then
        'cargo todos los motivos ,y vacia la de asignados
        cargarGrillaMotivo
    Else
         'busco todos los motivos
          sql = "SELECT MOT_CODIGO,MOT_DESCRI FROM MOTIVO"
          rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
         'cargar en las dos grillas
         Do While rec.EOF = False
            'para cada motivo
              For i = 1 To motAsignados.Count
                 If rec!MOT_CODIGO = motAsignados(i) Then 'si es asignado
                     grdMotivoAsignado.AddItem Trim(rec!MOT_CODIGO) & Chr(9) & _
                     Trim(rec!MOT_DESCRI) & Chr(9) & "NO" 'lo agrego a grilla de asignados
                     Exit For
                 End If
                 If i = motAsignados.Count Then
                    grdMotivo.AddItem Trim(rec!MOT_CODIGO) & Chr(9) & _
                    Trim(rec!MOT_DESCRI) & Chr(9) & "NO" 'lo agrego a grilla motivos
                End If
             Next
             rec.MoveNext
         Loop
         rec.Close
    End If
End Sub



Private Sub cargarGrillaMotivo()
    cSQL = "SELECT * FROM MOTIVO"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
       Do While rec.EOF = False
            grdMotivo.AddItem Trim(rec!MOT_CODIGO) & Chr(9) & _
            Trim(rec!MOT_DESCRI) & Chr(9) & "NO"
            rec.MoveNext
       Loop
    End If
    rec.Close
End Sub
Private Sub configurogrilla()
    'motivos asignados
    grdMotivoAsignado.FormatString = "Codigo|Descripcion|Seleccionado"
    grdMotivoAsignado.ColWidth(0) = 0 'codigo
    grdMotivoAsignado.ColWidth(1) = 2400 'descipcion
    grdMotivoAsignado.ColWidth(2) = 0 'selecc
    grdMotivoAsignado.Rows = 1
    grdMotivoAsignado.Cols = 3
    grdMotivoAsignado.BorderStyle = flexBorderNone
    grdMotivoAsignado.row = 0
    Dim i As Integer
    For i = 0 To grdMotivoAsignado.Cols - 1
        grdMotivoAsignado.Col = i
        grdMotivoAsignado.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdMotivoAsignado.CellBackColor = &H808080    'GRIS OSCURO
        grdMotivoAsignado.CellFontBold = True
    Next
    grdMotivoAsignado.AddItem ("" & "" & "")
    'todos los motivos
    grdMotivo.FormatString = "Codigo|Descripcion|Seleccionado"
    grdMotivo.ColWidth(0) = 0 'codigo
    grdMotivo.ColWidth(1) = 2400 'descipcion
    grdMotivo.ColWidth(2) = 0 'selecc
    grdMotivo.Rows = 1
    grdMotivo.Cols = 3
    grdMotivo.BorderStyle = flexBorderNone
    grdMotivo.row = 0
    For i = 0 To grdMotivo.Cols - 1
        grdMotivo.Col = i
        grdMotivo.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdMotivo.CellBackColor = &H808080    'GRIS OSCURO
        grdMotivo.CellFontBold = True
    Next
    grdMotivo.AddItem ("" & "" & "")
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

Private Sub grdMotivo_DblClick()
    Dim J As Integer
    If grdMotivo.TextMatrix(grdMotivo.RowSel, 0) <> "" Then
        If grdMotivo.TextMatrix(grdMotivo.RowSel, 2) = "NO" Then
            grdMotivo.TextMatrix(grdMotivo.RowSel, 2) = "SI"
            'CAMBIAR COLOR
            'backColor = &HC000&
            'foreColor = &HFFFFFF
            For J = 0 To grdMotivo.Cols - 1
                grdMotivo.Col = J
                grdMotivo.CellForeColor = &HFFFFFF
                grdMotivo.CellBackColor = &HC000&
                grdMotivo.CellFontBold = True
            Next
        Else
            grdMotivo.TextMatrix(grdMotivo.RowSel, 2) = "NO"
            For J = 0 To grdMotivo.Cols - 1
                grdMotivo.Col = J
                grdMotivo.CellForeColor = &H80000008
                grdMotivo.CellBackColor = &H80000005
                grdMotivo.CellFontBold = False
            Next
        End If
    End If
End Sub

Private Sub grdMotivoAsignado_DblClick()
    Dim J As Integer
    If grdMotivoAsignado.TextMatrix(grdMotivoAsignado.RowSel, 0) <> "" Then
        If grdMotivoAsignado.TextMatrix(grdMotivoAsignado.RowSel, 2) = "NO" Then
            grdMotivoAsignado.TextMatrix(grdMotivoAsignado.RowSel, 2) = "SI"
            'CAMBIAR COLOR
            'backColor = &HC000&
            'foreColor = &HFFFFFF
            For J = 0 To grdMotivoAsignado.Cols - 1
                grdMotivoAsignado.Col = J
                grdMotivoAsignado.CellForeColor = &HFFFFFF
                grdMotivoAsignado.CellBackColor = &HC000&
                grdMotivoAsignado.CellFontBold = True
            Next
        Else
            grdMotivoAsignado.TextMatrix(grdMotivoAsignado.RowSel, 2) = "NO"
            For J = 0 To grdMotivoAsignado.Cols - 1
                grdMotivoAsignado.Col = J
                grdMotivoAsignado.CellForeColor = &H80000008
                grdMotivoAsignado.CellBackColor = &H80000005
                grdMotivoAsignado.CellFontBold = False
            Next
        End If
    End If
End Sub

Private Sub txtConsul_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtcoseguro_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtcoseguro_GotFocus()
    seltxt
End Sub

Private Sub txtcoseguro_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtcoseguro, KeyAscii)
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

Private Sub txtPorcentCom_Change()
    cmdAceptar.Enabled = True
End Sub

Private Sub txtPorcentCom_GotFocus()
    SelecTexto txtPorcentCom
End Sub

Private Sub txtPorcentCom_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentCom, KeyAscii)
End Sub

Private Sub txtPorcentCom_LostFocus()
    If txtPorcentCom.Text <> "" Then
        txtPorcentCom.Text = Valido_Importe(txtPorcentCom.Text)
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
