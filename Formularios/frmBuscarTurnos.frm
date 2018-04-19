VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBuscarTurnos 
   Caption         =   "Buscar Turnos"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11010
   Icon            =   "frmBuscarTurnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   9840
      Picture         =   "frmBuscarTurnos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   8880
      Picture         =   "frmBuscarTurnos.frx":190C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "Buscar por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   30
      TabIndex        =   14
      Top             =   0
      Width           =   10875
      Begin MSComCtl2.DTPicker DTFechaDesde 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   650
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   110034945
         CurrentDate     =   40073
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         MaxLength       =   40
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CheckBox chkPaciente 
         Caption         =   "Paciente"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkHoras 
         Caption         =   "Horas"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkFechas 
         Caption         =   "Fechas"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkDoctor 
         Caption         =   "Doctor"
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboDoctor 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   255
         Width           =   2400
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar Turnos"
         Height          =   930
         Left            =   9000
         MaskColor       =   &H80000006&
         Picture         =   "frmBuscarTurnos.frx":294E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Buscar "
         Top             =   555
         UseMaskColor    =   -1  'True
         Width           =   1245
      End
      Begin VB.TextBox txtDesCli 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4185
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "Descripción"
         Top             =   1020
         Width           =   3810
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1020
         Width           =   1110
      End
      Begin VB.ComboBox cboDesde 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker DTFechaHasta 
         Height          =   315
         Left            =   6840
         TabIndex        =   4
         Top             =   650
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   110034945
         CurrentDate     =   40073
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Desde:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   19
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lbltipoFac 
         AutoSize        =   -1  'True
         Caption         =   "Doctor:"
         Height          =   195
         Left            =   2415
         TabIndex        =   18
         Top             =   285
         Width           =   525
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   5820
         TabIndex        =   17
         Top             =   660
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1950
         TabIndex        =   16
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Paciente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2265
         TabIndex        =   15
         Top             =   1050
         Width           =   675
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   4550
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   8017
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   12632319
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBuscarTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim J As Integer
Dim hDesde As Integer
Dim hHasta As Integer

Private Sub CmdBuscAprox_Click()
    Dim sColor As String
    Dim USUARIO As String
    sql = "SELECT T.*,V.VEN_NOMBRE,C.CLI_RAZSOC"
    sql = sql & " FROM TURNOS T, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND T.VEN_CODIGO = V.VEN_CODIGO"
    If txtCliente.Text <> "" Then
        sql = sql & " AND T.CLI_CODIGO = " & XN(txtCodigo)
    End If
    If cboDoctor.ListIndex <> -1 Then
        sql = sql & " AND T.VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    If DTFechaDesde.Value <> "" Then sql = sql & " AND T.TUR_FECHA>=" & XDQ(DTFechaDesde.Value)
    If DTFechaHasta.Value <> "" Then sql = sql & " AND T.TUR_FECHA<=" & XDQ(DTFechaHasta.Value)
    If cboDesde.ListIndex <> -1 Then
        sql = sql & " AND T.TUR_HORAD = " & cboDesde.ItemData(cboDesde.ListIndex)
    End If
    sql = sql & " ORDER BY TUR_FECHA,TUR_HORAD"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    GrdModulos.Rows = 1
    i = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
            If Chk0(rec!TUR_USER) <> 0 Then
                USUARIO = BuscarUser(ChkNull(rec!TUR_USER))
            End If
            GrdModulos.AddItem rec!TUR_FECHA & Chr(9) & Format(rec!TUR_HORAD, "hh:mm") & Chr(9) & _
                               Format(rec!TUR_HORAH, "hh:mm") & Chr(9) & rec!CLI_RAZSOC & Chr(9) & _
                               rec!TUR_MOTIVO & Chr(9) & rec!VEN_NOMBRE & Chr(9) & _
                               rec!CLI_CODIGO & Chr(9) & rec!VEN_CODIGO & Chr(9) & _
                               ChkNull(rec!TUR_FECALTA) & Chr(9) & USUARIO
                                                              
            Select Case rec!TUR_ASISTIO
            Case 0
                sColor = &H80000008
            Case 1
                sColor = &HC000&
            Case 2
                sColor = &HFF&
            End Select
            GrdModulos.row = i
            For J = 0 To GrdModulos.Cols - 1
                GrdModulos.Col = J
                GrdModulos.CellForeColor = sColor          'FUENTE COLOR NEGRO
                GrdModulos.CellBackColor = &HC0C0FF      'ROSA
                GrdModulos.CellFontBold = True
            Next
            i = i + 1
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub
Private Function BuscarUser(Codigo As Integer) As String
    If Codigo = 99 Then
        BuscarUser = "GISELA BOTTI"
    Else
        sql = "SELECT VEN_NOMBRE FROM VENDEDOR WHERE VEN_CODIGO = " & Codigo
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscarUser = Rec1!VEN_NOMBRE
        End If
        Rec1.Close
    End If
End Function
Private Sub CmdNuevo_Click()
    chkDoctor.Value = Unchecked
    chkFechas.Value = Unchecked
    chkPaciente.Value = Unchecked
    chkFechas.Value = Unchecked
    cboDoctor.ListIndex = -1
    DTFechaDesde.Value = ""
    DTFechaHasta.Value = ""
    txtCliente.Text = ""
    txtDesCli.Text = ""
    cboDesde.ListIndex = -1
    GrdModulos.Rows = 1
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmBuscarTurnos = Nothing
        'Set rec = Nothing
        'Set Rec1 = Nothing
        'Set Rec2 = Nothing
        Unload Me
    End If
End Sub

Private Sub chkDoctor_Click()
    If cboDoctor.ListIndex <> -1 Then
        cboDoctor.ListIndex = -1
    Else
        cboDoctor.ListIndex = 0
        cboDoctor.SetFocus
    End If
End Sub

Private Sub chkFechas_Click()
    If DTFechaDesde.Value <> "" Or DTFechaHasta.Value <> "" Then
        DTFechaDesde.Value = ""
        DTFechaHasta.Value = ""
    Else
        DTFechaDesde.Value = Date
        DTFechaDesde.SetFocus
    End If
End Sub

Private Sub chkHoras_Click()
    If cboDesde.ListIndex <> -1 Then
        cboDesde.ListIndex = -1
    Else
        cboDesde.ListIndex = 0
        cboDesde.SetFocus
    End If
End Sub

Private Sub chkPaciente_Click()
    If txtCliente.Text <> "" Then
        txtCliente.Text = ""
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If

End Sub

Private Sub Form_Activate()
    If txtCliente.Text <> "" Then
        DTFechaDesde.Value = Date
        txtCliente_LostFocus
        CmdBuscAprox_Click
    End If
    Centrar_pantalla Me
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
    Set rec = New ADODB.Recordset
    DTFechaDesde.Value = Null
    DTFechaHasta.Value = Null
    LlenarComboDoctor
    LlenarComboHoras
    
    'GRILLA BUSQUEDA
    GrdModulos.FormatString = "Fecha|^Hora Desde|^Hora Hasta|Paciente|Motivo|Doctor|CodPaciente|CodDoctor|Fec Alta|Usuario"
    GrdModulos.ColWidth(0) = 1300 'FECHA
    GrdModulos.ColWidth(1) = 1300 'HORA DESDE
    GrdModulos.ColWidth(2) = 1300 'HORA HASTA
    GrdModulos.ColWidth(3) = 2500 'PACIENTE
    GrdModulos.ColWidth(4) = 2200    'MOTIVO
    GrdModulos.ColWidth(5) = 2200 'DOCTOR
    GrdModulos.ColWidth(6) = 0 'COD PACIENTE
    GrdModulos.ColWidth(7) = 0 'COD DOCTOR
    GrdModulos.ColWidth(8) = 1300 'Fecha alta
    GrdModulos.ColWidth(9) = 2500 'Usuario
    GrdModulos.Cols = 10
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    Centrar_pantalla Me

End Sub

Private Sub LlenarComboDoctor()
    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO =1"
    sql = sql & " ORDER BY VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        'cboFactura1.AddItem "(Todas)"
        Do While rec.EOF = False
            cboDoctor.AddItem rec!VEN_NOMBRE
            cboDoctor.ItemData(cboDoctor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        cboDoctor.ListIndex = -1
    End If
    rec.Close
End Sub
Private Sub LlenarComboHoras()
    Dim cItems As Integer
    rec.Open "SELECT HS_DESDE,HS_HASTA FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        hDesde = Hour(rec!HS_DESDE)
        hHasta = Hour(rec!HS_HASTA)
    End If
    rec.Close
    cItems = (hHasta - hDesde) * 2 + 1
    i = 0
    For J = hDesde To hHasta
        cboDesde.AddItem Format(J, "00") & ":00 "
        cboDesde.ItemData(cboDesde.NewIndex) = i
        cboDesde.AddItem Format(J, "00") & ":30 "
        cboDesde.ItemData(cboDesde.NewIndex) = i + 0.5
        
        'cbohasta.AddItem Format(J, "00") & ":00 "
        'cbohasta.ItemData(cbohasta.NewIndex) = i
        'cbohasta.AddItem Format(J, "00") & ":30 "
        'cbohasta.ItemData(cbohasta.NewIndex) = i + 0.5
        
        i = i + 1
    Next
    cboDesde.ListIndex = -1
    'cbohasta.ListIndex = -1
    
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
        txtCodigo.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCliente", "CODIGO"
    End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtCliente.Text <> "" Then
            sql = sql & " CLI_CODIGO=" & XN(txtCliente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtCliente) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
            txtCodigo.Text = rec!CLI_CODIGO
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub
Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
        txtCodigo.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtCliente", "CODIGO"
    End If
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtCliente.Text <> "" Then
            sql = sql & " CLI_DNI=" & XN(txtCliente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtCliente) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtCliente", "CADENA", Trim(txtDesCli.Text)
                If rec.State = 1 Then rec.Close
                txtDesCli.SetFocus
            Else
                txtCliente.Text = rec!CLI_DNI
                txtDesCli.Text = rec!CLI_RAZSOC
                txtCodigo.Text = rec!CLI_CODIGO
            End If
        Else
            MsgBox "No se encontro el Paciente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub
Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO,CLI_NRODOC"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código, dni"
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .Field = "CLI_NRODOC"
        campo3 = .Field
        
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtcodCli" Then
                txtCliente.Text = .ResultFields(2)
                txtCliente_LostFocus
            Else
                txtCliente.Text = .ResultFields(2)
                txtCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub
