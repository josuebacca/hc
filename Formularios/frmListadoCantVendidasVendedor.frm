VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoCantVendidasVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Atenciones Medicas"
   ClientHeight    =   3075
   ClientLeft      =   1515
   ClientTop       =   1740
   ClientWidth     =   5520
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListadoCantVendidasVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2520
      TabIndex        =   7
      Top             =   2685
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3825
      TabIndex        =   8
      Top             =   2685
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listar por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   45
      TabIndex        =   26
      Top             =   30
      Width           =   5415
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cboOSocial 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   3555
      End
      Begin VB.TextBox txtBuscaOS 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtBuscarOSNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2820
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "Descripción"
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtBuscarCliDescri 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2790
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Descripción"
         Top             =   1200
         Width           =   2355
      End
      Begin VB.TextBox txtBuscaCliente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1200
         Width           =   1155
      End
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   405
         Width           =   3555
      End
      Begin MSComCtl2.DTPicker FechaDesde 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54853633
         CurrentDate     =   43233
      End
      Begin MSComCtl2.DTPicker FechaHasta 
         Height          =   315
         Left            =   3720
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54853633
         CurrentDate     =   43233
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Paciente"
         Height          =   195
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Obra Social"
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Doctor"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   28
         Top             =   1740
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   1740
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   6735
      TabIndex        =   16
      Top             =   210
      Visible         =   0   'False
      Width           =   6915
      Begin VB.TextBox txtEmpresaCuit 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   24
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox txtEmp_Id 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4365
         TabIndex        =   23
         Top             =   1065
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   75
         Left            =   90
         TabIndex        =   22
         Top             =   1560
         Width           =   6795
      End
      Begin VB.TextBox txtTipoLibro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4365
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEmpresa 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   20
         Top             =   375
         Width           =   3075
      End
      Begin VB.TextBox txtMes_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         MaxLength       =   2
         TabIndex        =   9
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtAnio_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   10
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txtLibro_IdI 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4380
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "C.U.I.T."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1005
         Width           =   540
      End
   End
   Begin VB.Frame fraImpresion 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         Picture         =   "frmListadoCantVendidasVendedor.frx":0442
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         Picture         =   "frmListadoCantVendidasVendedor.frx":0544
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "frmListadoCantVendidasVendedor.frx":0646
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoCantVendidasVendedor.frx":0748
         Left            =   450
         List            =   "frmListadoCantVendidasVendedor.frx":0755
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   270
         Width           =   1635
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2340
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Modo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6690
      TabIndex        =   12
      Top             =   2985
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmListadoCantVendidasVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cboListar_Click()
    If frmListadoCantVendidasVendedor.Visible = True Then
        If cboListar.ListIndex = 0 Then
            cboAgrupar.Enabled = True
            cboAgrupar.ListIndex = 0
        Else
            cboAgrupar.Enabled = False
            cboAgrupar.ListIndex = -1
        End If
    End If
End Sub
Private Sub BuscarAtenciones()
    Dim i As Integer
    'LIMIPIO LA TABLA TEMPORAL DE TURNOS
    sql = "DELETE FROM TMP_TURNOS"
    DBConn.Execute sql
        
    sql = "SELECT T.*,V.VEN_CODIGO,V.VEN_NOMBRE,C.CLI_RAZSOC,C.CLI_NRODOC,C.CLI_TELEFONO,C.CLI_CELULAR,C.CLI_EDAD"
    sql = sql & " FROM TURNOS T, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND T.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND T.TUR_ASISTIO = 1" 'BUSCAMOS LOS TURNOS QUE ASISTIERON
    If FechaDesde.Value <> "" Then
        sql = sql & " AND T.TUR_FECHA >= " & XDQ(FechaDesde.Value)
    End If
    If FechaHasta.Value <> "" Then
        sql = sql & " AND T.TUR_FECHA <= " & XDQ(FechaHasta.Value)
    End If
    If cboVendedor.List(cboVendedor.ListIndex) <> "(Todos)" Then
        sql = sql & " AND T.VEN_CODIGO = " & cboVendedor.ItemData(cboVendedor.ListIndex)
    End If
    If cboOSocial.List(cboOSocial.ListIndex) <> "(Todos)" Then
        sql = sql & " AND T.TUR_OSOCIAL LIKE '" & cboOSocial.Text & "'"
    End If
    If txtCodigo.Text <> "" Then
        sql = sql & " AND T.CLI_CODIGO LIKE " & txtCodigo.Text
    End If
    sql = sql & " ORDER BY T.TUR_FECHA,T.TUR_HORAD"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        i = 1
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_TURNOS "
            sql = sql & " (TMP_ID,TMP_HORA,TMP_FECHA,TMP_DOCTOR,TMP_PACIENTE,TMP_EDAD,TMP_TELEFONO,TMP_CELULAR,TMP_OSOCIAL,TMP_MOTIVO,TMP_DRSOLICITA,TMP_IMPORTE,VEN_CODIGO)"
            sql = sql & " VALUES ( "
            sql = sql & i & ","
            sql = sql & XS(Format(rec!TUR_HORAD, "hh:mm") & " a " & Format(rec!TUR_HORAH, "hh:mm")) & ","
            sql = sql & XDQ(rec!TUR_FECHA) & ","
            sql = sql & XS(ChkNull(rec!VEN_NOMBRE)) & ","
            sql = sql & XS(ChkNull(rec!CLI_RAZSOC)) & "," 'PACIENTE
            sql = sql & XN(Chk0(rec!CLI_EDAD)) & "," ' EDAD
            sql = sql & XS(ChkNull(rec!CLI_TELEFONO)) & "," ' TELEFONO
            sql = sql & XS(ChkNull(rec!CLI_CELULAR)) & "," ' CELU
            sql = sql & XS(ChkNull(rec!TUR_OSOCIAL)) & "," ' O SOCIAL
            sql = sql & XS(ChkNull(rec!TUR_MOTIVO)) & "," ' MOTIVO
            sql = sql & XS(ChkNull(rec!TUR_DRSOLICITA)) & "," ' DR SOLICITA
            sql = sql & XN(Chk0(rec!TUR_IMPORTE)) & "," ' IMPORTE TURNO
            sql = sql & XN(Chk0(rec!VEN_CODIGO)) & ")" ' VENDEDOR CODIGO
            DBConn.Execute sql
        
            'grdGrilla.AddItem Format Rec!CLI_RAZSOC & Chr(9) & edad & Chr(9) & ChkNull(Rec!CLI_TELEFONO) & Chr(9) & ChkNull(Rec!CLI_CELULAR) & Chr(9) & Rec!TUR_OSOCIAL & Chr(9) & ChkNull(Rec!TUR_MOTIVO) & Chr(9) & _
                                     ChkNull(Rec!TUR_DRSOLICITA) & Chr(9) &  Rec!CLI_CODIGO & Chr(9) & Rec!TUR_ASISTIO & Chr(9) & ChkNull(Rec!CLI_NRODOC) & Chr(9) & ChkNull(Rec!TUR_DESDE) & Chr(9) & Rec!TUR_TIENEMUTUAL & Chr(9) & Format(Chk0(Rec!TUR_IMPORTE), "#,##0.00")
                
            i = i + 1
            rec.MoveNext
        Loop
    End If
    
    rec.Close


End Sub
    
    
    
    
    
Private Sub Check1_Click()

End Sub

Private Sub cmdAceptar_Click()

    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
    BuscarAtenciones
    
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    

    
    If FechaDesde.Value <> "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value <> "" And FechaHasta.Value = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Value & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Value & "'"
    ElseIf FechaDesde.Value = "" And FechaHasta.Value = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    If cboOSocial.List(cboOSocial.ListIndex) <> "(Todos)" Then
        Rep.Formulas(1) = "OBRA_SOCIAL='" & " Obra Social: " & cboOSocial.Text & "'"
    Else
        Rep.Formulas(1) = "OBRA_SOCIAL='" & " Obra Social: " & " Todas " & "'"
    End If
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Listado de Atenciones Medicas"
    Rep.ReportFileName = DirReport & "rptresumenesmedicos.rpt"
    Rep.Action = 1
End Sub

Private Sub cmdCancelar_Click()
    Set frmListadoCantVendidasVendedor = Nothing
    'mQuienLlamo = "ABMProducto"
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    Me.Top = 0
    Me.Left = 0
    'cboVendedor.AddItem "(Todos)"
    'CargoComboBox cboVendedor, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    LlenarComboDoctor
    LlenarComboOSocial
    If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 0
    If cboOSocial.ListCount > 0 Then cboOSocial.ListIndex = 0
    'CARGO COMBO LINEA
    cboDestino.ListIndex = 0
    'FechaDesde.value = Date
    'FechaHasta.value = Date
End Sub
Private Sub LlenarComboDoctor()

    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO > 1"
    sql = sql & " ORDER BY VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboVendedor.AddItem "(Todos)"
        Do While rec.EOF = False
            cboVendedor.AddItem rec!VEN_NOMBRE
            cboVendedor.ItemData(cboVendedor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop

    End If
    rec.Close
End Sub
Private Sub LlenarComboOSocial()
    cboOSocial.Clear
    sql = "SELECT * FROM OBRA_SOCIAL"
    sql = sql & " ORDER BY OS_NUMERO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboOSocial.AddItem "(Todos)"
        Do While rec.EOF = False
            cboOSocial.AddItem rec!OS_NOMBRE
            cboOSocial.ItemData(cboOSocial.NewIndex) = rec!OS_NUMERO
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub
Private Sub txtBuscaCliente_GotFocus()
    SelecTexto txtBuscaCliente
End Sub

Private Sub txtBuscaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
        'ActivoGrid = 0
    End If
End Sub

Private Sub txtBuscaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBuscaCliente_LostFocus()
    If txtBuscaCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,CLI_NROAFIL,CLI_CUMPLE,CLI_EDAD"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            If Len(Trim(txtBuscaCliente.Text)) < 7 Then
                sql = sql & " CLI_CODIGO=" & XN(txtCodigo)
            Else
                sql = sql & " CLI_NRODOC=" & XN(txtBuscaCliente)
            End If
             'sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
'        Else
'            sql = sql & " CLI_CODIGO LIKE '" & Trim(txtcodigo) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtBuscarCliDescri.Text = rec!CLI_RAZSOC
            txtCodigo.Text = rec!CLI_CODIGO
            
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
           ' txtBuscaCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtBuscaOS_GotFocus()
    SelecTexto txtBuscaOS
End Sub

Private Sub txtBuscaOS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarOS "txtBuscaOS", "CODIGO"
    End If
End Sub

Private Sub txtBuscaOS_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBuscaOS_LostFocus()
    If txtBuscaOS.Text <> "" Then
        cSQL = "SELECT OS_NUMERO, OS_NOMBRE FROM OBRA_SOCIAL WHERE OS_NUMERO = " & XN(txtBuscaOS.Text)
        rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtBuscaOS.Text = ChkNull(rec!OS_NUMERO)
            txtBuscarOSNombre.Text = ChkNull(rec!OS_NOMBRE)
        Else
            MsgBox "Obra Social inexistente", vbExclamation, TIT_MSGBOX
            'txtBuscaOS.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtBuscarCliDescri_GotFocus()
    SelecTexto txtBuscarCliDescri
End Sub

Private Sub txtBuscarCliDescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
        ActivoGrid = 0
    End If
End Sub

Private Sub txtBuscarCliDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtBuscarOSNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarOS "txtBuscaOS", "CODIGO"
    End If
End Sub

Private Sub txtBuscarOSNombre_LostFocus()
    If txtBuscaOS.Text = "" And txtBuscarOSNombre.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT OS_NUMERO,OS_NOMBRE FROM OBRA_SOCIAL WHERE OS_NOMBRE LIKE '" & Trim(txtBuscarOSNombre.Text) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarOS "txtBuscaOS", "CADENA", Trim(txtBuscarOSNombre)
                If rec.State = 1 Then rec.Close
                txtBuscarOSNombre.SetFocus
            Else
                txtBuscaOS.Text = rec!OS_NUMERO
                txtBuscarOSNombre.Text = ChkNull(rec!OS_NOMBRE)
            End If
            
        Else
            If MsgBox("La Obra Social no existe!  ¿Desea agregarla al Sistema?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            'preguntar si quiere agregarlo y abrir abm de tratamientos
            'MsgBox "Tratamiento inexistente", vbExclamation, TIT_MSGBOX
                gObraS = 1
                ABMObraSocial.txtDescri.Text = txtBuscarOSNombre.Text
                ABMObraSocial.Show vbModal
                txtBuscarOSNombre.SetFocus
            Else
                txtBuscaOS.SetFocus
            End If
            gObraS = 0
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub
Public Sub BuscarOS(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
        
    Set B = New CBusqueda
    With B
        cSQL = "SELECT OS_NOMBRE, OS_NUMERO"
        cSQL = cSQL & " FROM OBRA_SOCIAL "
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE OS_NOMBRE LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "OS_NOMBRE"
        campo1 = .Field
        .Field = "OS_NUMERO"
        campo2 = .Field
        
        .OrderBy = "OS_NOMBRE"
        camponumerico = False
        .Titulo = "Busqueda de Obras Sociales :"
        .MaxRecords = 1
        .Show
    
        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtBuscaOS" Then
                txtBuscaOS.Text = .ResultFields(2)
                txtBuscaOS_LostFocus
            Else
                'txtBuscaCliente.Text = .ResultFields(2)
                'txtBuscaCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
    
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
        
        hSQL = "Nombre, Código, DNI"
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
                txtCodigo.Text = .ResultFields(2)
                'txtCodCli_LostFocus
            Else
                If .ResultFields(3) = "" Then
                    txtBuscaCliente.Text = .ResultFields(2)
                    txtCodigo.Text = .ResultFields(2)
                Else
                    txtBuscaCliente.Text = .ResultFields(3)
                End If
                txtBuscaCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub
