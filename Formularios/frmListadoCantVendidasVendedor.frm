VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmListadoCantVendidasVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cantidades Vendidas por Vendedor"
   ClientHeight    =   2250
   ClientLeft      =   1515
   ClientTop       =   1740
   ClientWidth     =   5145
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
   ScaleHeight     =   2250
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   1845
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3825
      TabIndex        =   4
      Top             =   1845
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
      Height          =   1335
      Left            =   45
      TabIndex        =   22
      Top             =   30
      Width           =   5055
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   405
         Width           =   3195
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   750
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   315
         Left            =   3285
         TabIndex        =   1
         Top             =   750
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   345
         TabIndex        =   26
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   24
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   345
         TabIndex        =   23
         Top             =   780
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
      TabIndex        =   12
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   13
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
         TabIndex        =   21
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
         TabIndex        =   15
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
         TabIndex        =   14
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
      TabIndex        =   7
      Top             =   1440
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmListadoCantVendidasVendedor.frx":0748
         Left            =   450
         List            =   "frmListadoCantVendidasVendedor.frx":0755
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   2340
      Top             =   1485
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
      TabIndex        =   8
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

Private Sub CmdAceptar_Click()
'    If FechaDesde.Text = "" Then
'        MsgBox "Falta Ingresar la Fecha Desde", vbExclamation, TIT_MSGBOX
'        FechaDesde.SetFocus
'        Exit Sub
'    End If
'    If FechaHasta.Text = "" Then
'        MsgBox "Falta Ingresar la Fecha Hasta", vbExclamation, TIT_MSGBOX
'        FechaHasta.SetFocus
'        Exit Sub
'    End If
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    
    'SOLO FACTURAS DEFINITIVAS
    Rep.SelectionFormula = " {FACTURA_CLIENTE.EST_CODIGO}=3"
    
    If cboVendedor.List(cboVendedor.ListIndex) <> "(Todos)" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.VEN_CODIGO}=" & XN(cboVendedor.ItemData(cboVendedor.ListIndex))
        End If
    End If
    If FechaDesde.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}>= DATE (" & Mid(FechaDesde.Text, 7, 4) & "," & Mid(FechaDesde.Text, 4, 2) & "," & Mid(FechaDesde.Text, 1, 2) & ")"
        End If
    End If
    If FechaHasta.Text <> "" Then
        If Rep.SelectionFormula = "" Then
            Rep.SelectionFormula = " {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
        Else
            Rep.SelectionFormula = Rep.SelectionFormula & " AND {FACTURA_CLIENTE.FCL_FECHA}<= DATE (" & Mid(FechaHasta.Text, 7, 4) & "," & Mid(FechaHasta.Text, 4, 2) & "," & Mid(FechaHasta.Text, 1, 2) & ")"
        End If
    End If
    If FechaDesde.Text <> "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text <> "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: " & FechaDesde.Text & "   Hasta: " & Date & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text <> "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & FechaHasta.Text & "'"
    ElseIf FechaDesde.Text = "" And FechaHasta.Text = "" Then
        Rep.Formulas(0) = "FECHA='" & "Desde: Inicio" & "   Hasta: " & Date & "'"
    End If
    
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    
    Rep.WindowTitle = "Listado de Cantidades Vendidas"
    Rep.ReportFileName = DRIVE & DirReport & "cantidades_vendidas_vendedor.rpt"
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
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
    Me.Top = 0
    Me.Left = 0
    cboVendedor.AddItem "(Todos)"
    CargoComboBox cboVendedor, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 0
    'CARGO COMBO LINEA
    cboDestino.ListIndex = 0
    'FechaDesde.Text = Date
    'FechaHasta.Text = Date
End Sub
