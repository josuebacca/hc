VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmListadoProvedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores...."
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   30
      TabIndex        =   18
      Top             =   2220
      Width           =   7815
      Begin VB.CommandButton CBImpresora 
         Caption         =   "&Configurar Impresora"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   660
         Width           =   1710
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         Height          =   225
         Left            =   4020
         TabIndex        =   10
         Top             =   330
         Width           =   780
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "Pantalla"
         Height          =   195
         Left            =   945
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optImpresora 
         Caption         =   "Impresora"
         Height          =   195
         Left            =   2370
         TabIndex        =   9
         Top             =   330
         Width           =   1050
      End
      Begin VB.Label lblImpresora 
         AutoSize        =   -1  'True
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1995
         TabIndex        =   20
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   330
         Width           =   600
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Sentido..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   5505
      TabIndex        =   24
      Top             =   1200
      Width           =   2340
      Begin VB.OptionButton optDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Descendente"
         Height          =   210
         Left            =   570
         TabIndex        =   7
         Top             =   645
         Width           =   1275
      End
      Begin VB.OptionButton optAsc 
         Alignment       =   1  'Right Justify
         Caption         =   "Ascendente"
         Height          =   255
         Left            =   675
         TabIndex        =   6
         Top             =   345
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordenado por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   3165
      TabIndex        =   23
      Top             =   1200
      Width           =   2340
      Begin VB.OptionButton optCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Código"
         Height          =   255
         Left            =   945
         TabIndex        =   4
         Top             =   345
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optRazSoc 
         Alignment       =   1  'Right Justify
         Caption         =   "Razón Social"
         Height          =   210
         Left            =   495
         TabIndex        =   5
         Top             =   645
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmListadoProvedores.frx":0000
      Height          =   705
      Left            =   6195
      Picture         =   "frmListadoProvedores.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3390
      Width           =   810
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   4110
      Top             =   3405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ver..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   17
      Top             =   1200
      Width           =   3135
      Begin VB.OptionButton optDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "Listado Detallado"
         Height          =   255
         Left            =   630
         TabIndex        =   2
         Top             =   615
         Width           =   1770
      End
      Begin VB.OptionButton optGeneral 
         Alignment       =   1  'Right Justify
         Caption         =   "Listado General"
         Height          =   345
         Left            =   780
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "&Listar"
      Height          =   705
      Left            =   5370
      Picture         =   "frmListadoProvedores.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3390
      Width           =   810
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   705
      Left            =   7020
      Picture         =   "frmListadoProvedores.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3390
      Width           =   810
   End
   Begin VB.Frame Frame2 
      Caption         =   "   Proveedor ......"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   30
      TabIndex        =   13
      Top             =   -15
      Width           =   7815
      Begin VB.ComboBox cboBuscaTipoProv 
         Height          =   315
         Left            =   1815
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   4095
      End
      Begin VB.TextBox txtBuscaProv 
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   1
         Top             =   705
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   28
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   14
         Top             =   720
         Width           =   810
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   3660
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   16
      Top             =   3615
      Width           =   660
   End
End
Attribute VB_Name = "frmListadoProvedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wSentido As String
Dim wCondicion As String

Private Sub CBImpresora_Click()
  CDImpresora.PrinterDefault = True
  CDImpresora.ShowPrinter
  lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cmdListar_Click()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized 'crptMinimized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.SortFields(0) = ""

    If optGeneral.Value = True Then 'LISTADO GENERAL DE DE UN TIPO DE PROVEEDOR SELECCIONADO
        If cboBuscaTipoProv.List(cboBuscaTipoProv.ListIndex) <> "(Todos)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = "{PROVEEDOR.TPR_CODIGO}=" & cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex)
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {PROVEEDOR.TPR_CODIGO}=" & cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex)
            End If
        End If
        If txtBuscaProv.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {PROVEEDOR.PROV_RAZSOC} LIKE '" & Trim(txtBuscaProv.Text) & "%'"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {PROVEEDOR.PROV_RAZSOC} LIKE '" & Trim(txtBuscaProv.Text) & "%'"
            End If
        End If
            
        If optAsc.Value = True Then
           wSentido = "+"
        Else
           wSentido = "-"
        End If
        If optCodigo.Value = True Then
            wCondicion = wSentido & "{PROVEEDOR.PROV_CODIGO}"
        ElseIf optRazSoc.Value = True Then
            wCondicion = wSentido & "{PROVEEDOR.PROV_RAZSOC}"
        End If
        Rep.SortFields(0) = wCondicion
        
        Rep.WindowTitle = "Maestro de Proveedores..."
        Rep.ReportFileName = DRIVE & DirReport & "MaestroProveedores.rpt"
    End If
    
    If optDetallado.Value = True Then 'LISTADO DETALLADO
        
        If cboBuscaTipoProv.List(cboBuscaTipoProv.ListIndex) <> "(Todos)" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = "{PROVEEDOR.TPR_CODIGO}=" & cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex)
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {PROVEEDOR.TPR_CODIGO}=" & cboBuscaTipoProv.ItemData(cboBuscaTipoProv.ListIndex)
            End If
        End If
        If txtBuscaProv.Text <> "" Then
            If Rep.SelectionFormula = "" Then
                Rep.SelectionFormula = " {PROVEEDOR.PROV_RAZSOC} LIKE '" & Trim(txtBuscaProv.Text) & "%'"
            Else
                Rep.SelectionFormula = Rep.SelectionFormula & " AND {PROVEEDOR.PROV_RAZSOC} LIKE '" & Trim(txtBuscaProv.Text) & "%'"
            End If
        End If
        Rep.WindowTitle = "Maestro de Proveedores - Detallado..."
        Rep.ReportFileName = DRIVE & DirReport & "maestroproveedoresDetalle.rpt"
    End If
    
    If optPantalla.Value = True Then
         Rep.Destination = crptToWindow
    ElseIf optImpresora.Value = True Then
        Rep.Destination = crptToPrinter
        Rep.PrintFileType = crptExcel50
    ElseIf optExcel.Value = True Then
        Rep.Destination = crptToFile
        Rep.PrintFileType = crptExcel50
    End If
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    txtBuscaProv.Text = ""
    cboBuscaTipoProv.ListIndex = 0
    optCodigo.Value = True
    optAsc.Value = True
    optDetallado.Value = True
    cboBuscaTipoProv.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmListadoProvedores = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    lblEstado.Caption = ""

    txtBuscaProv.Text = ""
       
    CargoComboTipoProveedor
    'impresora actual
    lblImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Public Sub CargoComboTipoProveedor()
    'Cargo el combo Tipo de Proveedor
    cboBuscaTipoProv.Clear
    
    sql = "SELECT * FROM TIPO_PROVEEDOR ORDER BY TPR_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
            cboBuscaTipoProv.AddItem "(Todos)"
        Do While Not rec.EOF
            cboBuscaTipoProv.AddItem rec.Fields!TPR_DESCRI
            cboBuscaTipoProv.ItemData(cboBuscaTipoProv.NewIndex) = rec.Fields!TPR_CODIGO
            rec.MoveNext
        Loop
        cboBuscaTipoProv.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtBuscaProv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
