VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmFacturaProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas de Proveedores..."
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   705
   ClientWidth     =   10050
   ControlBox      =   0   'False
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10050
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   7515
      TabIndex        =   13
      Top             =   6105
      Width           =   1230
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   6270
      TabIndex        =   12
      Top             =   6105
      Width           =   1230
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   8760
      TabIndex        =   14
      Top             =   6105
      Width           =   1230
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Eliminar"
      Height          =   405
      Left            =   4440
      TabIndex        =   17
      Top             =   6105
      Visible         =   0   'False
      Width           =   1230
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   6015
      Left            =   45
      TabIndex        =   24
      Top             =   60
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Datos"
      TabPicture(0)   =   "frmFacturaProveedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProveedor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmFacturaProveedores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Facturas..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4905
         Left            =   135
         TabIndex        =   25
         Top             =   1050
         Width           =   9705
         Begin VB.TextBox txtCodInt 
            Height          =   345
            Left            =   7995
            TabIndex        =   47
            Top             =   810
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.TextBox txtPrecio 
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
            Height          =   315
            Left            =   6600
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1185
            Width           =   960
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
            Height          =   330
            Left            =   8535
            TabIndex        =   15
            ToolTipText     =   "Quitar Producto"
            Top             =   1185
            Width           =   900
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
            Height          =   315
            Left            =   150
            TabIndex        =   7
            Top             =   1185
            Width           =   1170
         End
         Begin VB.CommandButton cmdAsignar 
            Caption         =   "A&gregar"
            Height          =   330
            Left            =   7605
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Asignar Producto"
            Top             =   1185
            Width           =   900
         End
         Begin VB.TextBox txtCantidad 
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
            Height          =   315
            Left            =   5685
            MaxLength       =   10
            TabIndex        =   9
            Top             =   1185
            Width           =   885
         End
         Begin VB.TextBox txtdescri 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1335
            TabIndex        =   8
            Top             =   1185
            Width           =   4320
         End
         Begin VB.ComboBox CboGastos 
            Height          =   315
            Left            =   1230
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   3825
         End
         Begin VB.ComboBox cboComprobante 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   585
            Width           =   1665
         End
         Begin VB.CommandButton cmdNuevoGasto 
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
            Left            =   5070
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedores.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Agregar País"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   7470
            MaxLength       =   40
            TabIndex        =   16
            Top             =   4440
            Width           =   1710
         End
         Begin VB.TextBox txtNroSucursal 
            Height          =   315
            Left            =   3990
            MaxLength       =   4
            TabIndex        =   4
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtNroComprobante 
            Height          =   315
            Left            =   4515
            MaxLength       =   8
            TabIndex        =   5
            Top             =   600
            Width           =   960
         End
         Begin FechaCtl.Fecha FechaComprobante 
            Height          =   315
            Left            =   6525
            TabIndex        =   6
            Top             =   615
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin MSFlexGridLib.MSFlexGrid GrdDetalleFactura 
            Height          =   2850
            Left            =   120
            TabIndex        =   42
            Top             =   1530
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   5027
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   290
            BackColorSel    =   16761024
            FocusRect       =   0
            HighLight       =   2
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6660
            TabIndex        =   46
            ToolTipText     =   "Agregar Producto"
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1395
            TabIndex        =   45
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5730
            TabIndex        =   44
            ToolTipText     =   "Agregar Producto"
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   210
            TabIndex        =   43
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6765
            TabIndex        =   41
            Top             =   4485
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Gasto:"
            Height          =   195
            Left            =   690
            TabIndex        =   39
            Top             =   255
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Comprobante:"
            Height          =   195
            Left            =   165
            TabIndex        =   38
            Top             =   630
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   3315
            TabIndex        =   27
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   6000
            TabIndex        =   26
            Top             =   645
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1575
         Left            =   -74835
         TabIndex        =   30
         Top             =   375
         Width           =   9645
         Begin VB.TextBox txtDesProv 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
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
            Left            =   2985
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Descripción"
            Top             =   540
            Width           =   4440
         End
         Begin VB.TextBox txtProveedor 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1545
            MaxLength       =   40
            TabIndex        =   18
            Top             =   540
            Width           =   975
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "&Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5490
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   21
            ToolTipText     =   "Buscar "
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   1980
         End
         Begin VB.CommandButton cmdBuscarProveedor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2550
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedores.frx":03C2
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Buscar Proveedor"
            Top             =   540
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   4095
            TabIndex        =   20
            Top             =   1005
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   1545
            TabIndex        =   19
            Top             =   990
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   35
            Top             =   570
            Width           =   810
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   480
            TabIndex        =   34
            Top             =   1020
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   3030
            TabIndex        =   33
            Top             =   1050
            Width           =   960
         End
      End
      Begin VB.Frame FrameProveedor 
         Caption         =   "Proveedor..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   135
         TabIndex        =   28
         Top             =   345
         Width           =   9705
         Begin VB.TextBox txtCodTipoProv 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   37
            Top             =   285
            Width           =   345
         End
         Begin VB.CommandButton cmdBuscarProveedor1 
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
            Left            =   2280
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaProveedores.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Buscar Proveedor"
            Top             =   285
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtProvRazSoc 
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
            Left            =   2715
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "Descripción"
            Top             =   285
            Width           =   5340
         End
         Begin VB.TextBox txtCodProveedor 
            Height          =   315
            Left            =   1620
            MaxLength       =   40
            TabIndex        =   0
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   645
            TabIndex        =   29
            Top             =   300
            Width           =   555
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3870
         Left            =   -74865
         TabIndex        =   22
         Top             =   1995
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   6826
         _Version        =   393216
         Cols            =   16
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorSel    =   16761024
         AllowBigSelection=   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
      TabIndex        =   23
      Top             =   6180
      Width           =   660
   End
End
Attribute VB_Name = "frmFacturaProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub CalculoTotal()
    txtTotal.Text = "0,00"
    For i = 1 To GrdDetalleFactura.Rows - 1
       If Trim(GrdDetalleFactura.TextMatrix(i, 0)) <> "" Then
          If GrdDetalleFactura.TextMatrix(i, 4) <> "" Then
            txtTotal.Text = Valido_Importe(CDbl(txtTotal.Text) + CDbl(GrdDetalleFactura.TextMatrix(i, 4)))
          End If
       End If
    Next i
End Sub

Private Sub LlenarComboGastos()
    SQL = "SELECT TGT_CODIGO, TGT_DESCRI " & _
          "  FROM TIPO_GASTO " & _
          " ORDER BY TGT_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
       Do While rec.EOF = False
          CboGastos.AddItem rec!TGT_DESCRI
          CboGastos.ItemData(CboGastos.NewIndex) = rec!TGT_CODIGO
          rec.MoveNext
       Loop
       CboGastos.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboComprobante()
    SQL = "SELECT TCO_CODIGO,TCO_DESCRI"
    SQL = SQL & " FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_DESCRI LIKE 'FACTU%'"
    SQL = SQL & " ORDER BY TCO_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboComprobante.AddItem rec!TCO_DESCRI
            cboComprobante.ItemData(cboComprobante.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboComprobante.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub cmdAsignar_Click()
    If txtCantidad = "" Then
        MsgBox "Debe Ingresar la cantidad", vbExclamation, TIT_MSGBOX
        txtCantidad.SetFocus
        Exit Sub
    End If
    If txtPrecio.Text = "" Then
        MsgBox "Debe Ingresar el Precio", vbExclamation, TIT_MSGBOX
        txtPrecio.SetFocus
        Exit Sub
    End If
    If txtcodigo.Text <> "" Then
        For i = 1 To GrdDetalleFactura.Rows - 1
            If GrdDetalleFactura.TextMatrix(i, 5) = CLng(txtCodInt.Text) Then
                'MsgBox "El producto ya fue ingresado", vbExclamation, TIT_MSGBOX
                'TxtCODIGO.SetFocus
                'Exit Sub
            End If
        Next
        GrdDetalleFactura.AddItem Trim(txtcodigo.Text) & Chr(9) & Trim(txtDescri.Text) & Chr(9) & _
                                  Trim(txtCantidad.Text) & Chr(9) & txtPrecio.Text & Chr(9) & _
                                  Valido_Importe(CInt(txtCantidad.Text) * CDbl(txtPrecio.Text)) & Chr(9) & Trim(txtCodInt.Text)
         
        CalculoTotal
        txtcodigo.Text = ""
        txtcodigo.SetFocus
     Else
        MsgBox "Debe seleccionar un Producto"
    End If
End Sub

Private Sub cmdNuevoGasto_Click()
    ABMTipoGatos.vMode = 1
    mOrigen = False
    ABMTipoGatos.Show vbModal
    CboGastos.Clear
    'CARGO COMBO GASTOS
    LlenarComboGastos
    BuscaCodigoProxItemData CboGastos.ListCount, CboGastos
    CboGastos.SetFocus
End Sub

Private Sub cmdBorrar_Click()
    
    If MsgBox("¿Seguro que desea eliminar la Factura?", vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
        On Error GoTo Seclavose
         lblEstado.Caption = "Eliminando..."
         Screen.MousePointer = vbHourglass
         DBConn.BeginTrans
         
         'DETALLE FACTURA
    
'         BORRO DE LA CUENTA CORRIENTE DEL PROVEEDOR
'         Descuento los Litros del Stock
                                          
         DBConn.CommitTrans
         lblEstado.Caption = ""
         Screen.MousePointer = vbNormal
         cmdNuevo_Click
    End If
    Exit Sub
    
Seclavose:
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    Set Rec1 = New ADODB.Recordset
    GrdModulos.Rows = 1
    SQL = "SELECT F.TPR_CODIGO,F.PROV_CODIGO, P.PROV_RAZSOC, F.TCO_CODIGO, C.TCO_DESCRI,"
    SQL = SQL & " F.FPR_NROSUC,F.FPR_NUMERO,F.FPR_FECHA, F.FPR_TOTAL"
    SQL = SQL & " FROM PROVEEDOR P, FACTURA_PROVEEDOR F, TIPO_COMPROBANTE C"
    SQL = SQL & " WHERE P.TPR_CODIGO = F.TPR_CODIGO"
    SQL = SQL & " AND P.PROV_CODIGO = F.PROV_CODIGO"
    SQL = SQL & " AND F.TCO_CODIGO = C.TCO_CODIGO"
    If txtProveedor.Text <> "" Then SQL = SQL & " AND P.PROV_CODIGO=" & XN(txtProveedor)
    If FechaDesde <> "" Then SQL = SQL & " AND F.FPR_FECHA >=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then SQL = SQL & " AND F.FPR_FECHA <=" & XDQ(FechaHasta)
    SQL = SQL & " ORDER BY F.FPR_FECHA DESC"
    
    Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            GrdModulos.AddItem Rec1!FPR_FECHA & Chr(9) & _
                               Rec1!TCO_DESCRI & Chr(9) & _
                               Format(Rec1!FPR_NROSUC, "0000") & "-" & Format(Rec1!FPR_NUMERO, "00000000") & Chr(9) & _
                               Rec1!PROV_RAZSOC & Chr(9) & _
                               Valido_Importe(Rec1!FPR_TOTAL) & Chr(9) & _
                               Rec1!TCO_CODIGO & Chr(9) & _
                               Rec1!PROV_CODIGO & Chr(9) & _
                               Rec1!TPR_CODIGO & Chr(9)

            Rec1.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightWithFocus
        GrdModulos.Col = 0
        GrdModulos.SetFocus
    Else
        MsgBox "No se encontraron Facturas. Verifique!", vbExclamation, TIT_MSGBOX
        txtProveedor.SetFocus
    End If
    Rec1.Close
End Sub

Private Sub cmdBuscarProveedor_Click()
    BusquedaDeProveedor txtProveedor, "CODIGO"
    txtProveedor_LostFocus
    cmdBuscarProveedor.SetFocus
End Sub

Private Sub cmdBuscarProveedor1_Click()
    BusquedaDeProveedor txtCodProveedor, "CODIGO"
    txtCodProveedor_LostFocus
    txtCodProveedor.SetFocus
End Sub

Private Sub cmdGrabar_Click()
    Set rec = New ADODB.Recordset
    
    If ValidarGastosProveedor = False Then Exit Sub
        
    If MsgBox("¿Confirma Factura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorCarga
    
    DBConn.BeginTrans
    SQL = "SELECT FPR_FECHA "
    SQL = SQL & " FROM FACTURA_PROVEEDOR "
    SQL = SQL & " WHERE TPR_CODIGO  = " & XN(txtCodTipoProv.Text)
    SQL = SQL & "   AND PROV_CODIGO = " & XN(txtCodProveedor.Text)
    SQL = SQL & "   AND TCO_CODIGO  = " & XN(cboComprobante.ItemData(cboComprobante.ListIndex))
    SQL = SQL & "   AND FPR_NROSUC  = " & XN(txtNroSucursal.Text)
    SQL = SQL & "   AND FPR_NUMERO  = " & XN(txtNroComprobante.Text)
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        SQL = "INSERT INTO FACTURA_PROVEEDOR(TPR_CODIGO,PROV_CODIGO,"
        SQL = SQL & " TCO_CODIGO,FPR_NROSUC,FPR_NUMERO,"
        SQL = SQL & " FPR_FECHA,FPR_TOTAL,FPR_SALDO,TGT_CODIGO)"
        SQL = SQL & " VALUES ("
        SQL = SQL & XN(txtCodTipoProv.Text) & ", "
        SQL = SQL & XN(txtCodProveedor.Text) & ", "
        SQL = SQL & XN(cboComprobante.ItemData(cboComprobante.ListIndex)) & ", "
        SQL = SQL & XN(txtNroSucursal.Text) & ", "
        SQL = SQL & XN(txtNroComprobante.Text) & ", "
        SQL = SQL & XDQ(FechaComprobante.Text) & ", "
        SQL = SQL & XN(txtTotal.Text) & ", " 'TOTAL FACTURA
        SQL = SQL & XN("0") & ", " 'SALDO
        SQL = SQL & XN(CboGastos.ItemData(CboGastos.ListIndex)) & ")"
        DBConn.Execute SQL
        
        
        'DETALLE FACTURA
        For i = 1 To GrdDetalleFactura.Rows - 1
            If GrdDetalleFactura.TextMatrix(i, 0) <> "" And GrdDetalleFactura.TextMatrix(i, 1) <> "" Then
                SQL = "INSERT INTO DETALLE_FACTURA_PROVEEDOR(TPR_CODIGO,PROV_CODIGO,TCO_CODIGO,"
                SQL = SQL & " FPR_NROSUC,FPR_NUMERO,FPR_NROITEM,PTO_CODIGO,FPR_DESCRI,"
                SQL = SQL & " FPR_CANTIDAD,FPR_PRECIO,FPR_SUBTOT)"
                SQL = SQL & " VALUES ("
                SQL = SQL & XN(txtCodTipoProv.Text) & ", "
                SQL = SQL & XN(txtCodProveedor.Text) & ", "
                SQL = SQL & XN(cboComprobante.ItemData(cboComprobante.ListIndex)) & ", "
                SQL = SQL & XN(txtNroSucursal.Text) & ", "
                SQL = SQL & XN(txtNroComprobante.Text) & ", "
                SQL = SQL & i & ", "
                SQL = SQL & XN(GrdDetalleFactura.TextMatrix(i, 5)) & ", "
                SQL = SQL & XS(GrdDetalleFactura.TextMatrix(i, 1)) & ", "
                SQL = SQL & XN(GrdDetalleFactura.TextMatrix(i, 2)) & ", "
                SQL = SQL & XN(GrdDetalleFactura.TextMatrix(i, 3)) & ", "
                SQL = SQL & XN(GrdDetalleFactura.TextMatrix(i, 4)) & ")"
                DBConn.Execute SQL
                
                'CATUALIZO EL STOCK
                SQL = "UPDATE STOCK SET"
                SQL = SQL & " DST_STKFIS = DST_STKFIS + " & XN(GrdDetalleFactura.TextMatrix(i, 2))
                SQL = SQL & " WHERE STK_CODIGO = " & XN(Sucursal)
                SQL = SQL & " AND PTO_CODIGO = " & XN(GrdDetalleFactura.TextMatrix(i, 5))
                DBConn.Execute SQL
            End If
        Next
        
    Else
        MsgBox "La Factura ya fue ingresada", vbCritical, TIT_MSGBOX
    End If
    rec.Close
        
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    cmdNuevo_Click
    Exit Sub
    
HayErrorCarga:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarGastosProveedor() As Boolean
    If txtCodProveedor.Text = "" Then
        MsgBox "Debe Ingresar el Proveedor. Verifique!", vbCritical, TIT_MSGBOX
        txtCodProveedor.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If CboGastos.ListIndex = -1 Then
        MsgBox "Debe Ingresar un Tipo de Gasto. Verifique!", vbCritical, TIT_MSGBOX
        CboGastos.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If cboComprobante.ListIndex = -1 Then
        MsgBox "Debe Ingresar un Tipo de Comprobante. Verifique!", vbCritical, TIT_MSGBOX
        cboComprobante.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtNroSucursal.Text = "" Then
        MsgBox "Debe Ingresar la Sucursal. Verifique!", vbCritical, TIT_MSGBOX
        txtNroSucursal.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If txtNroComprobante.Text = "" Then
        MsgBox "Debe Ingresar el Nro de Comprobante. Verifique!", vbCritical, TIT_MSGBOX
        txtNroComprobante.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If FechaComprobante.Text = "" Then
        MsgBox "La Fecha del comprobate es requerida", vbExclamation, TIT_MSGBOX
        FechaComprobante.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    If GrdDetalleFactura.Rows = 1 Then
        MsgBox "Debe Ingresar un Item de Producto. Verifique!", vbCritical, TIT_MSGBOX
        txtcodigo.SetFocus
        ValidarGastosProveedor = False
        Exit Function
    End If
    ValidarGastosProveedor = True
End Function

Private Sub cmdNuevo_Click()
    LimpiarBusqueda
    txtCodProveedor.Text = ""
    txtCodTipoProv.Text = ""
    FechaComprobante.Text = ""
    FechaComprobante.Enabled = True
    BuscaProx "MERCADERIAS", CboGastos
    txtNroSucursal.Text = ""
    txtNroComprobante.Text = ""
    
    txtTotal.Text = "0,00"
    GrdDetalleFactura.Rows = 1
    txtcodigo.Text = ""
    CmdBorrar.Enabled = False
    cmdGrabar.Enabled = True
    tabDatos.Tab = 0
    txtCodProveedor.SetFocus
End Sub

Private Sub cmdQuitar_Click()
    If GrdDetalleFactura.Rows <> 1 Then
        If MsgBox("¿Seguro desea Eliminar el Producto: " & Trim(GrdDetalleFactura.TextMatrix(GrdDetalleFactura.RowSel, 1)) & "? ", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            lblEstado.Caption = "Borrando..."
            Screen.MousePointer = vbHourglass
            If GrdDetalleFactura.Rows = 2 Then
                GrdDetalleFactura.HighLight = flexHighlightNever
                GrdDetalleFactura.Rows = 1
                CalculoTotal
                txtcodigo.SetFocus
            Else
                GrdDetalleFactura.RemoveItem (GrdDetalleFactura.RowSel)
                CalculoTotal
                txtcodigo.SetFocus
            End If
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmFacturaProveedores = Nothing
        Unload Me
    End If
End Sub

Private Sub FechaComprobante_LostFocus()
    If Trim(FechaComprobante) = "" Then FechaComprobante.Text = Date
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And ActiveControl.Name <> "txtcodigo" And _
       ActiveControl.Name <> "txtdescri" And ActiveControl.Name <> "txtCodProveedor" And _
       ActiveControl.Name <> "txtProvRazSoc" And ActiveControl.Name <> "txtProveedor" Then
        tabDatos.Tab = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MySendKeys Chr(9)
    End If
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub LimpiarBusqueda()
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    txtProveedor.Text = ""
    GrdModulos.Rows = 1
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    'Call Centrar_pantalla(Me)
    Me.Top = 0
    Me.Left = 0
    
    'CONFIGURO GRILLA BUSQUEDA
    GrdModulos.FormatString = "^Fecha|Comprobante|^Nro Comp.|Proveeor|>Importe|Tipo Comprobante|Nro Proveedor|Tipo Proveedor"
                            
    GrdModulos.ColWidth(0) = 1300   'Fecha
    GrdModulos.ColWidth(1) = 1800   'Comprobante
    GrdModulos.ColWidth(2) = 1500   'Numero comprobante
    GrdModulos.ColWidth(3) = 3500   'Proveedor
    GrdModulos.ColWidth(4) = 1200   'Total
    GrdModulos.ColWidth(5) = 0      'Tipo Comprobante
    GrdModulos.ColWidth(6) = 0      'Nro Proveedor
    GrdModulos.ColWidth(7) = 0      'Tipo Proveedor
    GrdModulos.Cols = 8
    GrdModulos.Rows = 1
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    
    'Configuro las Grillas
    'GRILLA DONDE SE CRAGAN LOS PRODUCTOS
    GrdDetalleFactura.FormatString = "^Código|<Producto|^Cantidad|>Precio|>Total|CodInt"
    GrdDetalleFactura.ColWidth(0) = 1200 'CODIGO PRODUCTO
    GrdDetalleFactura.ColWidth(1) = 4500 'PRODUCTO
    GrdDetalleFactura.ColWidth(2) = 1000 'CANTIDAD
    GrdDetalleFactura.ColWidth(3) = 1100 'PRECIO
    GrdDetalleFactura.ColWidth(4) = 1200 'TOTAL
    GrdDetalleFactura.ColWidth(5) = 0    'CODINT
    GrdDetalleFactura.Rows = 1
    GrdDetalleFactura.HighLight = flexHighlightWithFocus
    GrdDetalleFactura.BorderStyle = flexBorderNone
    GrdDetalleFactura.row = 0
    For i = 0 To GrdDetalleFactura.Cols - 1
        GrdDetalleFactura.Col = i
        GrdDetalleFactura.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdDetalleFactura.CellBackColor = &H808080    'GRIS OSCURO
        GrdDetalleFactura.CellFontBold = True
    Next
        
    tabDatos.Tab = 0
    
    cmdGrabar.Enabled = True
    CmdBorrar.Enabled = False
    lblEstado.Caption = ""
    txtTotal.Text = "0,00"
    
    'CARGO COMBO COMPROBANTES
    LlenarComboComprobante
    
    'CARGO COMBO GASTOS
    LlenarComboGastos
    BuscaProx "MERCADERIAS", CboGastos
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.Rows > 1 Then
        tabDatos.Tab = 0
        
        txtTotal.Text = "0,00"
        GrdDetalleFactura.Rows = 1
        
        BuscarfacturaProveedor (GrdModulos.RowSel)
        CmdBorrar.Enabled = True
        cmdGrabar.Enabled = False
        tabDatos.Tab = 0
    End If
End Sub

Private Sub BuscarfacturaProveedor(Fila As Integer)

    Dim rec As New ADODB.Recordset
    
    SQL = "SELECT * FROM FACTURA_PROVEEDOR"
    SQL = SQL & " WHERE"
    SQL = SQL & " TPR_CODIGO=" & XN(GrdModulos.TextMatrix(Fila, 7))
    SQL = SQL & " AND PROV_CODIGO=" & XN(GrdModulos.TextMatrix(Fila, 6))
    SQL = SQL & " AND TCO_CODIGO=" & XN(GrdModulos.TextMatrix(Fila, 5))
    SQL = SQL & " AND FPR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(Fila, 2), 4))
    SQL = SQL & " AND FPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(Fila, 2), 8))
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtCodProveedor.Text = rec!PROV_CODIGO
        txtCodProveedor_LostFocus
        Call BuscaCodigoProxItemData(rec!TGT_CODIGO, CboGastos)
        Call BuscaCodigoProxItemData(rec!TCO_CODIGO, cboComprobante)
        txtNroSucursal.Text = Format(rec!FPR_NROSUC, "0000")
        txtNroComprobante.Text = Format(rec!FPR_NUMERO, "00000000")
        FechaComprobante.Text = rec!FPR_FECHA
        txtTotal.Text = Valido_Importe(Chk0(rec!FPR_TOTAL))
    End If
    rec.Close
    
    'DETALLE FACTURA
    SQL = "SELECT D.*, P.PTO_CODBARRAS"
    SQL = SQL & " FROM DETALLE_FACTURA_PROVEEDOR D, PRODUCTO P"
    SQL = SQL & " WHERE"
    SQL = SQL & " D.TPR_CODIGO=" & XN(GrdModulos.TextMatrix(Fila, 7))
    SQL = SQL & " AND D.PROV_CODIGO=" & XN(GrdModulos.TextMatrix(Fila, 6))
    SQL = SQL & " AND D.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(Fila, 5))
    SQL = SQL & " AND D.FPR_NROSUC=" & XN(Left(GrdModulos.TextMatrix(Fila, 2), 4))
    SQL = SQL & " AND D.FPR_NUMERO=" & XN(Right(GrdModulos.TextMatrix(Fila, 2), 8))
    SQL = SQL & " AND P.PTO_CODIGO=D.PTO_CODIGO"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        i = 1
        Do While rec.EOF = False
            GrdDetalleFactura.AddItem IIf(IsNull(rec!PTO_CODBARRAS), rec!PTO_CODIGO, rec!PTO_CODBARRAS) & Chr(9) & Trim(rec!FPR_DESCRI) & Chr(9) & _
                                             Trim(rec!FPR_CANTIDAD) & Chr(9) & Valido_Importe(rec!FPR_PRECIO) & Chr(9) & _
                                             Valido_Importe(rec!FPR_SUBTOT) & Chr(9) & Trim(rec!PTO_CODIGO)
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GrdModulos_dblClick
    End If
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 1 Then
       LimpiarBusqueda
       If Me.Visible = True Then txtProveedor.SetFocus
    End If
End Sub

Private Sub txtCantidad_GotFocus()
    SelecTexto txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
    If txtcodigo.Text = "" Then
        txtcodigo.Text = ""
        txtDescri.Text = ""
        txtCantidad.Text = ""
        txtPrecio.Text = ""
        txtCodInt.Text = ""
        cmdAsignar.Enabled = False
    Else
        cmdAsignar.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto txtcodigo
End Sub

Private Sub txtcodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        txtcodigo.SetFocus
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    If txtcodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = " SELECT P.PTO_DESCRI, P.PTO_CODIGO"
        SQL = SQL & " FROM PRODUCTO P"
        SQL = SQL & " WHERE"
        If IsNumeric(txtcodigo.Text) Then
            SQL = SQL & " P.PTO_CODIGO =" & XN(txtcodigo.Text) & " OR P.PTO_CODBARRAS=" & XS(txtcodigo.Text)
        Else
            SQL = SQL & " P.PTO_CODBARRAS=" & XS(txtcodigo.Text)
        End If
        SQL = SQL & " ORDER BY P.PTO_CODIGO"
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDescri.Text = Trim(rec!PTO_DESCRI)
            txtCodInt.Text = rec!PTO_CODIGO
        Else
            MsgBox "El Código no existe", vbExclamation, TIT_MSGBOX
            txtcodigo.SetFocus
        End If
        rec.Close
    End If
End Sub


Private Sub txtCodProveedor_Change()
    If txtCodProveedor.Text = "" Then
        txtProvRazSoc.Text = ""
        txtCodTipoProv.Text = ""
    End If
End Sub

Private Sub txtCodProveedor_GotFocus()
    SelecTexto txtCodProveedor
End Sub

Private Sub txtCodProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BusquedaDeProveedor txtCodProveedor, "CODIGO"
        txtCodProveedor_LostFocus
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub txtCodProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodProveedor_LostFocus()
    If txtCodProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoProveedor(txtCodProveedor), DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtProvRazSoc.Text = Rec1!PROV_RAZSOC
            txtCodTipoProv.Text = Rec1!TPR_CODIGO
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtdescri_Change()
    If txtDescri.Text = "" Then
        txtcodigo.Text = ""
    End If
End Sub

Private Sub txtdescri_GotFocus()
    SelecTexto txtDescri
End Sub

Private Sub txtdescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarProducto "CODIGO"
        txtDescri.SetFocus
    End If
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDescri_LostFocus()
   If txtcodigo.Text = "" And txtDescri.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        Screen.MousePointer = vbHourglass
        SQL = "SELECT PTO_CODIGO,PTO_DESCRI,PTO_CODBARRAS"
        SQL = SQL & " FROM PRODUCTO"
        SQL = SQL & " WHERE PTO_DESCRI LIKE '" & txtDescri.Text & "%'"
        Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            If Rec1.RecordCount > 1 Then
                'grdGrilla.SetFocus
                BuscarProducto "CADENA", Trim(txtDescri.Text)
                txtDescri.SetFocus
            Else
                txtcodigo.Text = Trim(Rec1!PTO_CODBARRAS)
                txtDescri.Text = Trim(Rec1!PTO_DESCRI)
                txtCodInt.Text = Trim(Rec1!PTO_CODIGO)
            End If
        Else
                MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                txtDescri.Text = ""
        End If
        Rec1.Close
        Screen.MousePointer = vbNormal
    ElseIf txtcodigo.Text = "" And txtDescri.Text = "" Then
        cmdAsignar.Enabled = False
    End If
End Sub

Private Sub txtNroComprobante_GotFocus()
    SelecTexto txtNroComprobante
End Sub

Private Sub txtNroComprobante_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroComprobante_LostFocus()
    If txtNroComprobante.Text <> "" Then
       txtNroComprobante.Text = Format(txtNroComprobante.Text, "00000000")
       If Trim(txtCodProveedor) <> "" Then
          If Trim(txtNroSucursal) <> "" Then
             'Consulto la Factura
             Set rec = New ADODB.Recordset
             cSQL = "SELECT FPR_FECHA,FPR_TOTAL, " & _
                    "       FPR_SALDO, TGT_CODIGO  " & _
                    "  FROM FACTURA_PROVEEDOR " & _
                    " WHERE TPR_CODIGO = " & XN(txtCodTipoProv.Text) & _
                    "   AND PROV_CODIGO = " & XN(txtCodProveedor.Text) & _
                    "   AND TCO_CODIGO = " & XN(cboComprobante.ItemData(cboComprobante.ListIndex)) & _
                    "   AND FPR_NROSUC = " & XN(txtNroSucursal.Text) & _
                    "   AND FPR_NUMERO = " & XN(txtNroComprobante.Text)
             rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
             If rec.EOF = False Then
                'Consulto los Datos completos de la Factura
                FechaComprobante.Text = Trim(rec!FPR_FECHA)
                FechaComprobante.Enabled = False
                txtTotal.Text = Format(rec!FPR_TOTAL, "0.00")
                BuscaProx Trim(rec!TGT_CODIGO), CboGastos
             End If
             rec.Close
             
             'Consulto detalle Factura
             cSQL = "SELECT PTO_CODIGO, FPR_DESCRI, FPR_CANTIDAD, FPR_PRECIO, FPR_SUBTOT " & _
                    "  FROM DETALLE_FACTURA_PROVEEDOR " & _
                    " WHERE TPR_CODIGO = " & XN(txtCodTipoProv) & _
                    "   AND PROV_CODIGO = " & XN(txtCodProveedor) & _
                    "   AND TCO_CODIGO = " & XN(cboComprobante.ItemData(cboComprobante.ListIndex)) & _
                    "   AND FPR_NROSUC = " & XN(txtNroSucursal) & _
                    "   AND FPR_NUMERO = " & XN(txtNroComprobante) & _
                    " ORDER BY FPR_NROITEM"
             rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
             If rec.EOF = False Then
                Do While rec.EOF = False
                   GrdDetalleFactura.AddItem Trim(rec!PTO_CODIGO) & Chr(9) & Trim(rec!FPR_DESCRI) & Chr(9) & _
                                             Trim(rec!FPR_CANTIDAD) & Chr(9) & Valido_Importe(rec!FPR_PRECIO) & Chr(9) & _
                                             Valido_Importe(rec!FPR_SUBTOT)
                   rec.MoveNext
                Loop
            End If
            rec.Close
            Set rec = Nothing
            CmdBorrar.Enabled = True
          Else
            'Falta el Nro. de Sucursal
            'NO puede pasar porque le pongo 1 por defecto
          End If
       Else
          'Falta el Nro. de Proveedor
          'NO puede pasar porque lo controlo antes
       End If
    Else
       'Falta el Nro. de Comprobante
       MsgBox "Ingrese el Nº de Comprobante a registrar!", vbExclamation, TIT_MSGBOX
       txtNroComprobante.SetFocus
    End If
End Sub

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
        txtNroSucursal.Text = "1"
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    SelecTexto txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPrecio, KeyAscii)
End Sub

Private Sub txtPrecio_LostFocus()
    If txtPrecio.Text = "" Then
        If txtcodigo.Text <> "" Then
            txtPrecio.Text = "0,00"
        End If
    Else
        txtPrecio.Text = Valido_Importe(txtPrecio.Text)
    End If
End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BusquedaDeProveedor txtProveedor, "CODIGO"
        txtProveedor_LostFocus
        txtProveedor.SetFocus
    End If
End Sub

Private Sub txtProvRazSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BusquedaDeProveedor txtCodProveedor, "CODIGO"
        txtCodProveedor_LostFocus
        txtCodProveedor.SetFocus
    End If
End Sub

Private Sub TxtTotal_GotFocus()
    seltxt
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTotal.Text, KeyAscii)
End Sub

Private Sub txtPendientes_GotFocus()
    seltxt
End Sub

Private Sub txtProveedor_Change()
    If txtProveedor.Text = "" Then txtDesProv.Text = ""
End Sub

Private Sub txtProveedor_GotFocus()
    SelecTexto txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtProveedor_LostFocus()
    If txtProveedor.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        SQL = "SELECT PROV_CODIGO, PROV_RAZSOC"
        SQL = SQL & " FROM PROVEEDOR"
        SQL = SQL & " WHERE PROV_CODIGO=" & XN(txtProveedor.Text)
        Rec1.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtDesProv.Text = Rec1!PROV_RAZSOC
        Else
            MsgBox "El Proveedor no existe", vbExclamation, TIT_MSGBOX
            txtProveedor.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtProvRazSoc_Change()
    If txtProvRazSoc.Text = "" Then
        txtCodProveedor.Text = ""
        txtCodTipoProv.Text = ""
    End If
End Sub

Private Sub txtProvRazSoc_GotFocus()
    SelecTexto txtProvRazSoc
End Sub

Private Sub txtProvRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtProvRazSoc_LostFocus()
    If txtCodProveedor.Text = "" And txtProvRazSoc.Text <> "" Then
        rec.Open BuscoProveedor(txtProvRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BusquedaDeProveedor txtCodProveedor, "CADENA", Trim(txtProvRazSoc.Text)
                txtCodProveedor_LostFocus
                txtProvRazSoc.SetFocus
            Else
                txtCodTipoProv.Text = rec!TPR_CODIGO
                txtCodProveedor.Text = rec!PROV_CODIGO
                txtProvRazSoc.Text = rec!PROV_RAZSOC
            End If
        Else
            MsgBox "No se encontró el Proveedor", vbExclamation, TIT_MSGBOX
            txtCodProveedor.SetFocus
        End If
        rec.Close
    ElseIf txtCodProveedor.Text = "" And txtProvRazSoc.Text = "" Then
        'MsgBox "Debe elegir un Proveedor", vbExclamation, TIT_MSGBOX
        MsgBox "Ingrese el Código de Proveedor a registrar!", vbExclamation, TIT_MSGBOX
        txtCodProveedor.SetFocus
    End If
End Sub

Private Function BuscoProveedor(Pro As String) As String
    SQL = "SELECT PROV_CODIGO, PROV_RAZSOC, TPR_CODIGO"
    SQL = SQL & " FROM PROVEEDOR "
    SQL = SQL & " WHERE "
    If txtCodProveedor.Text <> "" Then
        SQL = SQL & " PROV_CODIGO=" & XN(Pro)
    Else
        SQL = SQL & " PROV_RAZSOC LIKE '" & Pro & "%'"
    End If
    BuscoProveedor = SQL
End Function

Private Sub TxtTotal_LostFocus()
    If txtTotal.Text <> "" Then
        txtTotal.Text = Format(txtTotal, "0.00")
    Else
        txtTotal.Text = "0.00"
    End If
End Sub

Public Sub BuscarProducto(mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        'Set .Conn = DBConn
        cSQL = "SELECT PTO_DESCRI, PTO_CODIGO"
        cSQL = cSQL & " FROM PRODUCTO"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE"
            cSQL = cSQL & " PTO_DESCRI LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Descripción, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "PTO_DESCRI"
        campo1 = .Field
        .Field = "PTO_CODIGO"
        campo2 = .Field
        .OrderBy = "PTO_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Productos :"
        .MaxRecords = 1
        .Show
        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
                txtcodigo.Text = .ResultFields(2)
                TxtCodigo_LostFocus
        End If
    End With
    Set B = Nothing
End Sub

Public Sub BusquedaDeProveedor(Txt As TextBox, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        'Set .Conn = DBConn
        cSQL = "SELECT PROV_RAZSOC, PROV_CODIGO"
        cSQL = cSQL & " FROM PROVEEDOR"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE"
            cSQL = cSQL & " PROV_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Razón Social, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "PROV_RAZSOC"
        campo1 = .Field
        .Field = "PROV_CODIGO"
        campo2 = .Field
        .OrderBy = "PROV_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Proveedores :"
        .MaxRecords = 1
        .Show
        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
                Txt.Text = .ResultFields(2)
        End If
    End With
    Set B = Nothing
End Sub

