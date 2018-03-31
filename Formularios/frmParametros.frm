VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
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
   ScaleHeight     =   6735
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabDatos 
      Height          =   5835
      Left            =   45
      TabIndex        =   26
      Top             =   15
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   10292
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
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
      TabCaption(0)   =   "Sucursal "
      TabPicture(0)   =   "frmParametros.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame Frame5 
         Caption         =   "Horario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   51
         Top             =   3825
         Width           =   4560
         Begin VB.TextBox txtHasta 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3210
            MaxLength       =   5
            TabIndex        =   10
            Top             =   240
            Width           =   720
         End
         Begin VB.TextBox txtDesde 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1395
            MaxLength       =   5
            TabIndex        =   9
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta:"
            Height          =   315
            Left            =   2520
            TabIndex        =   53
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde:"
            Height          =   315
            Left            =   675
            TabIndex        =   52
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Aviso o Promoción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   50
         Top             =   4560
         Width           =   9615
         Begin VB.TextBox txtaviso 
            Height          =   750
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   9375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Condición Impositiva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   120
         TabIndex        =   37
         Top             =   1830
         Width           =   4560
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1335
            TabIndex        =   8
            Top             =   1590
            Width           =   1080
         End
         Begin VB.TextBox txtIngBrutos 
            Height          =   315
            Left            =   1335
            MaxLength       =   10
            TabIndex        =   4
            Top             =   225
            Width           =   1350
         End
         Begin VB.ComboBox cboIva 
            Height          =   315
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   915
            Width           =   3150
         End
         Begin FechaCtl.Fecha fechaInicio 
            Height          =   315
            Left            =   1335
            TabIndex        =   7
            Top             =   1260
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin MSMask.MaskEdBox txtCuit 
            Height          =   315
            Left            =   1335
            TabIndex        =   5
            Top             =   570
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   13
            Mask            =   "##-########-#"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "I.V.A.:"
            Height          =   315
            Left            =   105
            TabIndex        =   47
            Top             =   1590
            Width           =   1200
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio Actividad:"
            Height          =   285
            Left            =   105
            TabIndex        =   46
            Top             =   1260
            Width           =   1200
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IVA Condición:"
            Height          =   315
            Left            =   105
            TabIndex        =   45
            Top             =   915
            Width           =   1200
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ing. Brutos:"
            Height          =   315
            Left            =   105
            TabIndex        =   44
            Top             =   225
            Width           =   1200
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "C.U.I.T.:"
            Height          =   315
            Left            =   105
            TabIndex        =   43
            Top             =   570
            Width           =   1200
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Numeración Comprobantes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Left            =   4680
         TabIndex        =   27
         Top             =   1830
         Width           =   5070
         Begin VB.TextBox txtNroFacturaB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            TabIndex        =   14
            Top             =   1395
            Width           =   1080
         End
         Begin VB.TextBox txtNroRemito 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            TabIndex        =   12
            Top             =   705
            Width           =   1080
         End
         Begin VB.TextBox txtNroFacturaA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            TabIndex        =   13
            Top             =   1050
            Width           =   1080
         End
         Begin VB.TextBox txtNotaCreditoA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            TabIndex        =   15
            Top             =   1740
            Width           =   1080
         End
         Begin VB.TextBox txtNotaCreditoB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            TabIndex        =   16
            Top             =   2085
            Width           =   1080
         End
         Begin VB.TextBox txtNotaDebitoA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3900
            TabIndex        =   22
            Top             =   1740
            Width           =   1080
         End
         Begin VB.TextBox txtNotaDebitoB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3900
            TabIndex        =   21
            Top             =   1395
            Width           =   1080
         End
         Begin VB.TextBox txtNroReciboB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3900
            TabIndex        =   20
            Top             =   1050
            Width           =   1080
         End
         Begin VB.TextBox txtNroReciboA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3900
            TabIndex        =   19
            Top             =   705
            Width           =   1080
         End
         Begin VB.TextBox txtRecepcionMerca 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            TabIndex        =   11
            Top             =   360
            Width           =   1080
         End
         Begin VB.TextBox txtSalidaDeposito 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3900
            TabIndex        =   18
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. Factura C:"
            Height          =   315
            Left            =   90
            TabIndex        =   49
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. Remito:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   705
            Width           =   1365
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. Factura B:"
            Height          =   315
            Left            =   90
            TabIndex        =   36
            Top             =   1395
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nota Crédito C:"
            Height          =   315
            Left            =   90
            TabIndex        =   35
            Top             =   1740
            Width           =   1365
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nota Crédito B:"
            Height          =   315
            Left            =   90
            TabIndex        =   34
            Top             =   2085
            Width           =   1365
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nota Débito C:"
            Height          =   315
            Left            =   2625
            TabIndex        =   33
            Top             =   1740
            Width           =   1245
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nota Débito B:"
            Height          =   315
            Left            =   2625
            TabIndex        =   32
            Top             =   1395
            Width           =   1245
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. Recibo B:"
            Height          =   315
            Left            =   2625
            TabIndex        =   31
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. Recibo C:"
            Height          =   315
            Left            =   2625
            TabIndex        =   30
            Top             =   705
            Width           =   1245
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entrada Depósito:"
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Salida Depósito:"
            Height          =   315
            Left            =   2625
            TabIndex        =   28
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Generales...       "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   135
         TabIndex        =   38
         Top             =   405
         Width           =   9615
         Begin VB.TextBox txtSucursal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7770
            TabIndex        =   3
            Top             =   315
            Width           =   1080
         End
         Begin VB.TextBox txtRazSoc 
            Height          =   315
            Left            =   1215
            MaxLength       =   50
            TabIndex        =   0
            Top             =   315
            Width           =   4860
         End
         Begin VB.TextBox txtDireccion 
            Height          =   315
            Left            =   1215
            MaxLength       =   50
            TabIndex        =   1
            Top             =   660
            Width           =   4860
         End
         Begin VB.TextBox txtTelefono 
            Height          =   315
            Left            =   1215
            MaxLength       =   30
            TabIndex        =   2
            Top             =   1005
            Width           =   2070
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sucursal:"
            Height          =   315
            Left            =   6270
            TabIndex        =   42
            Top             =   315
            Width           =   1470
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Teléfono:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   1005
            Width           =   1125
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dirección:"
            Height          =   315
            Left            =   60
            TabIndex        =   40
            Top             =   660
            Width           =   1125
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Razón Social:"
            Height          =   315
            Left            =   60
            TabIndex        =   39
            Top             =   315
            Width           =   1125
         End
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmParametros.frx":001C
      Height          =   750
      Left            =   8160
      Picture         =   "frmParametros.frx":0326
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5895
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmParametros.frx":0630
      Height          =   750
      Left            =   9030
      Picture         =   "frmParametros.frx":093A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5895
      Width           =   870
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
      Left            =   165
      TabIndex        =   25
      Top             =   6120
      Width           =   660
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrabar_Click()
    If Validar_Parametros = False Then Exit Sub
    
    On Error GoTo HayError
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Actualizando..."
    
    DBConn.BeginTrans
    sql = "UPDATE PARAMETROS"
    sql = sql & " SET RAZ_SOCIAL=" & XS(txtRazSoc.Text)
    sql = sql & " ,DIRECCION=" & XS(txtDireccion.Text)
    sql = sql & " ,TELEFONO=" & XS(txtTelefono.Text)
    sql = sql & " ,CUIT=" & XS(txtCuit.Text)
    sql = sql & " ,ING_BRUTOS=" & XS(txtIngBrutos.Text)
    sql = sql & " ,IVA_CODIGO=" & cboIva.ItemData(cboIva.ListIndex)
    sql = sql & " ,INICIO_ACTIVIDAD=" & XDQ(fechaInicio.Text)
    sql = sql & " ,NRO_REMITO=" & XN(txtNroRemito.Text)
    sql = sql & " ,FACTURA_C=" & XN(txtNroFacturaA.Text)
    sql = sql & " ,FACTURA_B=" & XN(txtNroFacturaB.Text)
    sql = sql & " ,NOTA_CREDITO_C=" & XN(txtNotaCreditoA.Text)
    sql = sql & " ,NOTA_CREDITO_B=" & XN(txtNotaCreditoB.Text)
    sql = sql & " ,NOTA_DEBITO_C=" & XN(txtNotaDebitoA.Text)
    sql = sql & " ,NOTA_DEBITO_B=" & XN(txtNotaDebitoB.Text)
    sql = sql & " ,IVA=" & XN(txtIva.Text)
    sql = sql & " ,SUCURSAL=" & XN(txtSucursal.Text)
    sql = sql & " ,RECIBO_C=" & XN(txtNroReciboA.Text)
    sql = sql & " ,RECIBO_B=" & XN(txtNroReciboB.Text)
    sql = sql & " ,RECEPCION_MERCADERIA=" & XN(txtRecepcionMerca.Text)
    sql = sql & " ,SALIDA_MERCADERIA=" & XN(txtSalidaDeposito.Text)
    sql = sql & " ,AVISO =" & XS(txtaviso.Text)
    sql = sql & " ,HS_DESDE = #" & txtDesde.Text & "#"
    sql = sql & " ,HS_HASTA = #" & txtHasta.Text & "#"
    DBConn.Execute sql
    
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    Exit Sub
    
HayError:
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function Validar_Parametros() As Boolean
    If txtSucursal.Text = "" Then
        MsgBox "Debe Ingresar el número de Sucursal", vbExclamation, TIT_MSGBOX
        txtSucursal.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtRecepcionMerca.Text = "" Then
        MsgBox "Debe Ingresar el número de Recepción de Mercadería", vbExclamation, TIT_MSGBOX
        txtRecepcionMerca.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNroRemito.Text = "" Then
        MsgBox "Debe Ingresar el número de Remito", vbExclamation, TIT_MSGBOX
        txtNroRemito.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNroFacturaA.Text = "" Then
        MsgBox "Debe Ingresar el número de Factura A", vbExclamation, TIT_MSGBOX
        txtNroFacturaA.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNroFacturaB.Text = "" Then
        MsgBox "Debe Ingresar el número de Factura B", vbExclamation, TIT_MSGBOX
        txtNroFacturaB.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaCreditoA.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Crédito A", vbExclamation, TIT_MSGBOX
        txtNotaCreditoA.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaCreditoB.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Crédito B", vbExclamation, TIT_MSGBOX
        txtNotaCreditoB.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaDebitoA.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Débito A", vbExclamation, TIT_MSGBOX
        txtNotaDebitoA.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
    If txtNotaDebitoB.Text = "" Then
        MsgBox "Debe Ingresar el número de Nota de Débito B", vbExclamation, TIT_MSGBOX
        txtNotaDebitoB.SetFocus
        Validar_Parametros = False
        Exit Function
    End If
      Validar_Parametros = True
End Function

Private Sub cmdSalir_Click()
    Set frmParametros = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    tabDatos.Tab = 0
    'cargo combo iva
    LlenarComboIva
    'busco datos
    BuscarDatos
 
    lblEstado.Caption = ""
End Sub

Private Sub BuscarDatos()
    sql = "SELECT * FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        'DATOS
        txtRazSoc.Text = IIf(IsNull(rec!RAZ_SOCIAL), "", rec!RAZ_SOCIAL)
        txtDireccion.Text = IIf(IsNull(rec!DIRECCION), "", rec!DIRECCION)
        txtTelefono.Text = IIf(IsNull(rec!TELEFONO), "", rec!TELEFONO)
        txtCuit.Text = IIf(IsNull(rec!cuit), "", rec!cuit)
        txtIngBrutos.Text = IIf(IsNull(rec!ING_BRUTOS), "", rec!ING_BRUTOS)
        Call BuscaCodigoProxItemData(IIf(IsNull(rec!IVA_CODIGO), 1, rec!IVA_CODIGO), cboIva)
        fechaInicio.Text = IIf(IsNull(rec!INICIO_ACTIVIDAD), "", rec!INICIO_ACTIVIDAD)
        txtIva.Text = IIf(IsNull(rec!iva), "", Format(rec!iva, "0.00"))
        txtSucursal.Text = IIf(IsNull(rec!Sucursal), 1, rec!Sucursal)
        txtNroRemito.Text = IIf(IsNull(rec!NRO_REMITO), 1, rec!NRO_REMITO)
        txtNroFacturaA.Text = IIf(IsNull(rec!FACTURA_C), 1, rec!FACTURA_C)
        txtNroFacturaB.Text = IIf(IsNull(rec!FACTURA_B), 1, rec!FACTURA_B)
        txtNotaCreditoA.Text = IIf(IsNull(rec!NOTA_CREDITO_C), 1, rec!NOTA_CREDITO_C)
        txtNotaCreditoB.Text = IIf(IsNull(rec!NOTA_CREDITO_B), 1, rec!NOTA_CREDITO_B)
        txtNotaDebitoA.Text = IIf(IsNull(rec!NOTA_DEBITO_C), 1, rec!NOTA_DEBITO_C)
        txtNotaDebitoB.Text = IIf(IsNull(rec!NOTA_DEBITO_B), 1, rec!NOTA_DEBITO_B)
        txtNroReciboA.Text = IIf(IsNull(rec!RECIBO_C), 1, rec!RECIBO_C)
        txtNroReciboB.Text = IIf(IsNull(rec!RECIBO_B), 1, rec!RECIBO_B)
        txtRecepcionMerca.Text = IIf(IsNull(rec!RECEPCION_MERCADERIA), 1, rec!RECEPCION_MERCADERIA)
        txtSalidaDeposito.Text = IIf(IsNull(rec!SALIDA_MERCADERIA), 1, rec!SALIDA_MERCADERIA)
        txtaviso.Text = IIf(IsNull(rec!AVISO), "", rec!AVISO)
        txtDesde = IIf(IsNull(rec!HS_DESDE), "", Format(rec!HS_DESDE, "hh:mm"))
        txtHasta = IIf(IsNull(rec!HS_HASTA), "", Format(rec!HS_HASTA, "hh:mm"))
    End If
    rec.Close
End Sub

Private Sub LlenarComboIva()
    sql = "SELECT * FROM CONDICION_IVA ORDER BY IVA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboIva.AddItem rec!IVA_DESCRI
            cboIva.ItemData(cboIva.NewIndex) = rec!IVA_CODIGO
            rec.MoveNext
        Loop
        cboIva.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub txtCuit_GotFocus()
    SelecTexto txtCuit
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCuit_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtCuit.ClipText)) = 12 Then
      txtCuit.SelStart = 12
  End If
End Sub

Private Sub txtCuit_LostFocus()
    If txtCuit.Text <> "" Then
        'rutina de validación de CUIT
        If Not ValidoCuit(txtCuit) Then
            txtCuit.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDesde_GotFocus()
    SelecTexto txtDesde
End Sub

Private Sub txtDesde_LostFocus()
    txtDesde.Text = Format(txtDesde, "hh:mm")
End Sub

Private Sub txtDireccion_GotFocus()
    SelecTexto txtDireccion
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtHasta_GotFocus()
    SelecTexto txtHasta
End Sub

Private Sub txtHasta_LostFocus()
    txtDesde.Text = Format(txtDesde, "hh:mm")
End Sub

Private Sub txtIngBrutos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtIva_GotFocus()
   SelecTexto txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroDecimal(txtIva, KeyAscii)
End Sub

Private Sub txtIva_LostFocus()
    If txtIva.Text <> "" Then
        If ValidarPorcentaje(txtIva) = False Then txtIva.SetFocus
    End If
End Sub

Private Sub txtNotaCreditoA_GotFocus()
    SelecTexto txtNotaCreditoA
End Sub

Private Sub txtNotaCreditoA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNotaCreditoB_GotFocus()
    SelecTexto txtNotaCreditoB
End Sub

Private Sub txtNotaCreditoB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNotaDebitoA_GotFocus()
    SelecTexto txtNotaDebitoA
End Sub

Private Sub txtNotaDebitoA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNotaDebitoB_GotFocus()
    SelecTexto txtNotaDebitoB
End Sub

Private Sub txtNotaDebitoB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFacturaA_GotFocus()
    SelecTexto txtNroFacturaA
End Sub

Private Sub txtNroFacturaA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFacturaB_GotFocus()
    SelecTexto txtNroFacturaB
End Sub

Private Sub txtNroFacturaB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroReciboA_GotFocus()
    SelecTexto txtNroReciboA
End Sub

Private Sub txtNroReciboA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroReciboB_GotFocus()
    SelecTexto txtNroReciboB
End Sub

Private Sub txtNroReciboB_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroRemito_GotFocus()
    SelecTexto txtNroRemito
End Sub

Private Sub txtNroRemito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRecepcionMerca_GotFocus()
    SelecTexto txtRecepcionMerca
End Sub

Private Sub txtRecepcionMerca_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSalidaDeposito_GotFocus()
    SelecTexto txtSalidaDeposito
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto txtSucursal
End Sub

Private Sub txtTelefono_GotFocus()
    SelecTexto txtTelefono
End Sub
