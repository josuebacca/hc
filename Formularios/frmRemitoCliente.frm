VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmRemitoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito de Clientes..."
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8580
      TabIndex        =   10
      Top             =   7455
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10320
      TabIndex        =   12
      Top             =   7455
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7710
      TabIndex        =   9
      Top             =   7455
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9450
      TabIndex        =   11
      Top             =   7455
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7380
      Left            =   60
      TabIndex        =   22
      Top             =   45
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13018
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   512
      ForeColor       =   -2147483630
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
      TabPicture(0)   =   "frmRemitoCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameRemito"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FramePedido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmRemitoCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FramePedido 
         Caption         =   "Nota de Pedido ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   4050
         TabIndex        =   45
         Top             =   345
         Width           =   6990
         Begin VB.ComboBox cboRepresentada 
            Height          =   315
            Left            =   3375
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   750
            Width           =   3180
         End
         Begin VB.CommandButton cmdBuscarNotaPedido 
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
            Left            =   2205
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Buscar Nota de Pedido"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtNroNotaPedido 
            Height          =   315
            Left            =   990
            TabIndex        =   3
            Top             =   420
            Width           =   1125
         End
         Begin FechaCtl.Fecha FechaNotaPedido 
            Height          =   285
            Left            =   990
            TabIndex        =   4
            Top             =   765
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin MSFlexGridLib.MSFlexGrid grillaNotaPedido 
            Height          =   900
            Left            =   255
            TabIndex        =   46
            Top             =   1155
            Width           =   6450
            _ExtentX        =   11377
            _ExtentY        =   1588
            _Version        =   393216
            Rows            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   12648447
            BackColorBkg    =   -2147483633
            GridLinesFixed  =   1
            ScrollBars      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   2250
            TabIndex        =   62
            Top             =   825
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   270
            TabIndex        =   49
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   270
            TabIndex        =   47
            Top             =   450
            Width           =   615
         End
      End
      Begin VB.Frame FrameRemito 
         Caption         =   "Remito..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   120
         TabIndex        =   24
         Top             =   345
         Width           =   3915
         Begin VB.ComboBox cboRepRemito 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   225
            Width           =   2895
         End
         Begin VB.ComboBox cboListaPrecio 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1275
            Width           =   2895
         End
         Begin VB.TextBox txtNroSucursal 
            Enabled         =   0   'False
            Height          =   330
            Left            =   960
            MaxLength       =   4
            TabIndex        =   57
            Top             =   570
            Width           =   555
         End
         Begin VB.ComboBox cboStock 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1620
            Width           =   2895
         End
         Begin VB.TextBox txtNroRemito 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   21
            Top             =   570
            Width           =   1005
         End
         Begin FechaCtl.Fecha FechaRemito 
            Height          =   285
            Left            =   960
            TabIndex        =   13
            Top             =   945
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Stock:"
            Height          =   195
            Left            =   90
            TabIndex        =   54
            Top             =   1665
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Lst Precios:"
            Height          =   195
            Left            =   90
            TabIndex        =   51
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   90
            TabIndex        =   50
            Top             =   975
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   90
            TabIndex        =   44
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   90
            TabIndex        =   43
            Top             =   1965
            Width           =   555
         End
         Begin VB.Label lblEstadoRemito 
            AutoSize        =   -1  'True
            Caption         =   "EST. REMITO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   960
            TabIndex        =   42
            Top             =   1980
            Width           =   1050
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   52
         Top             =   2490
         Width           =   10935
         Begin VB.TextBox txtConcepto 
            Height          =   315
            Left            =   3105
            TabIndex        =   6
            Top             =   165
            Width           =   5160
         End
         Begin VB.CheckBox chkRemitoSinFactura 
            Alignment       =   1  'Right Justify
            Caption         =   "Remito sin Factura "
            Height          =   240
            Left            =   75
            TabIndex        =   5
            Top             =   225
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "<F1> Buscar Remitos"
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
            Left            =   8475
            TabIndex        =   58
            Top             =   210
            Width           =   2100
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Concepto:"
            Height          =   195
            Left            =   2295
            TabIndex        =   55
            Top             =   210
            Width           =   750
         End
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
         Height          =   1785
         Left            =   -74595
         TabIndex        =   30
         Top             =   480
         Width           =   10410
         Begin VB.TextBox txtOrden 
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
            Left            =   135
            TabIndex        =   63
            Text            =   "A"
            Top             =   645
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.CommandButton cmdBuscarVendedor 
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
            Left            =   3165
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Buscar Vendedor"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
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
            Left            =   3165
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarSuc 
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
            Left            =   3165
            MaskColor       =   &H000000FF&
            Picture         =   "frmRemitoCliente.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Buscar Sucursal"
            Top             =   615
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   315
            Left            =   2115
            TabIndex        =   16
            Top             =   945
            Width           =   990
         End
         Begin VB.TextBox txtDesVen 
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
            Left            =   3600
            TabIndex        =   37
            Top             =   960
            Width           =   4620
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "&Buscar"
            Height          =   360
            Left            =   8475
            MaskColor       =   &H000000FF&
            TabIndex        =   19
            ToolTipText     =   "Buscar "
            Top             =   945
            UseMaskColor    =   -1  'True
            Width           =   1665
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   4620
            TabIndex        =   18
            Top             =   1320
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   2115
            TabIndex        =   17
            Top             =   1320
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.TextBox txtDesCli 
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
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   315
            Left            =   2115
            MaxLength       =   40
            TabIndex        =   14
            Top             =   255
            Width           =   990
         End
         Begin VB.TextBox txtSucursal 
            Height          =   315
            Left            =   2115
            MaxLength       =   40
            TabIndex        =   15
            Top             =   600
            Width           =   990
         End
         Begin VB.TextBox txtDesSuc 
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
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   31
            Tag             =   "Descripción"
            Top             =   600
            Width           =   4620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   1020
            TabIndex        =   38
            Top             =   1005
            Width           =   750
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   3570
            TabIndex        =   36
            Top             =   1380
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1020
            TabIndex        =   35
            Top             =   1365
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   1020
            TabIndex        =   34
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1020
            TabIndex        =   33
            Top             =   660
            Width           =   660
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4320
         Left            =   -74610
         TabIndex        =   20
         Top             =   2355
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7620
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   262
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
         Height          =   4365
         Left            =   120
         TabIndex        =   25
         Top             =   2955
         Width           =   10935
         Begin VB.TextBox txtDeclarado 
            Alignment       =   1  'Right Justify
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
            Left            =   9270
            TabIndex        =   66
            Top             =   3660
            Width           =   990
         End
         Begin VB.TextBox txtBultos 
            Alignment       =   1  'Right Justify
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
            Left            =   7350
            TabIndex        =   64
            Top             =   3660
            Width           =   990
         End
         Begin VB.TextBox txtPeso 
            Alignment       =   1  'Right Justify
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
            Left            =   5640
            TabIndex        =   60
            Top             =   3660
            Width           =   990
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1365
            MaxLength       =   60
            TabIndex        =   8
            Top             =   4005
            Width           =   8895
         End
         Begin VB.CommandButton cmdBuscarProducto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":0C60
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":0F6A
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   510
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdQuitarProducto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10395
            MaskColor       =   &H8000000F&
            Picture         =   "frmRemitoCliente.frx":1274
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   855
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   195
            TabIndex        =   26
            Top             =   450
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3480
            Left            =   135
            TabIndex        =   7
            Top             =   180
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   6138
            _Version        =   393216
            Rows            =   3
            Cols            =   9
            FixedCols       =   0
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Declarado:"
            Height          =   195
            Left            =   8460
            TabIndex        =   67
            Top             =   3735
            Width           =   780
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bultos:"
            Height          =   195
            Left            =   6825
            TabIndex        =   65
            Top             =   3735
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Kg.:"
            Height          =   195
            Left            =   4890
            TabIndex        =   59
            Top             =   3735
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   195
            TabIndex        =   53
            Top             =   4050
            Width           =   1125
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
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
         Index           =   0
         Left            =   -74820
         TabIndex        =   23
         Top             =   570
         Width           =   1065
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
      Left            =   150
      TabIndex        =   41
      Top             =   7515
      Width           =   660
   End
End
Attribute VB_Name = "frmRemitoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim W As Integer
Dim TipoBusquedaDoc As Integer
Dim VEstadoRemito As Integer
Dim VStockPendiente As String
Dim VCodigoStock As String
Dim CantidadProducto As Integer
Dim VPedidoPendiente As Integer
Dim VBanderaBuscar As Boolean
Dim VBonificacion As Double
Dim VTotal As Double

Private Sub cboRepRemito_LostFocus()
    If ActiveControl.Name = "cmdNuevo" Or ActiveControl.Name = "cmdImprimir" _
       Or ActiveControl.Name = "CmdSalir" Or ActiveControl.Name = "cmdGrabar" Then Exit Sub
    
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    If VBanderaBuscar = False Then
        'Select Case cboRepRemito.ItemData(cboRepRemito.ListIndex)
        '    Case VRepresentada
        '        txtNroSucursal.Text = Sucursal
'            Case VRepresentada2
'                txtNroSucursal.Text = Sucursal2
'            Case VRepresentada3
'                txtNroSucursal.Text = Sucursal3
        '    Case Else
        '        MsgBox "La Representada seleccionada no es correcta", vbCritical, TIT_MSGBOX
        '        cboRepRemito.ListIndex = 0
        '        cboRepRemito.SetFocus
        '        Exit Sub
        'End Select
        txtNroRemito.Text = BuscoUltimoNumeroComprobante(cboRepRemito.ItemData(cboRepRemito.ListIndex), 0)
    End If
End Sub

Private Sub chkRemitoSinFactura_Click()
    If chkRemitoSinFactura.Value = Checked Then
        txtConcepto.Enabled = True
    Else
        txtConcepto.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
'    ABMProducto.Show vbModal
'    grdGrilla.SetFocus
'    grdGrilla.row = 1
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Select Case TipoBusquedaDoc
    
    Case 1 'BUSCA REMITOS
        
        sql = "SELECT RC.*, C.CLI_RAZSOC, S.SUC_DESCRI, V.VEN_NOMBRE,NP.REP_CODIGO"
        sql = sql & " FROM REMITO_CLIENTE RC,CLIENTE C, SUCURSAL S,"
        sql = sql & " NOTA_PEDIDO NP, VENDEDOR V"
        sql = sql & " WHERE"
        sql = sql & " RC.NPE_NUMERO=NP.NPE_NUMERO"
        sql = sql & " AND RC.NPE_FECHA=NP.NPE_FECHA"
        sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
        sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
        sql = sql & " AND NP.VEN_CODIGO=V.VEN_CODIGO"
        If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
        If txtSucursal.Text <> "" Then sql = sql & " AND NP.SUC_CODIGO=" & XN(txtSucursal)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If FechaDesde <> "" Then sql = sql & " AND RC.RCL_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND RC.RCL_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY RC.RCL_NUMERO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem Format(rec!RCL_SUCURSAL, "0000") & "-" & Format(rec!RCL_NUMERO, "00000000") _
                                & Chr(9) & rec!RCL_FECHA _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!SUC_DESCRI _
                                & Chr(9) & rec!VEN_NOMBRE & Chr(9) & rec!EST_CODIGO _
                                & Chr(9) & rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA _
                                & Chr(9) & rec!RCL_OBSERVACION & Chr(9) & rec!STK_CODIGO _
                                & Chr(9) & rec!RCL_SINFAC & Chr(9) & rec!RCL_CONCEPTO _
                                & Chr(9) & rec!REP_CODIGO
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
            GrdModulos.Col = 0
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
        
    Case 2 'BUSCA NOTA DE PEDIDO
        
        sql = "SELECT NP.NPE_NUMERO, NP.NPE_FECHA, C.CLI_RAZSOC, S.SUC_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, SUCURSAL S"
        sql = sql & " WHERE"
        sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
        sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
        sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
        sql = sql & " AND NP.EST_CODIGO=1" 'NOTA DE PEDIDO PENDIENTES
        If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
        If txtSucursal.Text <> "" Then sql = sql & " AND NP.SUC_CODIGO=" & XN(txtSucursal)
        If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
        If FechaDesde <> "" Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
        If FechaHasta <> "" Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
        sql = sql & " ORDER BY NPE_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Do While rec.EOF = False
                GrdModulos.AddItem rec!NPE_NUMERO & Chr(9) & rec!NPE_FECHA _
                                & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!SUC_DESCRI
                rec.MoveNext
            Loop
            GrdModulos.SetFocus
            GrdModulos.Col = 0
        Else
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            MsgBox "No se encontraron datos...", vbExclamation, TIT_MSGBOX
        End If
    End Select
    
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    rec.Close
End Sub

Private Sub cmdBuscarCli_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCliente.Text = frmBuscar.grdBuscar.Text
        txtCliente.SetFocus
        txtCliente_LostFocus
    Else
        txtCliente.SetFocus
    End If
End Sub



Private Sub cmdBuscarNotaPedido_Click()
    TipoBusquedaDoc = 2
    tabDatos.Tab = 1
End Sub

Private Sub cmdBuscarProducto_Click()
    grdGrilla.SetFocus
    frmBuscar.TipoBusqueda = 2
    frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    grdGrilla.Col = 0
    EDITAR grdGrilla, txtEdit, 13
    If Trim(frmBuscar.grdBuscar.Text) <> "" Then txtEdit.Text = frmBuscar.grdBuscar.Text
    TxtEdit_KeyDown vbKeyReturn, 0
End Sub

Private Sub cmdBuscarSuc_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.TxtDescriB = ""
    If txtCliente.Text <> "" Then
        frmBuscar.CodigoCli = txtCliente.Text
    Else
        frmBuscar.CodigoCli = ""
    End If
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 3
        txtCliente.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 0
        txtSucursal.Text = frmBuscar.grdBuscar.Text
        txtSucursal.SetFocus
        txtSucursal_LostFocus
    Else
        txtSucursal.SetFocus
    End If
End Sub

Private Sub cmdBuscarVendedor_Click()
    frmBuscar.TipoBusqueda = 4
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtVendedor.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtDesVen.Text = frmBuscar.grdBuscar.Text
        txtVendedor.SetFocus
    Else
        txtVendedor.SetFocus
    End If
End Sub

Private Sub CmdGrabar_Click()
    
    If ValidarRemito = False Then Exit Sub
    If MsgBox("¿Confirma Remito?" & Chr(13) & Chr(13) & _
            "Sucursal: " & cboRepRemito.List(cboRepRemito.ListIndex) & Chr(13) & _
            "Número:  " & txtNroSucursal.Text & "-" & txtNroRemito.Text, vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorRemito
    
    DBConn.BeginTrans
    
    sql = "SELECT * FROM REMITO_CLIENTE"
    sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
    sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then '----------NUEVO REMITO--------------------
            
        sql = "INSERT INTO REMITO_CLIENTE"
        sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,NPE_NUMERO,"
        sql = sql & "NPE_FECHA,RCL_OBSERVACION,"
        sql = sql & "EST_CODIGO,RCL_SINFAC,RCL_CONCEPTO,RCL_NUMEROTXT,STK_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroRemito) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaRemito) & ","
        sql = sql & XN(txtNroNotaPedido) & ","
        sql = sql & XDQ(FechaNotaPedido) & ","
        sql = sql & XS(txtObservaciones) & ","
        If chkRemitoSinFactura.Value = Checked Then
            sql = sql & "3,"   'ESTADO DEFINITIVO
            sql = sql & "'S'," 'REMITO SIN FACTURA
            sql = sql & XS(txtConcepto.Text) & "," 'CONCEPTO DEL REMITO SIN FACTURA
        Else
            sql = sql & "1,"    'ESTADO PENDIENTE
            sql = sql & "'N',"  'REMITO CON FACTURA
            sql = sql & "NULL," 'CONCEPTO DEL REMITO SIN FACTURA
        End If
        sql = sql & XS(Format(txtNroRemito.Text, "00000000")) & ","
        sql = sql & cboStock.ItemData(cboStock.ListIndex) & ")" 'STOCK
        DBConn.Execute sql
           
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                
                CantidadProducto = 0
                'ME FIJO SI LA CANTIDAD ES IGUAL QUE EN LA NOTA DE PEDIDO
                sql = "SELECT PTO_CODIGO, DNP_CANTIDAD AS CANTIDAD"
                sql = sql & " FROM DETALLE_NOTA_PEDIDO"
                sql = sql & " WHERE"
                sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    CantidadProducto = Rec1!Cantidad
                End If
                Rec1.Close
                
                'SI LA NOTA DE PEDIDO NO TIENE EL PRODUCTO NO AGO EL PASO SUGUIENTE
                If CantidadProducto > 0 Then
                    sql = " SELECT DR.PTO_CODIGO, SUM(DR.DRC_CANTIDAD) AS CANTIDAD"
                    sql = sql & " FROM DETALLE_REMITO_CLIENTE DR, REMITO_CLIENTE R"
                    sql = sql & " WHERE"
                    sql = sql & " R.NPE_NUMERO=" & XN(txtNroNotaPedido)
                    sql = sql & " AND R.NPE_FECHA=" & XDQ(FechaNotaPedido)
                    sql = sql & " AND R.RCL_NUMERO=DR.RCL_NUMERO"
                    sql = sql & " AND R.RCL_SUCURSAL=DR.RCL_SUCURSAL"
                    sql = sql & " AND DR.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                    sql = sql & " AND R.EST_CODIGO <> 2" 'BUSCA EN LOS REMITOS NO ANULADOS
                    sql = sql & " GROUP BY DR.PTO_CODIGO"
                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                        CantidadProducto = CantidadProducto - CInt(Rec1!Cantidad)
                    End If
                    Rec1.Close
                
                    If CInt(grdGrilla.TextMatrix(i, 2)) >= CantidadProducto Then
                        'MARCO LOS PRODUCTOS UTILIZADOS EN EL REMITO
                        sql = "UPDATE DETALLE_NOTA_PEDIDO"
                        sql = sql & " SET DNP_MARCA='X'"
                        sql = sql & " WHERE"
                        sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                        sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                        sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                        DBConn.Execute sql
                    End If
                End If
                
                'INSERTO EN EL DETALLE DEL REMITO
                sql = "INSERT INTO DETALLE_REMITO_CLIENTE"
                sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,DRC_NROITEM,"
                sql = sql & "PTO_CODIGO,DRC_CANTIDAD,DRC_PRECIO,DRC_BONIFICA,DRC_COSTO)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroRemito.Text) & ","
                sql = sql & XN(txtNroSucursal.Text) & ","
                sql = sql & XDQ(FechaRemito.Text) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 4)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 8)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        'ACTUALIZO EL STOCK (STOCK PENDIENTE)
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                sql = "UPDATE DETALLE_STOCK"
                sql = sql & " SET"
                sql = sql & " DST_STKPEN = DST_STKPEN + " & XN(grdGrilla.TextMatrix(i, 2))
                sql = sql & " WHERE STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
                sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                DBConn.Execute sql
            End If
        Next
    
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO AL REMITO
        Call ActualizoNumeroComprobantes(cboRepRemito.ItemData(cboRepRemito.ListIndex), 0, txtNroRemito.Text)
        
        DBConn.CommitTrans
        
        DBConn.BeginTrans
        sql = "SELECT PTO_CODIGO FROM DETALLE_NOTA_PEDIDO"
        sql = sql & " WHERE"
        sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
        sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
        sql = sql & " AND DNP_MARCA IS NULL"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = True Then
            'CAMBIO ESTADO DE LA NOTA DE PEDIDO (LE PONGO DEFINITIVO)
            sql = "UPDATE NOTA_PEDIDO SET EST_CODIGO=3"
            sql = sql & " WHERE"
            sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
        End If
        Rec1.Close
        
        DBConn.CommitTrans
        
    Else '---------MODIFICA EL REMITO-----------------------------------
    
        If MsgBox("¿Seguro que modifica el Remito Nro: " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroRemito.Text) _
                   , vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
            
            'CONTROLO QUE EL REMITO NO FUE UTILIZADO EN SALIDA DE DEPOSITO
            sql = "SELECT EGA_CODIGO,EGA_FECHA"
            sql = sql & " FROM ENTREGA_PRODUCTO"
            sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                MsgBox "El Remito no puede ser modificado ya que fue utilizado" & Chr(13) & _
                       "en la Salida de Mercadería Nro: " & Rec1!EGA_CODIGO & " del " & Rec1!EGA_FECHA, vbCritical, TIT_MSGBOX
                rec.Close
                Rec1.Close
                DBConn.CommitTrans
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                Exit Sub
            End If
            Rec1.Close
            
            'VUELVO ATRAS EL STOCK (STOCK PENDIENTE)
            VCodigoStock = ""
            sql = "SELECT STK_CODIGO FROM REMITO_CLIENTE"
            sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito.Text)
            sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal.Text)
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec2.EOF = False Then
                VCodigoStock = Rec2!STK_CODIGO
            Else
                VCodigoStock = "0"
            End If
            Rec2.Close
            
            VStockPendiente = ""
            'Set Rec2 = New ADODB.Recordset
            sql = "SELECT DS.STK_CODIGO, DS.PTO_CODIGO, DS.DST_STKPEN, DR.DRC_CANTIDAD"
            sql = sql & " FROM DETALLE_STOCK DS, DETALLE_REMITO_CLIENTE DR"
            sql = sql & " WHERE"
            sql = sql & " DS.STK_CODIGO=" & XN(VCodigoStock)
            sql = sql & " AND DR.RCL_NUMERO=" & XN(txtNroRemito.Text)
            sql = sql & " AND DR.RCL_SUCURSAL=" & XN(txtNroSucursal.Text)
            sql = sql & " AND DS.PTO_CODIGO=DR.PTO_CODIGO"
            Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
            If Rec2.EOF = False Then
                Do While Rec2.EOF = False
                    VStockPendiente = CStr(CInt(Rec2!DST_STKPEN) - CInt(Rec2!DRC_CANTIDAD))
                    sql = "UPDATE DETALLE_STOCK"
                    sql = sql & " SET"
                    sql = sql & " DST_STKPEN=" & XN(VStockPendiente)
                    sql = sql & " WHERE STK_CODIGO=" & XN(VCodigoStock)
                    sql = sql & " AND PTO_CODIGO=" & XN(Rec2!PTO_CODIGO)
                    DBConn.Execute sql
                    
                    'LE SACO LA MARCA A LOS PRODUCTOS EN LA NOTA DE PEDIDO
                    sql = "UPDATE DETALLE_NOTA_PEDIDO"
                    sql = sql & " SET DNP_MARCA=NULL"
                    sql = sql & " WHERE"
                    sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                    sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                    sql = sql & " AND PTO_CODIGO=" & XN(Rec2!PTO_CODIGO)
                    DBConn.Execute sql
                    
                    Rec2.MoveNext
                Loop
            End If
            Rec2.Close
            '----HASTA ACA LA VUELTA ATRAS DEL STOCK-----
            '---AHORA COMIENZA LA MODIFICACION DEL REMITO--
            
            'ACTUALIZA REMITO_CLIENTE
            sql = "UPDATE REMITO_CLIENTE"
            sql = sql & " SET RCL_OBSERVACION=" & XS(txtObservaciones.Text)
            If chkRemitoSinFactura.Value = Checked Then
                sql = sql & " ,EST_CODIGO= 3" 'ESTADO DEFINITIVO
                sql = sql & " ,RCL_SINFAC='S'"
                sql = sql & " ,RCL_CONCEPTO=" & XS(txtConcepto.Text)
            Else
                sql = sql & " ,EST_CODIGO=1" 'ESTADO PENDIENTE
                sql = sql & " ,RCL_SINFAC='N'"
                sql = sql & " ,RCL_CONCEPTO=NULL"
            End If
            sql = sql & " ,RCL_NUMEROTXT=" & XS(Format(txtNroRemito.Text, "00000000"))
            sql = sql & " ,STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
            sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
            DBConn.Execute sql
            
            '-BORRO EL DETALLE DEL REMITO Y LO INSERTO DE NUEVO---------
            sql = "DELETE FROM DETALLE_REMITO_CLIENTE"
            sql = sql & " WHERE RCL_NUMERO=" & XN(txtNroRemito)
            sql = sql & " AND RCL_SUCURSAL=" & XN(txtNroSucursal)
            DBConn.Execute sql
            
            For i = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(i, 0) <> "" Then
                    CantidadProducto = 0
                    'ME FIJO SI LA CANTIDAD ES IGUAL QUE EN LA NOTA DE PEDIDO
                    sql = "SELECT PTO_CODIGO, DNP_CANTIDAD AS CANTIDAD"
                    sql = sql & " FROM DETALLE_NOTA_PEDIDO"
                    sql = sql & " WHERE"
                    sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                    sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                    sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If Rec1.EOF = False Then
                        CantidadProducto = Rec1!Cantidad
                    End If
                    Rec1.Close
                    
                    'SI LA NOTA DE PEDIDO NO TIENE EL PRODUCTO NO AGO EL PASO SUGUIENTE
                    If CantidadProducto > 0 Then
                        sql = " SELECT DR.PTO_CODIGO, SUM(DR.DRC_CANTIDAD) AS CANTIDAD"
                        sql = sql & " FROM DETALLE_REMITO_CLIENTE DR, REMITO_CLIENTE R"
                        sql = sql & " WHERE"
                        sql = sql & " R.NPE_NUMERO=" & XN(txtNroNotaPedido)
                        sql = sql & " AND R.NPE_FECHA=" & XDQ(FechaNotaPedido)
                        sql = sql & " AND R.RCL_NUMERO=DR.RCL_NUMERO"
                        sql = sql & " AND R.RCL_SUCURSAL=DR.RCL_SUCURSAL"
                        sql = sql & " AND DR.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                        sql = sql & " AND R.EST_CODIGO <> 2" 'BUSCA EN LOS REMITOS NO ANULADOS
                        sql = sql & " GROUP BY DR.PTO_CODIGO"
                        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                        If Rec1.EOF = False Then
                            CantidadProducto = CantidadProducto - CInt(Rec1!Cantidad)
                        End If
                        Rec1.Close
                    
                        If CInt(grdGrilla.TextMatrix(i, 2)) >= CantidadProducto Then
                            'MARCO LOS PRODUCTOS UTILIZADOS EN EL REMITO
                            sql = "UPDATE DETALLE_NOTA_PEDIDO"
                            sql = sql & " SET DNP_MARCA='X'"
                            sql = sql & " WHERE"
                            sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                            sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                            DBConn.Execute sql
                        End If
                    End If
                    
                    'INSERTO EN EL DETALLE DEL REMITO
                    sql = "INSERT INTO DETALLE_REMITO_CLIENTE"
                    sql = sql & " (RCL_NUMERO,RCL_SUCURSAL,RCL_FECHA,DRC_NROITEM,"
                    sql = sql & "PTO_CODIGO,DRC_CANTIDAD,DRC_PRECIO,DRC_BONIFICA,DRC_COSTO)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroRemito.Text) & ","
                    sql = sql & XN(txtNroSucursal.Text) & ","
                    sql = sql & XDQ(FechaRemito.Text) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 2)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 4)) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 8)) & ")"
                    DBConn.Execute sql
                End If
            Next
            
            'ACTUALIZO EL STOCK (STOCK PENDIENTE)
            For i = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(i, 0) <> "" Then
                    sql = "UPDATE DETALLE_STOCK"
                    sql = sql & " SET"
                    sql = sql & " DST_STKPEN = DST_STKPEN + " & XN(grdGrilla.TextMatrix(i, 2))
                    sql = sql & " WHERE STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
                    sql = sql & " AND PTO_CODIGO=" & XN(grdGrilla.TextMatrix(i, 0))
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
        
            DBConn.BeginTrans
            sql = "SELECT PTO_CODIGO FROM DETALLE_NOTA_PEDIDO"
            sql = sql & " WHERE"
            sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP_MARCA IS NULL"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = True Then
                'CAMBIO ESTADO DE LA NOTA DE PEDIDO (LE PONGO DEFINITIVO)
                sql = "UPDATE NOTA_PEDIDO SET EST_CODIGO=3"
                sql = sql & " WHERE"
                sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                DBConn.Execute sql
            Else
                'CAMBIO ESTADO DE LA NOTA DE PEDIDO (LE PONGO PENDIENTE)
                'YA QUE LE QUEDAN PRODUCTOS PARA ASIGNAR
                sql = "UPDATE NOTA_PEDIDO SET EST_CODIGO=1"
                sql = sql & " WHERE"
                sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido)
                sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
                DBConn.Execute sql
            End If
            Rec1.Close
            
            DBConn.CommitTrans
        Else 'SI NO MODIFICO TERMINA EL BIGINN
            DBConn.CommitTrans
        End If
    End If '----FIN DEL REMITO-----
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    CmdNuevo_Click
    Exit Sub
    
HayErrorRemito:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    If Rec1.State = 1 Then Rec1.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarRemito() As Boolean
    
    If FechaRemito.Text = "" Then
        MsgBox "La Fecha del Remito es requerida", vbExclamation, TIT_MSGBOX
        FechaRemito.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If txtNroNotaPedido.Text = "" Then
        MsgBox "El número de Nota de Pedido es requerido", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If FechaNotaPedido.Text = "" Then
        MsgBox "La Fecha de la Nota de pedido es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaPedido.SetFocus
        ValidarRemito = False
        Exit Function
    End If
    If chkRemitoSinFactura.Value = Checked Then
        If txtConcepto.Text = "" Then
            MsgBox "Debe ingresar un concepto", vbExclamation, TIT_MSGBOX
            txtConcepto.SetFocus
            ValidarRemito = False
            Exit Function
        End If
    End If
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 0) <> "" Then
            If ESTA_EN_STOCK(grdGrilla.TextMatrix(i, 0), CStr(cboStock.ItemData(cboStock.ListIndex)), _
                             grdGrilla.TextMatrix(i, 1)) = False Then
                ValidarRemito = False
                Exit Function
            End If
        End If
    Next
    ValidarRemito = True
End Function

Private Sub cmdImprimir_Click()
    If txtNroSucursal.Text = "" Or txtNroRemito.Text = "" _
       Or txtNroNotaPedido.Text = "" Or FechaNotaPedido.Text = "" Then Exit Sub
       
    If MsgBox("¿Confirma Impresión Remito?" & Chr(13) & Chr(13) & _
            "Sucursal: " & cboRepRemito.List(cboRepRemito.ListIndex) & Chr(13) & _
            "Número:  " & txtNroSucursal.Text & "-" & txtNroRemito.Text, vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
            
'PONE A LA IMPRESORA  COMO PREDETERMINADA
    Dim X As Printer
    Dim mDriver As String
    mDriver = Impresora
    For Each X In Printers
        If X.DeviceName = mDriver Then
            ' La define como predeterminada del sistema.
            Set Printer = X
            Exit For
        End If
    Next
'-----------------------------------
    Set_Impresora
    ImprimirRemito
End Sub

Public Sub ImprimirEncabezado()
    '---------------IMPRIME EL ENCABEZADO DEL REMITO-------------------
    Printer.FontSize = 10
    Imprimir 13.8, 0.6, True, "REMITO Nº  " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroRemito.Text)
    Imprimir 15.5, 2.1, False, Format(FechaRemito, "dd/mm/yyyy")
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI, NP.VEN_CODIGO, V.VEN_NOMBRE"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L, NOTA_PEDIDO NP, "
    sql = sql & " PROVINCIA P, CONDICION_IVA CI, VENDEDOR V"
    sql = sql & " WHERE  NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
    sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
    sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"
    sql = sql & " AND NP.VEN_CODIGO=V.VEN_CODIGO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Imprimir 1.3, 4.6, True, "(" & Trim(Rec1!CLI_CODIGO) & ") " & Trim(Rec1!CLI_RAZSOC) & _
                               "  - Vend: (" & Trim(Rec1!VEN_CODIGO) & ")" & Trim(Rec1!VEN_NOMBRE)
                               
        Imprimir 1.3, 5, False, Trim(Rec1!CLI_DOMICI)
        'nota de pedido
        Imprimir 13.3, 5, True, "Nro.Pedido: " & Format(txtNroNotaPedido.Text, "00000000")
        Imprimir 1.3, 5.4, False, "Loc: " & Trim(Rec1!LOC_DESCRI) & " -- Prov: " & Trim(Rec1!PRO_DESCRI)
        'fecha nota pedido
        Imprimir 13.3, 5.4, True, "Fecha: " & Format(FechaNotaPedido.Text, "dd/mm/yyyy")
        Imprimir 1.7, 6.2, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 7.9, 6.2, False, IIf(IsNull(Rec1!CLI_CUIT), "NO INFORMADO", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 15.7, 6.2, False, IIf(IsNull(Rec1!CLI_INGBRU), "NO INFORMADO", Format(Rec1!CLI_INGBRU, "###-#####-##"))
    End If
    Rec1.Close
    'Imprimir 18.4, 7.9, False, CStr(VCantidadBultos)
    'Imprimir 0, 9.1, False, "Código"
    'Imprimir 2, 9.1, False, "Descripción"
    'Imprimir 13, 9.1, False, "Cant."
    'Imprimir 15, 9.1, False, "Rubro"
End Sub

Public Sub ImprimirRemito()
    Dim Renglon As Double
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For W = 1 To 3 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DEL REINTEGRO ------------------
        Renglon = 9.3
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                Printer.FontSize = 8
                Imprimir 1.3, Renglon, False, Format(grdGrilla.TextMatrix(i, 0), "000000")  'codigo
                If Len(grdGrilla.TextMatrix(i, 1)) < 60 Then
                    Imprimir 6.3, Renglon, False, Trim(grdGrilla.TextMatrix(i, 1))   'descripcion
                Else
                    Imprimir 6.3, Renglon, False, Trim(Left(grdGrilla.TextMatrix(i, 1), 59)) & "..." 'descripcion
                End If
                Printer.FontSize = 9
                Imprimir 4.3, Renglon, False, Trim(grdGrilla.TextMatrix(i, 2)) 'cantidad
                'Imprimir 15, Renglon, False, Trim(Left(grdGrilla.TextMatrix(I, 4), 20)) 'rubro
                Renglon = Renglon + 0.5
            End If
        Next i
        Printer.FontSize = 9
        '-----OBSERVACIONES------------------------------------------
        If txtObservaciones.Text <> "" Then
            Imprimir 1.2, Renglon + 1.5, False, "Observaciones: " & Trim(txtObservaciones.Text)
        End If
        'txtObservaciones
          '------------DATOS SUCURSAL-------------------------
           sql = "SELECT S.SUC_DESCRI,S.SUC_DOMICI, L.LOC_DESCRI, S.SUC_ENTREGAR"
           sql = sql & " FROM SUCURSAL S, NOTA_PEDIDO NP, LOCALIDAD L, CLIENTE C"
           sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido.Text)
           sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido.Text)
           sql = sql & " AND NP.CLI_CODIGO=C.CLI_CODIGO"
           sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
           sql = sql & " AND S.CLI_CODIGO=C.CLI_CODIGO"
           sql = sql & " AND S.LOC_CODIGO=L.LOC_CODIGO"
           sql = sql & " AND S.PRO_CODIGO=L.PRO_CODIGO"
           sql = sql & " AND S.PAI_CODIGO=L.PAI_CODIGO"
           Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
           If Rec1.EOF = False Then
                Printer.FontSize = 8
                Imprimir 1.2, 21, True, "Entregar:  " & Left(Trim(Rec1!SUC_DESCRI), 25) & " -- " & Left(Trim(Rec1!SUC_DOMICI), 30) & " (" & Left(Trim(Rec1!LOC_DESCRI), 20) & ")"
                Imprimir 2.6, 21.4, True, Trim(ChkNull(Rec1!SUC_ENTREGAR))
                Printer.FontSize = 9
           End If
           Rec1.Close
          '---------PORCUENTA Y ORDEN--------------------------
           Imprimir 1.2, 21.9, True, "Por Cuenta y Orden de: " & Trim(cboRepresentada.List(cboRepresentada.ListIndex))
          
            Printer.FontSize = 10
          '--------- BULTOS -------------------------------------------
            Imprimir 10, 23.2, True, "      Total Bultos:  " & CStr(Trim(txtBultos.Text))
          '--------- TOTAL KILOGRAMOS ---------------------------------
            Imprimir 10, 23.7, True, "           Total Kg.:  " & CStr(Trim(txtPeso.Text))
          '--------- TOTAL DECLARADO ----------------------------------
            Imprimir 10, 24.2, True, "Total Declarado:  $ " & CStr(Trim(txtDeclarado.Text))
            
        Printer.EndDoc
    Next W
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    'significa que no estoy buacando
    VBanderaBuscar = False
    'LIMPIO REMITO
    Limpiar_Remito
    'BUSCO EL NUMERO DE REMITO QUE CORRESPONDE
    'txtNroRemito.Text = BuscoUltimoRenito
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    '--------------
    tabDatos.Tab = 0
    TipoBusquedaDoc = 1
End Sub

Private Sub Limpiar_Remito()
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 0) = ""
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.TextMatrix(i, 5) = ""
        grdGrilla.TextMatrix(i, 6) = ""
        grdGrilla.TextMatrix(i, 7) = i
        grdGrilla.TextMatrix(i, 8) = ""
    Next
    grillaNotaPedido.TextMatrix(0, 1) = ""
    grillaNotaPedido.TextMatrix(1, 1) = ""
    grillaNotaPedido.TextMatrix(2, 1) = ""
    FechaNotaPedido.Text = ""
    txtNroNotaPedido.Text = ""
    cboRepresentada.ListIndex = -1
    chkRemitoSinFactura.Value = Unchecked
    txtConcepto.Text = ""
    lblEstadoRemito.Caption = ""
    txtObservaciones.Text = ""
    lblEstado.Caption = ""
    txtBultos.Text = ""
    txtPeso.Text = ""
    txtDeclarado.Text = ""
    cmdGrabar.Enabled = True
    '--------------
    txtNroSucursal.Text = ""
    txtNroRemito.Text = ""
    FrameRemito.Enabled = True
    FramePedido.Enabled = True
    '--------------
    FechaRemito.Text = Date
    cboListaPrecio.ListIndex = 0
    cboRepRemito.ListIndex = 0
    cboRepRemito.SetFocus
End Sub

Private Sub cmdQuitarProducto_Click()
    If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = grdGrilla.RowSel
        
        'RESTO LOS BULTOS
        txtBultos.Text = SumaBultos
        txtPeso.Text = SumaPeso
        txtDeclarado.Text = SumaDeclarado
        grdGrilla.SetFocus
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmRemitoCliente = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        TipoBusquedaDoc = 1
        tabDatos.Tab = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "grdGrilla" And _
        Me.ActiveControl.Name <> "txtEdit" And _
        KeyAscii = vbKeyReturn Then
        MySendKeys Chr(9)
    End If
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)

    grdGrilla.FormatString = "Código|Descripción|>Cantidad|>Precio|>Bonif.|>Importe|Peso|Orden|PRECIO COSTO"
    grdGrilla.ColWidth(0) = 900  'CODIGO
    grdGrilla.ColWidth(1) = 5200 'DESCRIPCION
    grdGrilla.ColWidth(2) = 900  'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 900  'BONOFICACION
    grdGrilla.ColWidth(5) = 1100 'IMPORTE
    grdGrilla.ColWidth(6) = 0    'PESO
    grdGrilla.ColWidth(7) = 0    'ORDEN
    grdGrilla.ColWidth(8) = 0    'PRECIO COSTO
    grdGrilla.Cols = 9
    grdGrilla.Rows = 1
    '-----------------------
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To 8
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    '-----------------------
    For i = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & "" & Chr(9) & "" & Chr(9) & (i - 1) & Chr(9) & ""
    Next
    
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Número|^Fecha|Cliente|Sucursal|Vendedor|Cod Estado|NP NUMERO|NP FECHA|OBSERVACIONES|" _
                              & "STOCK|REMITO SIN FACTURA|CONCEPTO|REPRESENTADA"
    GrdModulos.ColWidth(0) = 1300 'NUMERO
    GrdModulos.ColWidth(1) = 1000 'FECHA
    GrdModulos.ColWidth(2) = 3950 'CLIENTE
    GrdModulos.ColWidth(3) = 3950 'SUCURSAL
    GrdModulos.ColWidth(4) = 0    'VENDEDOR
    GrdModulos.ColWidth(5) = 0    'COD ESTADO
    GrdModulos.ColWidth(6) = 0    'NOTA PEDIDO NUMERO
    GrdModulos.ColWidth(7) = 0    'NOTA PEDIDO FECHA
    GrdModulos.ColWidth(8) = 0    'OBSERVACIONES
    GrdModulos.ColWidth(9) = 0    'STOCK
    GrdModulos.ColWidth(10) = 0   'REMITO SIN FACTURAS
    GrdModulos.ColWidth(11) = 0   'CONCEPTO
    GrdModulos.ColWidth(12) = 0   'REPRESENTADA
    GrdModulos.Rows = 1
    
    GrdModulos.BorderStyle = flexBorderNone
    
    GrdModulos.row = 0
    For i = 0 To 3
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    '------------------------------------
    grillaNotaPedido.BorderStyle = flexBorderNone
    grillaNotaPedido.ColWidth(0) = 950
    grillaNotaPedido.ColWidth(1) = 5350
    grillaNotaPedido.TextMatrix(0, 0) = "Cliente:"
    grillaNotaPedido.TextMatrix(1, 0) = "Sucursal:"
    grillaNotaPedido.TextMatrix(2, 0) = "Vendedor:"
    For i = 0 To 2
        grillaNotaPedido.Col = 0
        grillaNotaPedido.row = i
        grillaNotaPedido.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grillaNotaPedido.CellBackColor = &H808080    'GRIS OSCURO
        grillaNotaPedido.CellFontBold = True
    Next
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO EL COMBO DE LISTA DE PRECIOS
    Call CargoComboBox(cboListaPrecio, "LISTA_PRECIO", "LIS_CODIGO", "LIS_DESCRI")
    cboListaPrecio.ListIndex = 0
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoRemito) 'ESTADO PENDIENTE
    VEstadoRemito = 1
    FechaRemito.Text = Date
    
    'CARGO COMBO STOCK
    CargaCboStock
    
    'CARGO COMBO REPRESENTADA
    Call CargoComboBox(cboRepresentada, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboRepresentada.ListIndex = -1
    
    'CARGO COMBO CON LAS REPRESENTADAS QUE TIENEN DOCUMENTOS
    CargoComboRepresentadaRemito
    
    'PONGO ENABLE LOS DATOS DE LA FACTURA DE TERCEROS
    txtConcepto.Enabled = False
    TipoBusquedaDoc = 1 'ESTO ES PARA BUSCAR REMITOS(1), (2)PARA BUSCAR NOTA DE PEDIDO
    tabDatos.Tab = 0
    VPedidoPendiente = 0
    'significa que no estoy buacando
    VBanderaBuscar = False
End Sub

Private Function SumaBultos() As Integer
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 2) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 2))
        End If
    Next
    SumaBultos = VTotal
End Function

Private Function SumaPeso() As String
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 6) <> "" Then
            VTotal = VTotal + (CDbl(grdGrilla.TextMatrix(i, 6)) * CDbl(grdGrilla.TextMatrix(i, 2)))
        End If
    Next
    SumaPeso = Format(VTotal, "0.00")
End Function

Private Function SumaDeclarado() As String
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 5) <> "" Then
            VTotal = VTotal + (CDbl(grdGrilla.TextMatrix(i, 5)))
        End If
    Next
    SumaDeclarado = Valido_Importe(CStr(VTotal))
End Function

Private Sub CargoComboRepresentadaRemito()
    sql = "SELECT R.REP_RAZSOC,R.REP_CODIGO FROM REPRESENTADA R"
    'sql = sql & " WHERE R.REP_CODIGO = " & XN(VRepresentada) 'IN (" & XN(VRepresentada) & "," & XN(VRepresentada2) & "," & XN(VRepresentada3) & ")"
    sql = sql & " ORDER BY REP_RAZSOC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRepRemito.AddItem rec!REP_RAZSOC
            cboRepRemito.ItemData(cboRepRemito.NewIndex) = rec!REP_CODIGO
            rec.MoveNext
        Loop
        cboRepRemito.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CargaCboStock()
    sql = "SELECT S.STK_CODIGO,R.REP_RAZSOC"
    sql = sql & " FROM STOCK S, REPRESENTADA R"
    sql = sql & " WHERE S.REP_CODIGO=R.REP_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboStock.AddItem rec!REP_RAZSOC
            cboStock.ItemData(cboStock.NewIndex) = rec!STK_CODIGO
            rec.MoveNext
        Loop
        cboStock.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                grdGrilla.Col = 0
                
                'RESTO LOS BULTOS
                txtBultos.Text = SumaBultos
                txtPeso.Text = SumaPeso
                txtDeclarado.Text = SumaDeclarado
            End If
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 1
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                txtObservaciones.SetFocus
            End If
        Case 2
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" Then
                grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "1"
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or (grdGrilla.Col = 4) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 4 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    MySendKeys Chr(9)
                End If
            Else
                grdGrilla.Col = grdGrilla.Col + 1
            End If
        Else
            If (grdGrilla.Col <> 1) Then
                If KeyAscii > 47 And KeyAscii < 58 Then
                    EDITAR grdGrilla, txtEdit, KeyAscii
                End If
            Else
                EDITAR grdGrilla, txtEdit, KeyAscii
            End If
        End If
    End If
End Sub

Private Sub grdGrilla_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    'If Trim(TxtEdit) = "" Then TxtEdit = "0"
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then Exit Sub
        grdGrilla = txtEdit.Text
        txtEdit.Visible = False
        grdGrilla.row = 1
        grdGrilla.Col = 0
    End If
End Sub

Private Sub GrdModulos_Click()
    If GrdModulos.MouseRow = 0 Then
        GrdModulos.Col = GrdModulos.MouseCol
        GrdModulos.ColSel = GrdModulos.MouseCol
        
        If txtOrden.Text = "A" Then
            GrdModulos.Sort = 2
            txtOrden.Text = "B"
        Else
            GrdModulos.Sort = 1
            txtOrden.Text = "A"
        End If
    End If
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.Rows > 1 Then
        Select Case TipoBusquedaDoc
    
        Case 1 'BUSCA REMITOS
            lblEstado.Caption = "Buscando..."
            Screen.MousePointer = vbHourglass
            Limpiar_Remito
            Set Rec1 = New ADODB.Recordset
            
            txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4)
            If txtNroSucursal.Text = Sucursal Then
                'Call BuscaCodigoProxItemData(CInt(VRepresentada), cboRepRemito)
            'ElseIf txtNroSucursal.Text = Sucursal2 Then
                'Call BuscaCodigoProxItemData(CInt(VRepresentada2), cboRepRemito)
            'ElseIf txtNroSucursal.Text = Sucursal3 Then
                'Call BuscaCodigoProxItemData(CInt(VRepresentada3), cboRepRemito)
            End If
            txtNroRemito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8)
            FechaRemito.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            'CARGO EL ESTADO
            Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5)), lblEstadoRemito)
            VEstadoRemito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 5))
            txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 7)
            
            'REMITO SIN FACTURAS
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) = "S" Then
                chkRemitoSinFactura.Value = Checked
                txtConcepto.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 11)
            End If
            'BUSCO STOCK
            If GrdModulos.TextMatrix(GrdModulos.RowSel, 9) <> "" Then
                Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 9)), cboStock)
            Else
                cboStock.ListIndex = 0
            End If
            txtObservaciones.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 8)
            grillaNotaPedido.TextMatrix(0, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
            grillaNotaPedido.TextMatrix(1, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 3)
            grillaNotaPedido.TextMatrix(2, 1) = GrdModulos.TextMatrix(GrdModulos.RowSel, 4)
            Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 12)), cboRepresentada)
            
        '----BUSCO DETALLE DEL REMITO------------------
            sql = "SELECT DRC.*, P.PTO_DESCRI, TP.TPRE_DESCRI, TP.TPRE_PESO"
            sql = sql & " FROM DETALLE_REMITO_CLIENTE DRC, PRODUCTO P, TIPO_PRESENTACION TP"
            sql = sql & " WHERE DRC.RCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 4))
            sql = sql & " AND DRC.RCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 0), 8))
            sql = sql & " AND DRC.RCL_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 1))
            sql = sql & " AND DRC.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=TP.TPRE_CODIGO"
            sql = sql & " ORDER BY DRC.DRC_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                i = 1
                Do While Rec1.EOF = False
                    CambiaColorAFilaDeGrilla grdGrilla, i, vbBlack
                    grdGrilla.TextMatrix(i, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(i, 1) = Trim(Rec1!PTO_DESCRI) & " - " & Trim(Rec1!TPRE_DESCRI)
                    grdGrilla.TextMatrix(i, 2) = Rec1!DRC_CANTIDAD
                    grdGrilla.TextMatrix(i, 3) = Valido_Importe(Rec1!DRC_PRECIO)
                    grdGrilla.TextMatrix(i, 4) = IIf(IsNull(Rec1!DRC_BONIFICA), "", Format(Rec1!DRC_BONIFICA, "0.00"))
                    VBonificacion = 0
                    If Not IsNull(Rec1!DRC_BONIFICA) Then
                        VBonificacion = ((CDbl(Rec1!DRC_CANTIDAD) * CDbl(Chk0(Rec1!DRC_PRECIO))) * CDbl(Rec1!DRC_BONIFICA)) / 100
                        grdGrilla.TextMatrix(i, 5) = Valido_Importe(CStr((CDbl(Rec1!DRC_CANTIDAD) * CDbl(Chk0(Rec1!DRC_PRECIO))) - VBonificacion))
                    Else
                        grdGrilla.TextMatrix(i, 5) = Valido_Importe(CStr(CDbl(Rec1!DRC_CANTIDAD) * CDbl(Chk0(Rec1!DRC_PRECIO))))
                    End If
                    grdGrilla.TextMatrix(i, 6) = Rec1!TPRE_PESO
                    grdGrilla.TextMatrix(i, 7) = Rec1!DRC_NROITEM
                    grdGrilla.TextMatrix(i, 8) = Valido_Importe(Chk0(Rec1!DRC_COSTO))
                    i = i + 1
                    Rec1.MoveNext
                Loop
                txtBultos.Text = SumaBultos
                txtPeso.Text = SumaPeso
                txtDeclarado.Text = SumaDeclarado
            End If
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            '--------------
            FrameRemito.Enabled = False
            FramePedido.Enabled = False
            '--------------
            'significa que estoy buacando
            VBanderaBuscar = True
            
            tabDatos.Tab = 0
            grdGrilla.SetFocus
            grdGrilla.row = 1
            Rec1.Close
        '----------------------------------------------
        Case 2 'BUSCA NOTA PEDIDO
            txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
            FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
            tabDatos.Tab = 0
            txtNroNotaPedido_LostFocus
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
    If tabDatos.Tab = 1 Then
        cmdGrabar.Enabled = False
        cmdImprimir.Enabled = False
        LimpiarBusqueda
        If Me.Visible = True Then txtCliente.SetFocus
        If TipoBusquedaDoc = 1 Then
            frameBuscar.Caption = "Buscar Remito por..."
        Else
            frameBuscar.Caption = "Buscar Nota de Pedido por..."
        End If
    Else
        cmdImprimir.Enabled = True
        TipoBusquedaDoc = 1
        If VEstadoRemito = 1 Then
            cmdGrabar.Enabled = True
        Else
            cmdGrabar.Enabled = False
        End If
    End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    txtSucursal.Text = ""
    txtDesSuc.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    txtVendedor.Text = ""
    txtDesVen.Text = ""
    GrdModulos.Rows = 1
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(txtCliente)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtDesCli.Text = ""
            txtCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        BuscoCondicionIVA = rec!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    rec.Close
End Function

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 0 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 4 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
        Set Rec1 = New ADODB.Recordset
        CambiaColorAFilaDeGrilla grdGrilla, grdGrilla.RowSel, vbBlack
        Select Case grdGrilla.Col
        Case 0, 1
            If Trim(txtEdit) = "" Then
                txtEdit = ""
                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                grdGrilla.Col = 0
                grdGrilla.SetFocus
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, D.LIS_PRECIO, D.LIS_COSTO, TP.TPRE_DESCRI, TP.TPRE_PESO"
            sql = sql & " FROM PRODUCTO P, DETALLE_LISTA_PRECIO D, TIPO_PRESENTACION TP"
            sql = sql & " WHERE"
            If grdGrilla.Col = 0 Then
                sql = sql & " P.PTO_CODIGO=" & XN(txtEdit)
            Else
                sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
            End If
                sql = sql & " AND D.LIS_CODIGO=" & XN(cboListaPrecio.ItemData(cboListaPrecio.ListIndex))
                sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
                sql = sql & " AND P.TPRE_CODIGO=TP.TPRE_CODIGO"
                sql = sql & " AND P.PTO_ESTADO=1"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                If rec.RecordCount > 1 Then
                    grdGrilla.SetFocus
                    frmBuscar.TipoBusqueda = 2
                    'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                    frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                    frmBuscar.TxtDescriB.Text = txtEdit.Text
                    frmBuscar.Show vbModal
                    grdGrilla.Col = 0
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                    grdGrilla.Col = 1
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                    grdGrilla.Col = 3
                    grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                    grdGrilla.Col = 6
                    grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 6)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                    'COSTO
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(Chk0(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)))
                    grdGrilla.Col = 2
                Else
                    grdGrilla.Col = 0
                    grdGrilla.Text = Trim(rec!PTO_CODIGO)
                    grdGrilla.Col = 1
                    grdGrilla.Text = Trim(rec!PTO_DESCRI) & " - " & Trim(rec!TPRE_DESCRI)
                    grdGrilla.Col = 3
                    grdGrilla.Text = Valido_Importe(Trim(rec!LIS_PRECIO))
                    grdGrilla.Col = 6
                    grdGrilla.Text = Trim(Chk0(rec!TPRE_PESO))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                    'COSTO
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(Trim(Chk0(rec!LIS_COSTO)))
                    grdGrilla.Col = 2
                End If
                    If BuscoRepetetidos(CStr(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), grdGrilla.RowSel) = False Then
                     grdGrilla.Col = 0
                     grdGrilla_KeyDown vbKeyDelete, 0
                    End If
            Else
                    MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                    txtEdit.Text = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
            End If
            rec.Close
            Screen.MousePointer = vbNormal
            
        Case 2  'CANTIDAD
            If Trim(txtEdit) = "" Then grdGrilla.Text = "1"
            grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = txtEdit.Text
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                If Trim(txtEdit) <> "" Then
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 4) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) * CDbl(txtEdit.Text)) / 100)
                    Else
                        VBonificacion = 0
                    End If
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                End If
            End If
            
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                
                'ACA CALCULO Y AGREGO LOS PEDIDOS PENDIENTES A LA GRILLA
                'ADEMAS CALCULO EL DISPONIBLE
                VPedidoPendiente = 0
                sql = "SELECT DISTINCT DNP.PTO_CODIGO, SUM(DNP.DNP_CANTIDAD) AS PEDPEN"
                sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, NOTA_PEDIDO NP"
                sql = sql & " WHERE NP.NPE_NUMERO = DNP.NPE_NUMERO AND NP.EST_CODIGO = 1"
                sql = sql & " AND DNP_MARCA IS NULL"
                sql = sql & " AND DNP.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 0))
                sql = sql & " GROUP BY DNP.PTO_CODIGO"
                sql = sql & " ORDER BY DNP.PTO_CODIGO"
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    VPedidoPendiente = Rec1!PEDPEN
                Else
                    VPedidoPendiente = 0
                End If
                Rec1.Close
                
                'CONTROL DE STOCK FISICO (DEL STOCK SELECCIONADO EN EL MOMENTO)
                sql = "SELECT DS.DST_STKFIS, DS.DST_STKPEN, (DS.DST_STKFIS-DS.DST_STKPEN) AS DISPONIBLE"
                sql = sql & " FROM DETALLE_STOCK DS"
                sql = sql & " WHERE DS.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 0))
                sql = sql & " AND STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
    
                Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
                If Rec1.EOF = False Then
                    If CInt(Chk0(Rec1!DISPONIBLE)) < CInt(txtEdit.Text) Then
                        'BUSCO SI HAY STOCK DEL PRODUCTO EN TODOS LOS STOCKS PARA MOSTRARLE
                        sql = "SELECT SUM(DS.DST_STKFIS) AS FISICO, SUM(DS.DST_STKPEN) AS PENDIENTE"
                        sql = sql & " FROM DETALLE_STOCK DS"
                        sql = sql & " WHERE DS.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 0))
                        
                        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
                        If Rec2.EOF = False And Not IsNull(Rec2!PENDIENTE) Then
                            MsgBox "Producto: " & Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) & Chr(13) & _
                            "Stock: " & Trim(cboStock.List(cboStock.ListIndex)) & Chr(13) & _
                            " La cantidad ingresada supera al Stock Disponible" & Chr(13) & Chr(13) & _
                            " Pedido Pendiente = " & VPedidoPendiente & Chr(13) & _
                            " Stock Pendiente = " & Rec1!DST_STKPEN & Chr(13) & _
                            " Stock Fisico = " & Rec1!DST_STKFIS & Chr(13) & _
                            " Disponible = " & CInt(Chk0(Rec1!DISPONIBLE)) & Chr(13) & Chr(13) & _
                            " Otros Stocks " & Chr(13) & " Stock Pendiente = " & Rec2!PENDIENTE & Chr(13) & _
                            " Stock Fisico = " & Rec2!FISICO & Chr(13) & _
                            " Fisico - Pendiente = " & (CInt(Rec2!FISICO) - CInt(Rec2!PENDIENTE)), vbInformation, TIT_MSGBOX
                            
                        Else
                            MsgBox "Producto: " & Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) & Chr(13) & _
                            "Stock: " & Trim(cboStock.List(cboStock.ListIndex)) & Chr(13) & _
                            " La cantidad ingresada supera al Stock Disponible" & Chr(13) & Chr(13) & _
                            " Pedido Pendiente = " & VPedidoPendiente & Chr(13) & _
                            " Stock Pendiente = " & Rec1!DST_STKPEN & Chr(13) & _
                            " Stock Fisico = " & Rec1!DST_STKFIS & Chr(13) & _
                            " Disponible = " & CInt(Chk0(Rec1!DISPONIBLE)), vbInformation, TIT_MSGBOX
                        End If
                        
                        If cboStock.ItemData(cboStock.ListIndex) = 1 Then
                            If (CInt(Rec2!FISICO) - CInt(Rec2!PENDIENTE)) <= 0 Then
                                If MsgBox("No hay Disponibilidad del producto en Stock,  ¿ Agrega ?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
                                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.RowSel
                                    grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                                    grdGrilla.Col = 0
                                Else
                                    CambiaColorAFilaDeGrilla grdGrilla, grdGrilla.RowSel, vbRed
                                End If
                            Else
                                CambiaColorAFilaDeGrilla grdGrilla, grdGrilla.RowSel, vbRed
                            End If
                        Else
                            MsgBox "El producto no sera agregado al Remito", vbInformation, TIT_MSGBOX
                            LimpiarFilasDeGrilla grdGrilla, grdGrilla.RowSel
                            grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = grdGrilla.RowSel
                            grdGrilla.Col = 0
                        End If
                        Rec2.Close
                    End If
                End If
                Rec1.Close
                
                'CUENTO LOS BULTOS
                txtBultos.Text = SumaBultos
                txtPeso.Text = SumaPeso
                txtDeclarado.Text = SumaDeclarado
            End If
            
        Case 3 'PRECIO
'            If Trim(txtEdit) <> "" Then
'                txtEdit.Text = Valido_Importe(txtEdit)
'                grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = txtEdit.Text
'            Else
'                MsgBox "Debe ingresar el Importe", vbExclamation, TIT_MSGBOX
'                grdGrilla.Col = 3
'            End If
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                If Trim(txtEdit) <> "" Then
                    txtEdit.Text = Valido_Importe(txtEdit.Text)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = txtEdit.Text
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(CInt(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3))))
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 4) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 4))) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                    End If
                Else
                    MsgBox "Debe ingresar el Precio", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 1
                End If
            Else
                txtEdit.Text = ""
            End If
            
        Case 4 'BONIFICACION
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                If Trim(txtEdit) <> "" Then
                    If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                        Exit Sub
                    End If
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Format(txtEdit.Text, "0.00")
                    VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) * CDbl(txtEdit.Text)) / 100)
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                    txtDeclarado.Text = SumaDeclarado
                Else
                    MsgBox "Debe ingresar el Porcentaje", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 4
                End If
            End If
            
        End Select
        grdGrilla.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
       txtEdit.Visible = False
       grdGrilla.SetFocus
    End If
End Sub

Private Function ESTA_EN_STOCK(Producto As String, Stock As String, DESPRO As String) As Boolean
    sql = "SELECT PTO_CODIGO FROM DETALLE_STOCK DS"
    sql = sql & " WHERE DS.STK_CODIGO=" & XN(Stock)
    sql = sql & " AND PTO_CODIGO=" & XN(Producto)
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        ESTA_EN_STOCK = True 'SI ESTA EN EL STOCK
    Else
        If MsgBox("El producto: " & Trim(DESPRO) _
                 & Chr(13) & "no se encuentra en el Stock seleccionado" _
                 & Chr(13) & Chr(13) & "¿Desea agregarlo al Stock ahora?", vbExclamation + vbYesNo, TIT_MSGBOX) = vbYes Then
                 
             'frmDetalleStock.Show vbModal
        End If
        ESTA_EN_STOCK = False 'NO ESTA EN EL STOCK
    End If
    Rec1.Close
End Function

Private Function BuscoRepetetidos(Codigo As String, linea As Integer) As Boolean
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 0) <> "" Then
            If Codigo = CLng(grdGrilla.TextMatrix(i, 0)) And (i <> linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Sub txtNroNotaPedido_GotFocus()
    SelecTexto txtNroNotaPedido
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
       
    If txtNroNotaPedido.Text <> "" And VBanderaBuscar = False Then
        sql = "SELECT NP.*, E.EST_DESCRI"
        sql = sql & " FROM NOTA_PEDIDO NP, ESTADO_DOCUMENTO E"
        sql = sql & " WHERE NP.NPE_NUMERO=" & XN(txtNroNotaPedido)
        If FechaNotaPedido.Text <> "" Then
            sql = sql & " AND NP.NPE_FECHA=" & XDQ(FechaNotaPedido)
        End If
        sql = sql & " AND NP.EST_CODIGO=E.EST_CODIGO"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                MsgBox "Hay mas de una Nota de Pedido con el Número: " & txtNroNotaPedido.Text, vbInformation, TIT_MSGBOX
                Rec2.Close
                cmdBuscarNotaPedido_Click
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Text = Rec2!NPE_FECHA
            grillaNotaPedido.TextMatrix(0, 1) = BuscoCliente(Rec2!CLI_CODIGO)
            grillaNotaPedido.TextMatrix(1, 1) = BuscoSucursal(Rec2!SUC_CODIGO, Rec2!CLI_CODIGO)
            grillaNotaPedido.TextMatrix(2, 1) = BuscoVendedor(Rec2!VEN_CODIGO)
            
            'BUSCO LA REPRESENTADA
            Call BuscaCodigoProxItemData(CInt(Rec2!REP_CODIGO), cboRepresentada)
            
            'BUSCO SI LA NOTA DE PEDIDO TIENE OBSERVACIONES
            txtObservaciones.Text = IIf(IsNull(Rec2!NPE_OBSERVACION), "", Trim(Rec2!NPE_OBSERVACION))
              
            'LE DIGO QUE STOCK USAR
            Select Case Rec2!REP_CODIGO
                Case 5, 8, 10, 11
                    Call BuscaCodigoProxItemData(1, cboStock)
                    
                Case 3
                    Call BuscaCodigoProxItemData(3, cboStock)
                   
                Case 1, 2, 4
                    Call BuscaCodigoProxItemData(2, cboStock)
                    
                Case 7 'VIÑA DE MAIPU
                    Call BuscaCodigoProxItemData(6, cboStock)
                    
                Case 9
                    Call BuscaCodigoProxItemData(4, cboStock)
                   
                Case 14 'PEDRO CARRICONDO E HIJOS S.R.L.
                    Call BuscaCodigoProxItemData(5, cboStock)
            End Select
            
            'BUSCA EL NRO DE REMITO
            'cboRepRemito_LostFocus
            
            'lblEstadoNotaPedido.Caption = "Estado: " & Rec2!EST_DESCRI
            If Rec2!EST_CODIGO <> 1 Then
                MsgBox "La Nota de Pedido número: " & txtNroNotaPedido.Text & Chr(13) & Chr(13) & _
                       "No puede ser asignada al Remito por su estado (" & Rec2!EST_DESCRI & ")", vbExclamation, TIT_MSGBOX
                LimpiarNotaPedido
                cmdGrabar.Enabled = False
                Screen.MousePointer = vbNormal
                lblEstado.Caption = ""
                Rec2.Close
                Exit Sub
            Else
                cmdGrabar.Enabled = True
            End If
            Rec2.Close
            
        '-----BUSCO LOS DATOS DEL DETALLE DE LA NOTA DE PEDIDO---------
            sql = "SELECT DNP.*,P.PTO_DESCRI, TP.TPRE_DESCRI, TP.TPRE_PESO"
            sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, PRODUCTO P, TIPO_PRESENTACION TP"
            sql = sql & " WHERE DNP.NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=TP.TPRE_CODIGO"
            sql = sql & " ORDER BY DNP.DNP_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                i = 1
                Do While Rec1.EOF = False
                    If IsNull(Rec1!DNP_MARCA) Then 'ENTRA CUANDO EL PRODUCTO NO TIENE MARCA
                        'BUSCA LA CANTIDAD DEL PRODUCTO QUE YA FUE USADO EN OTROS REMITOS
                        sql = " SELECT DR.PTO_CODIGO, SUM(DR.DRC_CANTIDAD) AS CANTIDAD"
                        sql = sql & " FROM DETALLE_REMITO_CLIENTE DR, REMITO_CLIENTE R"
                        sql = sql & " WHERE"
                        sql = sql & " R.NPE_NUMERO=" & XN(txtNroNotaPedido)
                        sql = sql & " AND R.NPE_FECHA=" & XDQ(FechaNotaPedido)
                        sql = sql & " AND R.RCL_NUMERO=DR.RCL_NUMERO"
                        sql = sql & " AND R.RCL_SUCURSAL=DR.RCL_SUCURSAL"
                        sql = sql & " AND DR.PTO_CODIGO=" & XN(Rec1!PTO_CODIGO)
                        sql = sql & " AND R.EST_CODIGO <> 2" 'BUSCA EN LOS REMITOS NO ANULADOS
                        sql = sql & " GROUP BY DR.PTO_CODIGO"
                        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                        If rec.EOF = False Then
                            CantidadProducto = rec!Cantidad
                        Else
                            CantidadProducto = 0
                        End If
                        rec.Close
                        grdGrilla.TextMatrix(i, 0) = Rec1!PTO_CODIGO
                        grdGrilla.TextMatrix(i, 1) = Trim(Rec1!PTO_DESCRI) & " - " & Trim(Rec1!TPRE_DESCRI)
                        grdGrilla.TextMatrix(i, 2) = CStr(CInt(Rec1!DNP_CANTIDAD) - CantidadProducto)
                        grdGrilla.TextMatrix(i, 3) = Valido_Importe(Rec1!DNP_PRECIO)
                        grdGrilla.TextMatrix(i, 4) = IIf(IsNull(Rec1!DNP_BONIFICA), "", Format(Rec1!DNP_BONIFICA, "0.00"))
                        VBonificacion = 0
                        If Not IsNull(Rec1!DNP_BONIFICA) Then
                            VBonificacion = ((CDbl(Rec1!DNP_CANTIDAD) * CDbl(Chk0(Rec1!DNP_PRECIO))) * CDbl(Rec1!DNP_BONIFICA)) / 100
                            grdGrilla.TextMatrix(i, 5) = Valido_Importe(CStr((CDbl(Rec1!DNP_CANTIDAD) * CDbl(Chk0(Rec1!DNP_PRECIO))) - VBonificacion))
                        Else
                            grdGrilla.TextMatrix(i, 5) = Valido_Importe(CStr(CDbl(Rec1!DNP_CANTIDAD) * CDbl(Chk0(Rec1!DNP_PRECIO))))
                        End If
                        grdGrilla.TextMatrix(i, 6) = Trim(Chk0(Rec1!TPRE_PESO))
                        grdGrilla.TextMatrix(i, 7) = i 'Rec1!DNP_NROITEM
                        grdGrilla.TextMatrix(i, 8) = Valido_Importe(Chk0(Rec1!DNP_COSTO))
                        
                        CambiaColorAFilaDeGrilla grdGrilla, i, vbBlack
                        'CONTROLO EL STOCK FISICO
                        Call ControlStockFisico(i)
                        
                        i = i + 1
                    End If
                    Rec1.MoveNext
                Loop
                txtBultos.Text = SumaBultos
                txtPeso.Text = SumaPeso
                txtDeclarado.Text = SumaDeclarado
            End If
            Rec1.Close
            '--------------------------------------------------
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            chkRemitoSinFactura.SetFocus
        Else
            MsgBox "La Nota de Pedido no existe", vbExclamation, TIT_MSGBOX
            If Rec2.State = 1 Then Rec2.Close
            LimpiarNotaPedido
        End If
    End If
End Sub

Private Sub ControlStockFisico(Fila As Integer)
    'ACA CALCULO Y AGREGO LOS PEDIDOS PENDIENTES A LA GRILLA
    'ADEMAS CALCULO EL DISPONIBLE
    VPedidoPendiente = 0
    sql = "SELECT DISTINCT DNP.PTO_CODIGO, SUM(DNP.DNP_CANTIDAD) AS PEDPEN"
    sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, NOTA_PEDIDO NP"
    sql = sql & " WHERE NP.NPE_NUMERO = DNP.NPE_NUMERO AND NP.EST_CODIGO = 1"
    sql = sql & " AND DNP_MARCA IS NULL"
    sql = sql & " AND DNP.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(Fila, 0))
    'sql = sql & " AND NP.NPE_NUMERO <> " & XN(txtNroNotaPedido.Text)
    sql = sql & " GROUP BY DNP.PTO_CODIGO"
    sql = sql & " ORDER BY DNP.PTO_CODIGO"
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        VPedidoPendiente = Rec2!PEDPEN
    Else
        VPedidoPendiente = 0
    End If
    Rec2.Close
    
    'BUSCO EN EL STOCK SELECCIONADO
    sql = "SELECT DS.DST_STKFIS, DS.DST_STKPEN, (DS.DST_STKFIS-DS.DST_STKPEN) AS DISPONIBLE"
    sql = sql & " FROM DETALLE_STOCK DS"
    sql = sql & " WHERE DS.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(Fila, 0))
    sql = sql & " AND STK_CODIGO=" & XN(cboStock.ItemData(cboStock.ListIndex))
    
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        'If (CInt(Chk0(Rec2!DISPONIBLE)) - VPedidoPendiente) < CInt(grdGrilla.TextMatrix(Fila, 2)) Then
        If CInt(Chk0(Rec2!DISPONIBLE)) < CInt(grdGrilla.TextMatrix(Fila, 2)) Then
            'BUSCO SI HAY STOCK DEL PRODUCTO EN TODOS LOS STOCKS PARA MOTRALE
            sql = "SELECT SUM(DS.DST_STKFIS) AS FISICO, SUM(DS.DST_STKPEN) AS PENDIENTE"
            sql = sql & " FROM DETALLE_STOCK DS"
            sql = sql & " WHERE DS.PTO_CODIGO=" & XN(grdGrilla.TextMatrix(Fila, 0))
            
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False And Not IsNull(rec!PENDIENTE) Then
                MsgBox "Producto: " & Trim(grdGrilla.TextMatrix(Fila, 1)) & Chr(13) & _
                "Stock: " & Trim(cboStock.List(cboStock.ListIndex)) & Chr(13) & _
                " La cantidad ingresada supera al Stock Disponible" & Chr(13) & Chr(13) & _
                " Pedido Pendiente = " & VPedidoPendiente & Chr(13) & _
                " Stock Pendiente = " & Rec2!DST_STKPEN & Chr(13) & _
                " Stock Fisico = " & Rec2!DST_STKFIS & Chr(13) & _
                " Disponible = " & CInt(Chk0(Rec2!DISPONIBLE)) & Chr(13) & Chr(13) & _
                " Otros Stocks " & Chr(13) & " Stock Pendiente = " & rec!PENDIENTE & Chr(13) & _
                " Stock Fisico = " & rec!FISICO & Chr(13) & _
                " Fisico - Pendiente = " & (CInt(rec!FISICO) - CInt(rec!PENDIENTE)), vbInformation, TIT_MSGBOX
            Else
                MsgBox "Producto: " & Trim(grdGrilla.TextMatrix(Fila, 1)) & Chr(13) & _
                "Stock: " & Trim(cboStock.List(cboStock.ListIndex)) & Chr(13) & _
                " La cantidad ingresada supera al Stock Disponible" & Chr(13) & Chr(13) & _
                " Pedido Pendiente = " & VPedidoPendiente & Chr(13) & _
                " Stock Pendiente = " & Rec2!DST_STKPEN & Chr(13) & _
                " Stock Fisico = " & Rec2!DST_STKFIS & Chr(13) & _
                " Disponible = " & CInt(Chk0(Rec2!DISPONIBLE)), vbInformation, TIT_MSGBOX
            End If
            'rec.Close
            If cboStock.ItemData(cboStock.ListIndex) = 1 Then
                If (CInt(rec!FISICO) - CInt(rec!PENDIENTE)) <= 0 Then
                    If MsgBox("No hay Disponibilidad del producto en Stock,  ¿ Agrega ?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
                        LimpiarFilasDeGrilla grdGrilla, CLng(Fila)
                        grdGrilla.TextMatrix(Fila, 7) = Fila
                        i = i - 1
                    Else
                        CambiaColorAFilaDeGrilla grdGrilla, Fila, vbRed
                    End If
                Else
                    CambiaColorAFilaDeGrilla grdGrilla, Fila, vbRed
                End If
                
            Else
                MsgBox "El producto no sera agregado al Remito", vbInformation, TIT_MSGBOX
                LimpiarFilasDeGrilla grdGrilla, CLng(Fila)
                grdGrilla.TextMatrix(Fila, 7) = Fila
                i = i - 1
            End If
            rec.Close
        End If
    End If
    Rec2.Close
End Sub

Private Sub LimpiarNotaPedido()
    FrameRemito.Enabled = True
    FramePedido.Enabled = True
    txtNroNotaPedido.Text = ""
    FechaNotaPedido.Text = ""
    grillaNotaPedido.TextMatrix(0, 1) = ""
    grillaNotaPedido.TextMatrix(1, 1) = ""
    grillaNotaPedido.TextMatrix(2, 1) = ""
    cboRepresentada.ListIndex = -1
    txtNroNotaPedido.SetFocus
End Sub
Private Function BuscoVendedor(Codigo As String) As String
    sql = "SELECT VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO=" & XN(Codigo)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        BuscoVendedor = Trim(rec!VEN_NOMBRE)
    Else
        BuscoVendedor = "No se encontro el Vendedor"
    End If
    rec.Close
End Function

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT CLI_RAZSOC FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(Codigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            BuscoCliente = rec!CLI_RAZSOC
        Else
            BuscoCliente = "No se encontro el Cliente"
        End If
        rec.Close
End Function

Private Function BuscoSucursal(CodigoSuc As String, CodigoCli As String) As String
        sql = "SELECT * FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(CodigoSuc)
        sql = sql & " AND CLI_CODIGO=" & XN(CodigoCli)
        
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            BuscoSucursal = Rec1!SUC_DESCRI
        Else
            BuscoSucursal = "No se encontro la Sucursal"
        End If
        Rec1.Close
End Function

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtSucursal_Change()
    If txtSucursal.Text = "" Then
        txtDesSuc.Text = ""
    End If
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto txtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If txtSucursal.Text <> "" Then
        Set Rec1 = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, SUC_DESCRI FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtSucursal)
        If txtCliente.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(txtCliente)
        End If
        lblEstado.Caption = "Buscando..."
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            txtCliente.Text = Rec1!CLI_CODIGO
            txtCliente_LostFocus
            txtDesSuc.Text = Rec1!SUC_DESCRI
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDesSuc.Text = ""
            txtSucursal.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Sub txtVendedor_Change()
    If txtVendedor.Text = "" Then
        txtDesVen.Text = ""
    End If
End Sub

Private Sub txtVendedor_GotFocus()
    SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtVendedor_LostFocus()
    If txtVendedor.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT VEN_NOMBRE"
        sql = sql & " FROM VENDEDOR"
        sql = sql & " WHERE VEN_CODIGO=" & XN(txtVendedor)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtDesVen.Text = Trim(rec!VEN_NOMBRE)
        Else
            MsgBox "El Vendedor no existe", vbExclamation, TIT_MSGBOX
            txtDesVen.Text = ""
            txtVendedor.SetFocus
        End If
        rec.Close
    End If
End Sub
