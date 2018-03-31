VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmNotaDePedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Pedido"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8505
      TabIndex        =   13
      Top             =   7665
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7620
      TabIndex        =   12
      Top             =   7665
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10275
      TabIndex        =   15
      Top             =   7665
      Width           =   870
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9390
      TabIndex        =   14
      Top             =   7665
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7620
      Left            =   60
      TabIndex        =   24
      Top             =   15
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13441
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
      TabPicture(0)   =   "frmNotaDePedido.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FramePedido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDatos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaDePedido.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdModulos"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
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
         Height          =   1710
         Left            =   -74655
         TabIndex        =   44
         Top             =   570
         Width           =   10455
         Begin VB.TextBox txtOrden 
            Height          =   315
            Left            =   7590
            TabIndex        =   67
            Text            =   "A"
            Top             =   1320
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.CommandButton cmdBuscarVen 
            Height          =   315
            Left            =   3300
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Buscar Vendedor"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCli 
            Height          =   315
            Left            =   3300
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Buscar Cliente"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarSuc 
            Height          =   315
            Left            =   3300
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Buscar Sucursal"
            Top             =   615
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2265
            TabIndex        =   18
            Top             =   960
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
            Left            =   3750
            TabIndex        =   51
            Top             =   960
            Width           =   4620
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "&Buscar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8580
            MaskColor       =   &H000000FF&
            TabIndex        =   21
            ToolTipText     =   "Buscar "
            Top             =   915
            UseMaskColor    =   -1  'True
            Width           =   1650
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   4770
            TabIndex        =   20
            Top             =   1320
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   2265
            TabIndex        =   19
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
            Left            =   3750
            MaxLength       =   50
            TabIndex        =   46
            Tag             =   "Descripción"
            Top             =   270
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2265
            MaxLength       =   40
            TabIndex        =   16
            Top             =   270
            Width           =   975
         End
         Begin VB.TextBox txtSucursal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2265
            MaxLength       =   40
            TabIndex        =   17
            Top             =   615
            Width           =   975
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
            Left            =   3750
            MaxLength       =   50
            TabIndex        =   45
            Tag             =   "Descripción"
            Top             =   615
            Width           =   4620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1170
            TabIndex        =   52
            Top             =   1005
            Width           =   750
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3720
            TabIndex        =   50
            Top             =   1380
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1170
            TabIndex        =   49
            Top             =   1365
            Width           =   990
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   1170
            TabIndex        =   48
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1170
            TabIndex        =   47
            Top             =   660
            Width           =   660
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   " Datos del Cliente"
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
         Left            =   3555
         TabIndex        =   29
         Top             =   330
         Width           =   7470
         Begin VB.TextBox txtDomiSuc 
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
            Height          =   285
            Left            =   900
            MaxLength       =   50
            TabIndex        =   62
            Top             =   1410
            Width           =   4620
         End
         Begin VB.TextBox txtDescripcionSuc 
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
            Left            =   2805
            MaxLength       =   50
            TabIndex        =   23
            Tag             =   "Descripción"
            Top             =   1065
            Width           =   4590
         End
         Begin VB.TextBox txtCodigoSuc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   900
            MaxLength       =   40
            TabIndex        =   5
            Top             =   1065
            Width           =   975
         End
         Begin VB.CommandButton cmdBuscarSucursal 
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Buscar Sucursal"
            Top             =   1065
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevaSucursal 
            Height          =   315
            Left            =   2355
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0C60
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Agregar Sucursal"
            Top             =   1065
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdNuevoCliente 
            Height          =   315
            Left            =   2355
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":0FEA
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Agregar Cliente"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton cmdBuscarCliente 
            Height          =   315
            Left            =   1920
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":1374
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Buscar Cliente"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox TxtCodigoCli 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   900
            MaxLength       =   40
            TabIndex        =   3
            Top             =   270
            Width           =   975
         End
         Begin VB.TextBox txtRazSocCli 
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
            Left            =   2805
            MaxLength       =   50
            TabIndex        =   4
            Tag             =   "Descripción"
            Top             =   270
            Width           =   4590
         End
         Begin VB.TextBox txtDomici 
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
            Height          =   285
            Left            =   900
            MaxLength       =   50
            TabIndex        =   30
            Top             =   615
            Width           =   4620
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   64
            Top             =   885
            Width           =   750
         End
         Begin VB.Line Line1 
            X1              =   885
            X2              =   7380
            Y1              =   1005
            Y2              =   1005
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   63
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   43
            Top             =   1125
            Width           =   660
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   32
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   31
            Top             =   645
            Width           =   660
         End
      End
      Begin VB.Frame FramePedido 
         Caption         =   "Pedido..."
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
         Left            =   105
         TabIndex        =   26
         Top             =   330
         Width           =   3435
         Begin VB.ComboBox CboVendedor 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1005
            Width           =   2520
         End
         Begin FechaCtl.Fecha FechaNotaPedido 
            Height          =   285
            Left            =   840
            TabIndex        =   1
            Top             =   675
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.TextBox txtNroNotaPedido 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   315
            Width           =   1155
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   57
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lblEstadoNota 
            AutoSize        =   -1  'True
            Caption         =   "EST. NOTA PEDIDO"
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
            Left            =   900
            TabIndex        =   56
            Top             =   1455
            Width           =   1515
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   35
            Top             =   1050
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   28
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   75
            TabIndex        =   27
            Top             =   345
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4590
         Left            =   -74670
         TabIndex        =   22
         Top             =   2430
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   8096
         _Version        =   393216
         Cols            =   4
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
      Begin VB.Frame Frame1 
         Height          =   900
         Left            =   105
         TabIndex        =   58
         Top             =   2025
         Width           =   10920
         Begin VB.ComboBox cboRepresentada 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2445
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   510
            Width           =   4185
         End
         Begin VB.CheckBox chkDetalle 
            Alignment       =   1  'Right Justify
            Caption         =   "NP Detallada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   6
            Top             =   225
            Width           =   1260
         End
         Begin VB.ComboBox cboListaPrecio 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8340
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   165
            Width           =   2505
         End
         Begin VB.ComboBox cboCondicion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2445
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   165
            Width           =   4185
         End
         Begin VB.CommandButton cmdNuevoRubro 
            Height          =   315
            Left            =   6660
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDePedido.frx":167E
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1335
            TabIndex        =   65
            Top             =   585
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Lista de Precios:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7125
            TabIndex        =   61
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1665
            TabIndex        =   60
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4605
         Left            =   105
         TabIndex        =   36
         Top             =   2835
         Width           =   10920
         Begin VB.TextBox txtTotal 
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
            Left            =   9030
            TabIndex        =   70
            Top             =   3840
            Width           =   1170
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
            Left            =   6150
            TabIndex        =   69
            Top             =   3840
            Width           =   915
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1260
            MaxLength       =   60
            TabIndex        =   11
            Top             =   4185
            Width           =   8940
         End
         Begin VB.CommandButton cmdBuscarProducto 
            Height          =   330
            Left            =   10440
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDePedido.frx":1A08
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Producto"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdAgregarProducto 
            Height          =   330
            Left            =   10440
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDePedido.frx":1D12
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Agregar Producto"
            Top             =   540
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   10440
            MaskColor       =   &H8000000F&
            Picture         =   "frmNotaDePedido.frx":201C
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Producto"
            Top             =   885
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   255
            TabIndex        =   37
            Top             =   495
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3600
            Left            =   120
            TabIndex        =   10
            Top             =   165
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   6350
            _Version        =   393216
            Rows            =   13
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   262
            BackColorSel    =   12648447
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            GridColor       =   -2147483633
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8550
            TabIndex        =   72
            Top             =   3885
            Width           =   420
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Bultos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5610
            TabIndex        =   71
            Top             =   3885
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   68
            Top             =   4230
            Width           =   1125
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenado por :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   25
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Pedidos"
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
      Left            =   5235
      TabIndex        =   73
      Top             =   7740
      Width           =   2070
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
      Left            =   225
      TabIndex        =   55
      Top             =   7725
      Width           =   660
   End
End
Attribute VB_Name = "frmNotaDePedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim VBonificacion As Double
Dim VTotal As Double

Private Sub cboRepresentada_LostFocus()
'    grdGrilla.Col = 0
'    grdGrilla.row = 1
End Sub

Private Sub chkDetalle_Click()
    If chkDetalle.Value = Checked Then
        cboListaPrecio.ListIndex = 0
        cboCondicion.ListIndex = 0
        cboCondicion.Enabled = True
        cmdNuevoRubro.Enabled = True
    Else
        cboListaPrecio.ListIndex = 0
        cboCondicion.ListIndex = -1
        cboCondicion.Enabled = False
        cmdNuevoRubro.Enabled = False
    End If
End Sub

Private Sub cmdAgregarProducto_Click()
    'ABMProducto.Show vbModal
    'grdGrilla.SetFocus
    'grdGrilla.row = 1
End Sub

Private Sub cmdBorrar_Click()
    If txtNroNotaPedido.Text <> "" Then
        If MsgBox("Seguro que desea eliminar la Nota de Pedido Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo + vbDefaultButton2, TIT_MSGBOX) = vbYes Then
           On Error GoTo Seclavose
           
           sql = "SELECT P.EST_CODIGO, E.EST_DESCRI "
           sql = sql & " FROM NOTA_PEDIDO P, ESTADO_DOCUMENTO E"
           sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
           sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
           sql = sql & " AND P.EST_CODIGO=E.EST_CODIGO"
           rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
           
           If rec.EOF = False Then
                If rec!EST_CODIGO <> 1 Then
                    MsgBox "La Nota de Pedido no puede ser eliminada," & Chr(13) & _
                           " ya que esta en estado: " & Trim(rec!EST_DESCRI), vbExclamation, TIT_MSGBOX
                    rec.Close
                    Exit Sub
                End If
           End If
           rec.Close
            lblEstado.Caption = "Eliminando..."
            Screen.MousePointer = vbHourglass
            
            sql = "DELETE FROM DETALLE_NOTA_PEDIDO"
            sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            sql = "DELETE FROM NOTA_PEDIDO"
            sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido)
            DBConn.Execute sql
            
            lblEstado.Caption = ""
            Screen.MousePointer = vbNormal
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
Seclavose:
    DBConn.RollbackTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT NP.*, C.CLI_RAZSOC, S.SUC_DESCRI"
    sql = sql & " FROM NOTA_PEDIDO NP, CLIENTE C, SUCURSAL S"
    sql = sql & " WHERE"
    sql = sql & " NP.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND NP.SUC_CODIGO=S.SUC_CODIGO"
    sql = sql & " AND C.CLI_CODIGO=S.CLI_CODIGO"
    If txtCliente.Text <> "" Then sql = sql & " AND NP.CLI_CODIGO=" & XN(txtCliente)
    If txtSucursal.Text <> "" Then sql = sql & " AND NP.SUC_CODIGO=" & XN(txtSucursal)
    If txtVendedor.Text <> "" Then sql = sql & " AND NP.VEN_CODIGO=" & XN(txtVendedor)
    If FechaDesde <> "" Then sql = sql & " AND NP.NPE_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND NP.NPE_FECHA<=" & XDQ(FechaHasta)
    sql = sql & " ORDER BY NP.NPE_NUMERO, NP.NPE_FECHA"
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

Private Sub cmdBuscarCliente_Click()
    frmBuscar.TipoBusqueda = 1
    frmBuscar.TxtDescriB.Text = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        frmBuscar.grdBuscar.Col = 1
        txtRazSocCli.Text = frmBuscar.grdBuscar.Text
        txtCodigoSuc.SetFocus
    Else
        TxtCodigoCli.SetFocus
    End If
End Sub

Private Sub cmdBuscarProducto_Click()
    grdGrilla.SetFocus
    frmBuscar.TipoBusqueda = 2
    frmBuscar.CodListaPrecio = 0
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

Private Sub cmdBuscarSucursal_Click()
    frmBuscar.TipoBusqueda = 3
    frmBuscar.TxtDescriB.Text = ""
    If TxtCodigoCli.Text <> "" Then
        frmBuscar.CodigoCli = TxtCodigoCli.Text
    Else
        frmBuscar.CodigoCli = ""
    End If
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 3
        TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
        TxtCodigoCli_LostFocus
        frmBuscar.grdBuscar.Col = 0
        txtCodigoSuc.Text = frmBuscar.grdBuscar.Text
        txtCodigoSuc_LostFocus
    Else
        txtCodigoSuc.SetFocus
    End If
End Sub

Private Sub cmdBuscarVen_Click()
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
    If ValidarNotaPedido = False Then Exit Sub
    If MsgBox("¿Confirma Pedido?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    On Error GoTo HayErrorNota
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_PEDIDO"
    sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido.Text)
    sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = False Then
        If MsgBox("Seguro que modificar la Nota de Pedido Nro.: " & Trim(txtNroNotaPedido), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE NOTA_PEDIDO"
            sql = sql & " SET CLI_CODIGO=" & XN(TxtCodigoCli.Text)
            sql = sql & " ,SUC_CODIGO=" & XN(txtCodigoSuc.Text)
            sql = sql & " ,VEN_CODIGO=" & XN(CboVendedor.ItemData(CboVendedor.ListIndex))
            If chkDetalle.Value = Checked Then
                sql = sql & " ,FPG_CODIGO=" & XN(cboCondicion.ItemData(cboCondicion.ListIndex))
            Else
                sql = sql & " ,FPG_CODIGO=NULL"
            End If
            sql = sql & " ,REP_CODIGO=" & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex))
            sql = sql & " ,NPE_OBSERVACION=" & XS(txtObservaciones.Text)
            sql = sql & " WHERE"
            sql = sql & " NPE_NUMERO=" & XN(txtNroNotaPedido.Text)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido.Text)
            DBConn.Execute sql
            
            sql = "DELETE FROM DETALLE_NOTA_PEDIDO"
            sql = sql & " WHERE NPE_NUMERO=" & XN(txtNroNotaPedido.Text)
            sql = sql & " AND NPE_FECHA=" & XDQ(FechaNotaPedido.Text)
            DBConn.Execute sql
            
            For i = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(i, 0) <> "" Then
                    sql = "INSERT INTO DETALLE_NOTA_PEDIDO"
                    sql = sql & " (NPE_NUMERO,NPE_FECHA,DNP_NROITEM,PTO_CODIGO,"
                    sql = sql & "DNP_CANTIDAD,DNP_PRECIO,DNP_BONIFICA,DNP_COSTO)"
                    sql = sql & " VALUES ("
                    sql = sql & XN(txtNroNotaPedido.Text) & ","
                    sql = sql & XDQ(FechaNotaPedido.Text) & ","
                    sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & "," 'NRO ITEM
                    sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & "," 'PRODUCTO CODIGO
                    sql = sql & XN(grdGrilla.TextMatrix(i, 2)) & "," 'CANTIDAD
                    sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & "," 'PRECIO
                    sql = sql & XN(grdGrilla.TextMatrix(i, 4)) & "," 'BONIFICACION
                    sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ")" 'COSTO
                    DBConn.Execute sql
                End If
            Next
            DBConn.CommitTrans
        End If
        
    Else 'PEDIDO NUEVO
        sql = "INSERT INTO NOTA_PEDIDO"
        sql = sql & " (NPE_NUMERO,NPE_FECHA,CLI_CODIGO,"
        sql = sql & "SUC_CODIGO,VEN_CODIGO,FPG_CODIGO,REP_CODIGO,NPE_NUMEROTXT,"
        sql = sql & "NPE_OBSERVACION,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroNotaPedido.Text) & ","
        sql = sql & XDQ(FechaNotaPedido.Text) & ","
        sql = sql & XN(TxtCodigoCli.Text) & ","
        sql = sql & XN(txtCodigoSuc.Text) & ","
        sql = sql & XN(CboVendedor.ItemData(CboVendedor.ListIndex)) & ","
        If chkDetalle.Value = Checked Then
            sql = sql & XN(cboCondicion.ItemData(cboCondicion.ListIndex)) & ","
        Else
            sql = sql & "NULL,"
        End If
        sql = sql & XN(cboRepresentada.ItemData(cboRepresentada.ListIndex)) & ","
        sql = sql & XS(Format(txtNroNotaPedido.Text, "00000000")) & ","
        sql = sql & XS(txtObservaciones.Text) & ","
        sql = sql & "1)" 'ESTADO PENDIENTE
        DBConn.Execute sql
           
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_PEDIDO"
                sql = sql & " (NPE_NUMERO,NPE_FECHA,DNP_NROITEM,PTO_CODIGO,"
                sql = sql & " DNP_CANTIDAD,DNP_PRECIO,DNP_BONIFICA,DNP_COSTO)"
                sql = sql & " VALUES ("
                sql = sql & XN(txtNroNotaPedido.Text) & ","
                sql = sql & XDQ(FechaNotaPedido.Text) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & "," 'NRO ITEM
                sql = sql & XN(grdGrilla.TextMatrix(i, 0)) & "," 'PRODUCTO CODIGO
                sql = sql & XN(grdGrilla.TextMatrix(i, 2)) & "," 'CANTIDAD
                sql = sql & XN(grdGrilla.TextMatrix(i, 3)) & "," 'PRECIO
                sql = sql & XN(grdGrilla.TextMatrix(i, 4)) & "," 'BONIFOCACION
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ")" 'COSTO
                DBConn.Execute sql
            End If
        Next
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    CmdNuevo_Click
    Exit Sub
    
HayErrorNota:
    If rec.State = 1 Then rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub
Private Function ValidarNotaPedido() As Boolean
    
    If txtNroNotaPedido.Text = "" Then
        MsgBox "El número de Nota de Pedido es requerido", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If FechaNotaPedido.Text = "" Then
        MsgBox "La Fecha de la Nota de pedido es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaPedido.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If TxtCodigoCli.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    If txtCodigoSuc.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
        ValidarNotaPedido = False
        Exit Function
    End If
    ValidarNotaPedido = True
End Function

Private Sub cmdNuevaSucursal_Click()
'    ABMSucursal.Show vbModal
'    txtCodigoSuc.SetFocus
End Sub

Private Sub CmdNuevo_Click()
   For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 0) = ""
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.TextMatrix(i, 5) = ""
        grdGrilla.TextMatrix(i, 6) = i
        grdGrilla.TextMatrix(i, 7) = ""
   Next
   FramePedido.Enabled = True
   fraDatos.Enabled = True
   TxtCodigoCli.Text = ""
   txtCodigoSuc.Text = ""
   chkDetalle.Value = Unchecked
   TxtCodigoCli.Text = ""
   txtRazSocCli.Text = ""
   CboVendedor.ListIndex = 0
   FechaNotaPedido.Text = ""
   txtNroNotaPedido.Text = ""
   lblEstadoNota.Caption = ""
   txtObservaciones.Text = ""
   txtTotal.Text = ""
   txtBultos.Text = ""
   lblEstado.Caption = ""
   tabDatos.Tab = 0
   Call BuscoEstado(1, lblEstadoNota)
   cmdGrabar.Enabled = True
   cmdBorrar.Enabled = True
   txtNroNotaPedido.SetFocus
End Sub

Private Sub cmdNuevoCliente_Click()
'    ABMCliente.Show vbModal
'    TxtCodigoCli.SetFocus
End Sub

Private Sub cmdNuevoRubro_Click()
     ABMFormaPago.Show vbModal
     cboCondicion.Clear
     Call CargoComboBox(cboCondicion, "FORMA_PAGO", "FPG_CODIGO", "FPG_DESCRI")
     cboCondicion.ListIndex = 0
     cboCondicion.SetFocus
End Sub

Private Sub cmdQuitarProducto_Click()
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
        If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
            grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = grdGrilla.RowSel
            grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
            grdGrilla.SetFocus
            grdGrilla.Col = 0
        End If
    Else
        MsgBox "Debe seleccionar un Producto", vbExclamation, TIT_MSGBOX
        grdGrilla.SetFocus
        grdGrilla.Col = 0
        grdGrilla.row = 1
    End If
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaDePedido = Nothing
        Unload Me
    End If
End Sub

Private Sub FechaNotaPedido_LostFocus()
    If FechaNotaPedido.Text = "" Then FechaNotaPedido.Text = Date
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then tabDatos.Tab = 1
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
    grdGrilla.FormatString = "Código|Descripción|>Cantidad|>Precio|>Bonif.|>Importe|Orden|COSTO"
    grdGrilla.ColWidth(0) = 900  'CODIGO
    grdGrilla.ColWidth(1) = 5150 'DESCRIPCION
    grdGrilla.ColWidth(2) = 900  'CANTIDAD
    grdGrilla.ColWidth(3) = 1100 'PRECIO
    grdGrilla.ColWidth(4) = 900  'BONOFICACION
    grdGrilla.ColWidth(5) = 1100 'IMPORTE
    grdGrilla.ColWidth(6) = 0    'ORDEN
    grdGrilla.ColWidth(7) = 0    'COSTO
    grdGrilla.Cols = 8
    grdGrilla.Rows = 1
    grdGrilla.BorderStyle = flexBorderNone
    
    grdGrilla.row = 0
    For i = 0 To 7
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    For i = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                          & Chr(9) & "" & Chr(9) & "" & Chr(9) & (i - 1) & Chr(9) & ""
    Next
    
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = ">Número|^Fecha|Cliente|Sucursal"
    GrdModulos.ColWidth(0) = 1200
    GrdModulos.ColWidth(1) = 1000
    GrdModulos.ColWidth(2) = 4000
    GrdModulos.ColWidth(3) = 4000
    GrdModulos.Cols = 4
    GrdModulos.Rows = 1
    For i = 0 To 3
        GrdModulos.Col = i
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
    Next
    GrdModulos.BorderStyle = flexBorderNone
    '------------------------------------
    'CARGO COMBO LISTA DE PRECIOS
    Call CargoComboBox(cboListaPrecio, "LISTA_PRECIO", "LIS_CODIGO", "LIS_DESCRI")
    cboListaPrecio.ListIndex = 0
    
    'CARGO CONDICIONES DE PAGO
    Call CargoComboBox(cboCondicion, "FORMA_PAGO", "FPG_CODIGO", "FPG_DESCRI")
    
    'CARGO COMBO REPRESENTADA
    Call CargoComboBox(cboRepresentada, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboRepresentada.ListIndex = 0
    
    'CARGO COMBO VENDEDOR
    Call CargoComboBox(CboVendedor, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE")
    CboVendedor.ListIndex = 0
    '----------------
    cboCondicion.ListIndex = -1
    cboCondicion.Enabled = False
    cmdNuevoRubro.Enabled = False
    lblEstado.Caption = ""
    Call BuscoEstado(1, lblEstadoNota)
    tabDatos.Tab = 0
End Sub

Private Function SumaBonificacion() As Double
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 5) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 5))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBultos() As Integer
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 2) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 2))
        End If
    Next
    SumaBultos = VTotal
End Function

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1
            If MsgBox("Seguro que desea quitar el Producto: " & grdGrilla.TextMatrix(grdGrilla.RowSel, 1), vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                grdGrilla.Col = 0
            End If
        'Case Else
        '    grdGrilla.TextArray(GRIDINDEX(grdGrilla, grdGrilla.row, grdGrilla.Col)) = ""
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
        CmdNuevo_Click
        txtNroNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
        FechaNotaPedido.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 1)
        tabDatos.Tab = 0
        txtNroNotaPedido_LostFocus
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
'    txtCliente.Enabled = False
'    txtSucursal.Enabled = False
'    FechaDesde.Enabled = False
'    FechaHasta.Enabled = False
'    txtVendedor.Enabled = False
    cmdGrabar.Enabled = False
    cmdBorrar.Enabled = False
'    cmdBuscarCli.Enabled = False
'    cmdBuscarSuc.Enabled = False
'    cmdBuscarVen.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then txtCliente.SetFocus
  Else
    If Me.Visible = True Then txtNroNotaPedido.SetFocus
    cmdGrabar.Enabled = True
    cmdBorrar.Enabled = True
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
            Exit Sub
        End If
        rec.Close
    End If
'    If chkSucursal.Value = Unchecked And chkFecha.Value = Unchecked _
'        And chkVendedor.Value = Unchecked And ActiveControl.Name <> "cmdBuscarCli" _
'        And ActiveControl.Name <> "cmdNuevo" And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub

Private Sub TxtCodigoCli_Change()
    If TxtCodigoCli.Text = "" Then
        TxtCodigoCli.Text = ""
        txtRazSocCli.Text = ""
        txtDomici.Text = ""
        txtCodigoSuc.Text = ""
        txtDescripcionSuc.Text = ""
        txtDomiSuc.Text = ""
    End If
End Sub

Private Sub txtCodigoCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodigoSuc_Change()
    If txtCodigoSuc.Text = "" Then
        txtCodigoSuc.Text = ""
        txtDescripcionSuc.Text = ""
        txtDomiSuc.Text = ""
    End If
End Sub

Private Sub txtCodigoSuc_GotFocus()
    SelecTexto txtCodigoSuc
End Sub

Private Sub txtCodigoSuc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodigoSuc_LostFocus()
    If txtCodigoSuc.Text <> "" Then
        sql = "SELECT SUC_CODIGO,CLI_CODIGO,SUC_DESCRI,SUC_DOMICI"
        sql = sql & " FROM SUCURSAL"
        sql = sql & " WHERE SUC_CODIGO=" & XN(txtCodigoSuc)
        If TxtCodigoCli.Text <> "" Then
         sql = sql & " AND CLI_CODIGO=" & XN(TxtCodigoCli)
        End If
        lblEstado.Caption = "Buscando..."
        Set Rec1 = New ADODB.Recordset
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
                If Rec1.RecordCount > 1 Then
                    frmBuscar.TipoBusqueda = 3
                    frmBuscar.CodigoCli = ""
                    frmBuscar.TxtDescriB = txtCodigoSuc
                    frmBuscar.Show vbModal
                    frmBuscar.grdBuscar.Col = 3
                    TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
                    TxtCodigoCli_LostFocus
                    frmBuscar.grdBuscar.Col = 0
                    txtCodigoSuc.Text = frmBuscar.grdBuscar.Text
                    txtCodigoSuc_LostFocus
                    Rec1.Close
                    lblEstado.Caption = ""
                    Exit Sub
                End If
            TxtCodigoCli.Text = Rec1!CLI_CODIGO
            TxtCodigoCli_LostFocus
            txtDescripcionSuc.Text = Rec1!SUC_DESCRI
            txtDomiSuc.Text = Rec1!SUC_DOMICI
            lblEstado.Caption = ""
        Else
            lblEstado.Caption = ""
            MsgBox "La Sucursal no existe", vbExclamation, TIT_MSGBOX
            txtDescripcionSuc.Text = ""
            txtCodigoSuc.SetFocus
        End If
        Rec1.Close
    End If
End Sub

Private Function BuscoCondicionIVA(IVACodigo As String) As String
    Set Rec3 = New ADODB.Recordset
    sql = "SELECT * FROM CONDICION_IVA"
    sql = sql & " WHERE IVA_CODIGO=" & XN(IVACodigo)
    Rec3.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec3.EOF = False Then
        BuscoCondicionIVA = Rec3!IVA_DESCRI
    Else
        BuscoCondicionIVA = ""
    End If
    Rec3.Close
    Set Rec3 = Nothing
End Function
Private Sub TxtCodigoCli_GotFocus()
    SelecTexto TxtCodigoCli
End Sub

Private Sub TxtCodigoCli_LostFocus()
    If ActiveControl.Name = "cmdGrabar" Or ActiveControl.Name = "cmdBorrar" Or ActiveControl.Name = "CmdSalir" Then Exit Sub
    If TxtCodigoCli.Text <> "" Then
        sql = "SELECT CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,IVA_CODIGO,CLI_INGBRU"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE CLI_CODIGO=" & XN(TxtCodigoCli)
        'sql = sql & " AND CLI_ESTADO=1"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSocCli.Text = rec!CLI_RAZSOC
            txtDomici.Text = rec!CLI_DOMICI
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        rec.Close
    End If
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
            sql = "SELECT P.PTO_CODIGO, P.PTO_DESCRI, D.LIS_PRECIO, TP.TPRE_DESCRI, D.LIS_COSTO"
            sql = sql & " FROM PRODUCTO P, DETALLE_LISTA_PRECIO D,"
            sql = sql & " TIPO_PRESENTACION TP"
            sql = sql & " WHERE"
            If grdGrilla.Col = 0 Then
                sql = sql & " P.PTO_CODIGO=" & XN(txtEdit)
            Else
                sql = sql & " P.PTO_DESCRI LIKE '" & Trim(txtEdit) & "%'"
            End If
                sql = sql & " AND D.LIS_CODIGO=" & XN(cboListaPrecio.ItemData(cboListaPrecio.ListIndex))
                sql = sql & " AND P.PTO_CODIGO=D.PTO_CODIGO"
                sql = sql & " AND P.TPRE_CODIGO=TP.TPRE_CODIGO"
                sql = sql & " AND PTO_ESTADO=1"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                If rec.RecordCount > 1 Then
                    grdGrilla.SetFocus
                    frmBuscar.TipoBusqueda = 2
                    'LE DIGO EN QUE LISTA DE PRECIO BUSCAR LOS PRECIOS
                    frmBuscar.CodListaPrecio = cboListaPrecio.ItemData(cboListaPrecio.ListIndex)
                    frmBuscar.TxtDescriB.Text = txtEdit.Text
                    frmBuscar.Show vbModal
                    If frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0) <> "" Then
                        grdGrilla.Col = 0
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 0)
                        grdGrilla.Col = 1
                        grdGrilla.Text = frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 1)
                        grdGrilla.Col = 3
                        grdGrilla.Text = Valido_Importe(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 2))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = Valido_Importe(Chk0(frmBuscar.grdBuscar.TextMatrix(frmBuscar.grdBuscar.RowSel, 8)))
                    Else
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
                    End If
                    grdGrilla.Col = 2
                Else
                    grdGrilla.Col = 0
                    grdGrilla.Text = Trim(rec!PTO_CODIGO)
                    grdGrilla.Col = 1
                    grdGrilla.Text = (Trim(rec!PTO_DESCRI) & " - " & Trim(rec!TPRE_DESCRI))
                    grdGrilla.Col = 3
                    grdGrilla.Text = Valido_Importe(Trim(rec!LIS_PRECIO))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = Valido_Importe(Trim(Chk0(rec!LIS_COSTO)))
                    grdGrilla.Col = 2
                End If
                    CambiaColorAFilaDeGrilla grdGrilla, grdGrilla.RowSel, vbBlack
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                        If BuscoRepetetidos(CStr(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), grdGrilla.RowSel) = False Then
                         grdGrilla.Col = 0
                         grdGrilla_KeyDown vbKeyDelete, 0
                        End If
                    End If
            Else
                    MsgBox "No se ha encontrado el Producto", vbExclamation, TIT_MSGBOX
                    txtEdit.Text = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
            End If
            rec.Close
            Screen.MousePointer = vbNormal
            
        Case 2 'CANTIDAD
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
                    txtBultos = SumaBultos
                    txtTotal.Text = Valido_Importe(SumaBonificacion)
                End If
            End If
        
        Case 3 'PRECIO
        
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
                    txtTotal.Text = Valido_Importe(SumaBonificacion)
                Else
                    MsgBox "Debe ingresar el Precio", vbExclamation, TIT_MSGBOX
                    grdGrilla.Col = 1
                End If
            Else
                txtEdit.Text = ""
            End If
        
        Case 4
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                If Trim(txtEdit) <> "" Then
                    If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                        Exit Sub
                    End If
                    VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) * CDbl(txtEdit.Text)) / 100)
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 3)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                    txtTotal.Text = Valido_Importe(SumaBonificacion)
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

Private Sub txtNroNotaPedido_Change()
    If txtNroNotaPedido.Text = "" Then
        FechaNotaPedido.Text = ""
    End If
End Sub

Private Sub txtNroNotaPedido_GotFocus()
     FechaNotaPedido.Text = ""
End Sub

Private Sub txtNroNotaPedido_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaPedido_LostFocus()
    If ActiveControl.Name = "CmdSalir" Or ActiveControl.Name = "cmdNuevo" _
       Or ActiveControl.Name = "cmdBuscarNotaPedido" Or ActiveControl.Name = "txtCliente" Then Exit Sub
    
    If txtNroNotaPedido.Text <> "" Then
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
                tabDatos.Tab = 1
                Rec2.Close
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Buscando..."
            
            'CARGO CABECERA DE LA NOTA DE PEDIDO
            FechaNotaPedido.Text = Rec2!NPE_FECHA
            Call BuscaCodigoProxItemData(Rec2!VEN_CODIGO, CboVendedor)
            'BUSCA FORMA DE PAGO
            If Not IsNull(Rec2!FPG_CODIGO) Then
                chkDetalle.Value = Checked
                Call BuscaCodigoProxItemData(Rec2!FPG_CODIGO, cboCondicion)
            Else
                chkDetalle.Value = Unchecked
            End If
            'BUSCA REPRESENTADA
            Call BuscaCodigoProxItemData(Rec2!REP_CODIGO, cboRepresentada)
            txtObservaciones.Text = IIf(IsNull(Rec2!NPE_OBSERVACION), "", Rec2!NPE_OBSERVACION)
            
            TxtCodigoCli.Text = Rec2!CLI_CODIGO
            TxtCodigoCli_LostFocus
            txtCodigoSuc.Text = Rec2!SUC_CODIGO
            txtCodigoSuc_LostFocus
            Call BuscoEstado(Rec2!EST_CODIGO, lblEstadoNota)
            If Rec2!EST_CODIGO <> 1 Then
                cmdGrabar.Enabled = False
                cmdBorrar.Enabled = False
                FramePedido.Enabled = False
                fraDatos.Enabled = False
                grdGrilla.SetFocus
            Else
                cmdGrabar.Enabled = True
                cmdBorrar.Enabled = True
                FramePedido.Enabled = True
                fraDatos.Enabled = True
            End If
            
            'BUSCO LOS DATOS DEL DETALLE DE LA NOTA DE PEDIDO
            sql = "SELECT DNP.*,P.PTO_DESCRI, TP.TPRE_DESCRI"
            sql = sql & " FROM DETALLE_NOTA_PEDIDO DNP, PRODUCTO P,"
            sql = sql & " TIPO_PRESENTACION TP"
            sql = sql & " WHERE DNP.NPE_NUMERO=" & XN(txtNroNotaPedido)
            sql = sql & " AND DNP.NPE_FECHA=" & XDQ(FechaNotaPedido)
            sql = sql & " AND DNP.PTO_CODIGO=P.PTO_CODIGO"
            sql = sql & " AND P.TPRE_CODIGO=TP.TPRE_CODIGO"
            sql = sql & " ORDER BY DNP.DNP_NROITEM"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
                i = 1
                Do While Rec1.EOF = False
                    grdGrilla.TextMatrix(i, 0) = Rec1!PTO_CODIGO
                    grdGrilla.TextMatrix(i, 1) = Trim(Rec1!PTO_DESCRI) & " - " & Trim(Rec1!TPRE_DESCRI)
                    grdGrilla.TextMatrix(i, 2) = Rec1!DNP_CANTIDAD
                    If IsNull(Rec1!DNP_PRECIO) Then
                        grdGrilla.TextMatrix(i, 3) = ""
                    Else
                        grdGrilla.TextMatrix(i, 3) = Valido_Importe(Rec1!DNP_PRECIO)
                    End If
                    grdGrilla.TextMatrix(i, 4) = IIf(IsNull(Rec1!DNP_BONIFICA), "", Format(Rec1!DNP_BONIFICA, "0.00"))
                    
                    If Not IsNull(Rec1!DNP_BONIFICA) Then
                        VBonificacion = ((CDbl(Rec1!DNP_CANTIDAD) * CDbl(Chk0(Rec1!DNP_PRECIO))) * CDbl(Rec1!DNP_BONIFICA)) / 100
                        grdGrilla.TextMatrix(i, 5) = Valido_Importe(CStr((CDbl(Rec1!DNP_CANTIDAD) * CDbl(Chk0(Rec1!DNP_PRECIO))) - VBonificacion))
                    Else
                        grdGrilla.TextMatrix(i, 5) = Valido_Importe(CStr(CDbl(Rec1!DNP_CANTIDAD) * CDbl(Chk0(Rec1!DNP_PRECIO))))
                    End If
                    grdGrilla.TextMatrix(i, 6) = Rec1!DNP_NROITEM
                    grdGrilla.TextMatrix(i, 7) = Valido_Importe(Chk0(Rec1!DNP_COSTO))
                    If IsNull(Rec1!DNP_MARCA) Then
                        Call CambiaColorAFilaDeGrilla(grdGrilla, i, vbRed)
                    Else
                        Call CambiaColorAFilaDeGrilla(grdGrilla, i, vbBlack)
                    End If
                    i = i + 1
                    Rec1.MoveNext
                Loop
                txtTotal.Text = Valido_Importe(SumaBonificacion)
                txtBultos.Text = SumaBultos
            End If
            Rec1.Close
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
        Else
            Call BuscoEstado(1, lblEstadoNota)
        End If
        Rec2.Close
    Else
        MsgBox "Debe ingresar el Número de Nota de Pedido", vbExclamation, TIT_MSGBOX
        txtNroNotaPedido.SetFocus
    End If
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRazSocCli_Change()
    If txtRazSocCli.Text = "" Then
        TxtCodigoCli.Text = ""
        txtRazSocCli.Text = ""
        txtDomici.Text = ""
        txtCodigoSuc.Text = ""
        txtDescripcionSuc.Text = ""
        txtDomiSuc.Text = ""
    End If
End Sub

Private Sub txtRazSocCli_GotFocus()
    SelecTexto txtRazSocCli
End Sub

Private Sub txtRazSocCli_LostFocus()
    If TxtCodigoCli.Text = "" And txtRazSocCli.Text <> "" Then
        sql = "SELECT CLI_CODIGO,CLI_RAZSOC,CLI_DOMICI,CLI_CUIT,IVA_CODIGO,CLI_INGBRU"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE CLI_RAZSOC LIKE '" & Trim(txtRazSocCli.Text) & "%'"
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            If Rec2.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.TxtDescriB.Text = txtRazSocCli.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    TxtCodigoCli.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtRazSocCli.Text = frmBuscar.grdBuscar.Text
                    TxtCodigoCli_LostFocus
                    txtCodigoSuc.SetFocus
                Else
                    txtRazSocCli.Text = ""
                    TxtCodigoCli.SetFocus
                End If
            Else
                TxtCodigoCli.Text = Rec2!CLI_CODIGO
                txtRazSocCli.Text = Rec2!CLI_RAZSOC
                TxtCodigoCli_LostFocus
                txtCodigoSuc.SetFocus
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtRazSocCli.Text = ""
            TxtCodigoCli.SetFocus
        End If
        Rec2.Close
    ElseIf TxtCodigoCli.Text = "" And txtRazSocCli.Text = "" Then
        MsgBox "Debe elegir un cliente", vbExclamation, TIT_MSGBOX
        TxtCodigoCli.SetFocus
    End If
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
            Exit Sub
        End If
        Rec1.Close
    End If
'    If chkFecha.Value = Unchecked And chkVendedor.Value = Unchecked _
'        And ActiveControl.Name <> "cmdBuscarSuc" And ActiveControl.Name <> "cmdNuevo" _
'        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
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
            Exit Sub
        End If
        rec.Close
    End If
'    If chkFecha.Value = Unchecked And ActiveControl.Name <> "cmdNuevo" _
'        And ActiveControl.Name <> "CmdSalir" Then CmdBuscAprox.SetFocus
End Sub
