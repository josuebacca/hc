VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmFacturaClienteConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura de Clientes (Concepto)..."
   ClientHeight    =   7950
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8535
      TabIndex        =   16
      Top             =   7470
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10305
      TabIndex        =   18
      Top             =   7470
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7650
      TabIndex        =   15
      Top             =   7470
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9420
      TabIndex        =   17
      Top             =   7470
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7425
      Left            =   60
      TabIndex        =   26
      Top             =   15
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13097
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
      TabPicture(0)   =   "frmFacturaClienteConcepto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameFactura"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmFacturaClienteConcepto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   4470
         TabIndex        =   62
         Top             =   315
         Width           =   6540
         Begin VB.TextBox txtCodCliente 
            Height          =   330
            Left            =   750
            MaxLength       =   40
            TabIndex        =   5
            Top             =   420
            Width           =   960
         End
         Begin VB.TextBox txtCliRazSoc 
            Height          =   330
            Left            =   2190
            MaxLength       =   50
            TabIndex        =   6
            Tag             =   "Descripción"
            Top             =   420
            Width           =   4245
         End
         Begin VB.TextBox txtProvincia 
            BackColor       =   &H00C0C0C0&
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
            Height          =   300
            Left            =   750
            TabIndex        =   66
            Top             =   780
            Width           =   5685
         End
         Begin VB.TextBox txtCliLocalidad 
            BackColor       =   &H00C0C0C0&
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
            Left            =   750
            TabIndex        =   65
            Top             =   1110
            Width           =   5685
         End
         Begin VB.TextBox txtDomici 
            BackColor       =   &H00C0C0C0&
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
            Left            =   750
            MaxLength       =   50
            TabIndex        =   64
            Top             =   1425
            Width           =   5685
         End
         Begin VB.CommandButton cmdBuscarCliente 
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
            Left            =   1755
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaClienteConcepto.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Buscar Cliente"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   405
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
            Left            =   150
            TabIndex        =   70
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Prov.:"
            Height          =   195
            Left            =   150
            TabIndex        =   69
            Top             =   810
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            Height          =   195
            Left            =   150
            TabIndex        =   68
            Top             =   1155
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Dom.:"
            Height          =   195
            Left            =   150
            TabIndex        =   67
            Top             =   1455
            Width           =   435
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar por..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   -74640
         TabIndex        =   31
         Top             =   540
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
            Left            =   6480
            TabIndex        =   73
            Text            =   "A"
            Top             =   615
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.ComboBox cboBuscaRep 
            Height          =   315
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1245
            Width           =   3090
         End
         Begin VB.ComboBox cboFactura1 
            Height          =   315
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   900
            Width           =   2400
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
            Left            =   3750
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaClienteConcepto.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Buscar Cliente"
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   360
            Left            =   7140
            MaskColor       =   &H000000FF&
            TabIndex        =   24
            ToolTipText     =   "Buscar "
            Top             =   1200
            UseMaskColor    =   -1  'True
            Width           =   1665
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5205
            TabIndex        =   21
            Top             =   585
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   2700
            TabIndex        =   20
            Top             =   585
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
            Height          =   300
            Left            =   4185
            MaxLength       =   50
            TabIndex        =   32
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   2700
            MaxLength       =   40
            TabIndex        =   19
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   1560
            TabIndex        =   61
            Top             =   1275
            Width           =   1080
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Factura:"
            Height          =   195
            Left            =   1560
            TabIndex        =   58
            Top             =   945
            Width           =   960
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4155
            TabIndex        =   35
            Top             =   630
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1560
            TabIndex        =   34
            Top             =   615
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
            Left            =   1560
            TabIndex        =   33
            Top             =   315
            Width           =   555
         End
      End
      Begin VB.Frame FrameFactura 
         Caption         =   "Factura..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   105
         TabIndex        =   28
         Top             =   315
         Width           =   4350
         Begin VB.ComboBox cboRep 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   255
            Width           =   3090
         End
         Begin VB.TextBox txtNroSucursal 
            Height          =   330
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   2
            Top             =   945
            Width           =   555
         End
         Begin VB.ComboBox cboFactura 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   600
            Width           =   2190
         End
         Begin VB.TextBox txtNroFactura 
            Height          =   330
            Left            =   1770
            MaxLength       =   8
            TabIndex        =   3
            Top             =   945
            Width           =   1065
         End
         Begin FechaCtl.Fecha FechaFactura 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   1305
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   90
            TabIndex        =   60
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   90
            TabIndex        =   44
            Top             =   615
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   90
            TabIndex        =   41
            Top             =   1335
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   90
            TabIndex        =   40
            Top             =   975
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   90
            TabIndex        =   39
            Top             =   1650
            Width           =   555
         End
         Begin VB.Label lblEstadoFactura 
            AutoSize        =   -1  'True
            Caption         =   "EST. FACTURA"
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
            Left            =   1200
            TabIndex        =   38
            Top             =   1665
            Width           =   1170
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   4710
         Left            =   -74670
         TabIndex        =   25
         Top             =   2385
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8308
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
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
         Height          =   540
         Left            =   105
         TabIndex        =   42
         Top             =   2175
         Width           =   10935
         Begin VB.ComboBox CboVend 
            Height          =   315
            Left            =   915
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   165
            Width           =   3495
         End
         Begin VB.CommandButton cmdNuevoRubro 
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
            Left            =   10380
            MaskColor       =   &H000000FF&
            Picture         =   "frmFacturaClienteConcepto.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Agregar Condición de Venta"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboCondicion 
            Height          =   315
            Left            =   5580
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   165
            Width           =   4770
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Condición:"
            Height          =   195
            Left            =   4770
            TabIndex        =   55
            Top             =   210
            Width           =   750
         End
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
         Height          =   4695
         Left            =   105
         TabIndex        =   29
         Top             =   2640
         Width           =   10935
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   11
            Top             =   3960
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en % "
            Height          =   285
            Left            =   390
            TabIndex        =   10
            Top             =   3660
            Width           =   1290
         End
         Begin VB.TextBox txtSubTotalBoni 
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
            Left            =   4905
            TabIndex        =   56
            Top             =   3990
            Width           =   1155
         End
         Begin VB.TextBox txtImporteIva 
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
            Left            =   7320
            TabIndex        =   52
            Top             =   3990
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7320
            TabIndex        =   13
            Top             =   3660
            Width           =   1155
         End
         Begin VB.TextBox txtImporteBoni 
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
            Left            =   2850
            TabIndex        =   49
            Top             =   3990
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   12
            Top             =   3660
            Width           =   1155
         End
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
            Left            =   9390
            TabIndex        =   46
            Top             =   3990
            Width           =   1350
         End
         Begin VB.TextBox txtSubtotal 
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
            Left            =   9390
            TabIndex        =   45
            Top             =   3660
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1455
            MaxLength       =   60
            TabIndex        =   14
            Top             =   4335
            Width           =   9285
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1140
            TabIndex        =   30
            Top             =   450
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3435
            Left            =   150
            TabIndex        =   9
            Top             =   195
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   6059
            _Version        =   393216
            Rows            =   3
            Cols            =   7
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   4110
            TabIndex        =   57
            Top             =   4050
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6690
            TabIndex        =   54
            Top             =   4035
            Width           =   630
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6660
            TabIndex        =   53
            Top             =   3690
            Width           =   705
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   1890
            TabIndex        =   51
            Top             =   4035
            Width           =   630
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación:"
            Height          =   195
            Left            =   1890
            TabIndex        =   50
            Top             =   3690
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8595
            TabIndex        =   48
            Top             =   4035
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8595
            TabIndex        =   47
            Top             =   3690
            Width           =   750
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   43
            Top             =   4380
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
         TabIndex        =   27
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Facturas"
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
      Left            =   5175
      TabIndex        =   71
      Top             =   7575
      Width           =   2160
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
      TabIndex        =   37
      Top             =   7530
      Width           =   660
   End
End
Attribute VB_Name = "frmFacturaClienteConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim W As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoFactura As Integer
Dim VIva As String
Dim VIvaCalculo As Double
Dim VSucursal As String
Dim VNroFc As String
Dim VBanderaBuscar As Boolean

Private Sub cboFactura_Click()
    txtNroSucursal.Text = ""
    txtNroFactura.Text = ""
End Sub

Private Sub cboRep_Click()
    txtNroSucursal.Text = ""
    txtNroFactura.Text = ""
End Sub

Private Sub chkBonificaEnPesos_Click()
    If chkBonificaEnPesos.Value = Checked Then
        chkBonificaEnPorsentaje.Value = Unchecked
        chkBonificaEnPorsentaje.Enabled = False
    Else
        chkBonificaEnPorsentaje.Enabled = True
    End If
    txtPorcentajeBoni.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
End Sub

Private Sub chkBonificaEnPorsentaje_Click()
    If chkBonificaEnPorsentaje.Value = Checked Then
        chkBonificaEnPesos.Value = Unchecked
        chkBonificaEnPesos.Enabled = False
    Else
        chkBonificaEnPesos.Enabled = True
    End If
    txtPorcentajeBoni.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    sql = "SELECT FC.*, C.CLI_RAZSOC, TC.TCO_ABREVIA"
    sql = sql & " FROM FACTURA_CLIENTE FC, CLIENTE C, TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND FC.FCL_TIPO='C'" 'BUSCA FACTURAS POR CONCEPTO
    If txtCliente.Text <> "" Then sql = sql & " AND FC.CLI_CODIGO=" & XN(txtCliente)
    If FechaDesde <> "" Then sql = sql & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde.Text)
    If FechaHasta <> "" Then sql = sql & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta.Text)
    If cboFactura1.List(cboFactura1.ListIndex) <> "(Todas)" Then sql = sql & " AND FC.TCO_CODIGO=" & XN(cboFactura1.ItemData(cboFactura1.ListIndex))
    If cboBuscaRep.List(cboBuscaRep.ListIndex) <> "(Todas)" Then sql = sql & " AND FC.REP_CODIGO=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
    sql = sql & " ORDER BY FC.REP_CODIGO,FC.FCL_SUCURSAL,FC.FCL_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") _
                            & Chr(9) & rec!FCL_FECHA & Chr(9) & rec!CLI_RAZSOC _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!FCL_BONIFICA _
                            & Chr(9) & rec!FCL_IVA & Chr(9) & rec!FCL_OBSERVACION _
                            & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!FPG_CODIGO _
                            & Chr(9) & rec!FCL_BONIPESOS & Chr(9) & rec!REP_CODIGO _
                            & Chr(9) & rec!CLI_CODIGO & Chr(9) & rec!VEN_CODIGO
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
    frmBuscar.txtDescriB = ""
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
    frmBuscar.txtDescriB = ""
    frmBuscar.Show vbModal
    If frmBuscar.grdBuscar.Text <> "" Then
        frmBuscar.grdBuscar.Col = 0
        txtCodCliente.Text = frmBuscar.grdBuscar.Text
        txtCodCliente_LostFocus
        txtCliRazSoc.SetFocus
    Else
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub CmdGrabar_Click()
    
    If ValidarFactura = False Then Exit Sub
    If MsgBox("¿Confirma Factura?" & Chr(13) & Chr(13) & _
            "Representada: " & cboRep.List(cboRep.ListIndex) & Chr(13) & _
            "Tipo Fac:  " & cboFactura.List(cboFactura.ListIndex) & Chr(13) & _
            "Número:   " & txtNroSucursal.Text & "-" & txtNroFactura.Text, vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM FACTURA_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(cboFactura.ItemData(cboFactura.ListIndex))
    sql = sql & " AND FCL_NUMERO = " & XN(txtNroFactura)
    sql = sql & " AND FCL_SUCURSAL=" & XN(txtNroSucursal)
    sql = sql & " AND REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        'NUEVA FACTURA
        sql = "INSERT INTO FACTURA_CLIENTE"
        sql = sql & " (TCO_CODIGO,FCL_NUMERO,FCL_SUCURSAL,REP_CODIGO,FCL_FECHA,"
        sql = sql & "CLI_CODIGO, VEN_CODIGO, FCL_BONIFICA,FCL_IVA,FPG_CODIGO,FCL_OBSERVACION,"
        sql = sql & "FCL_BONIPESOS,FCL_NUMEROTXT,FCL_SUBTOTAL,FCL_TOTAL,FCL_SALDO,FCL_TIPO,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(cboFactura.ItemData(cboFactura.ListIndex)) & ","
        sql = sql & XN(txtNroFactura.Text) & ","
        sql = sql & XN(txtNroSucursal.Text) & ","
        sql = sql & XN(cboRep.ItemData(cboRep.ListIndex)) & ","
        sql = sql & XDQ(FechaFactura.Text) & ","
        sql = sql & XN(txtCodCliente.Text) & ","
        sql = sql & XN(CboVend.ItemData(CboVend.ListIndex)) & ","
        sql = sql & XN(txtPorcentajeBoni) & ","
        sql = sql & XN(VIva) & "," 'VALOR DEL IVA
        sql = sql & XN(cboCondicion.ItemData(cboCondicion.ListIndex)) & ","
        sql = sql & XS(txtObservaciones) & ","
        If chkBonificaEnPesos.Value = Checked Then
            sql = sql & "'S'" & "," 'BONIFICA EN PESOS
        ElseIf chkBonificaEnPorsentaje.Value = Checked Then
            sql = sql & "'N'" & "," 'BONIFICA EN PORCENTAJE
        Else
            sql = sql & "NULL" & "," 'NO HAY BONIFICACION
        End If
        sql = sql & XS(Format(txtNroFactura.Text, "00000000")) & ","
        If txtSubTotalBoni.Text <> "" Then 'SUBTOTAL
            If txtPorcentajeIva.Text <> "" Then
                sql = sql & XN(Valido_Importe(txtSubTotalBoni.Text)) & ","
            Else
                sql = sql & XN(Valido_Importe(CStr(CDbl(txtSubTotalBoni.Text) / VIvaCalculo))) & ","
            End If
        Else
            If txtPorcentajeIva.Text <> "" Then
                sql = sql & XN(Valido_Importe(txtSubtotal.Text)) & ","
            Else
                sql = sql & XN(Valido_Importe(CStr(CDbl(txtSubtotal.Text) / VIvaCalculo))) & ","
            End If
        End If
        sql = sql & XN(txtTotal) & ","
        sql = sql & XN(txtTotal) & "," 'SALDO FACTURA
        sql = sql & "'C'," 'FACTURA POR CONCEPTO
        sql = sql & "3)" 'ESTADO DEFINITIVO
        DBConn.Execute sql
           
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                sql = "INSERT INTO DETALLE_FACTURA_CLIENTE"
                sql = sql & " (TCO_CODIGO,FCL_NUMERO,FCL_SUCURSAL,REP_CODIGO,FCL_FECHA,"
                sql = sql & "DFC_NROITEM,DFC_CONCEPTO,DFC_CANTIDAD,DFC_PRECIO,DFC_BONIFICA)"
                sql = sql & " VALUES ("
                sql = sql & XN(cboFactura.ItemData(cboFactura.ListIndex)) & ","
                sql = sql & XN(txtNroFactura) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XN(cboRep.ItemData(cboRep.ListIndex)) & ","
                sql = sql & XDQ(FechaFactura) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 6)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(I, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 1)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 2)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(I, 3)) & ")"
                DBConn.Execute sql
            End If
        Next
                 
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA FACTURA QUE CORRESPONDE
        Call ActualizoNumeroComprobantes(cboRep.ItemData(cboRep.ListIndex), cboFactura.ItemData(cboFactura.ListIndex), txtNroFactura.Text)
        
        DBConn.CommitTrans
    Else
        MsgBox "La Factura ya fue Registrada", vbCritical, TIT_MSGBOX
        DBConn.CommitTrans
    End If
    rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    CmdNuevo_Click
    Exit Sub
    
HayErrorFactura:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Function ValidarFactura() As Boolean
    If FechaFactura.Text = "" Then
        MsgBox "La Fecha de la Factura es requerida", vbExclamation, TIT_MSGBOX
        FechaFactura.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If cboCondicion.ListIndex = -1 Then
        MsgBox "La Condición de pago es requerida", vbExclamation, TIT_MSGBOX
        cboCondicion.SetFocus
        ValidarFactura = False
        Exit Function
    End If
    If chkBonificaEnPesos.Value = Checked Or chkBonificaEnPorsentaje.Value = Checked Then
        If txtPorcentajeBoni.Text = "" Then
            MsgBox "Debe ingresar la Bonificación", vbExclamation, TIT_MSGBOX
            txtPorcentajeBoni.SetFocus
            ValidarFactura = False
            Exit Function
        End If
    End If
    ValidarFactura = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("¿Confirma Impresión Factura?" & Chr(13) & Chr(13) & _
            "Representada: " & cboRep.List(cboRep.ListIndex) & Chr(13) & _
            "Tipo Fac:  " & cboFactura.List(cboFactura.ListIndex) & Chr(13) & _
            "Número:   " & txtNroSucursal.Text & "-" & txtNroFactura.Text, vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
         
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
    ImprimirFactura
End Sub

Public Sub ImprimirFactura()
    Dim Renglon As Double
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For W = 1 To 3 'SE IMPRIME POR DUPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA FACTURA ------------------
        Renglon = 9.9
        For I = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(I, 0) <> "" Then
                Printer.FontSize = 8
                Imprimir 1.1, Renglon, False, Format(grdGrilla.TextMatrix(I, 0), "000000")  'codigo
                If Len(grdGrilla.TextMatrix(I, 0)) < 54 Then
                    Imprimir 1.2, Renglon, False, Trim(grdGrilla.TextMatrix(I, 0)) 'descripcion
                Else
                    Imprimir 1.2, Renglon, False, Trim(Left(grdGrilla.TextMatrix(I, 0), 53)) & "..." 'descripcion
                End If
                Printer.FontSize = 9
                Imprimir 12.1, Renglon, False, Trim(grdGrilla.TextMatrix(I, 1)) 'cantidad
                Imprimir 13.5, Renglon, False, Trim(grdGrilla.TextMatrix(I, 2)) 'precio
                Imprimir 15.5, Renglon, False, IIf(grdGrilla.TextMatrix(I, 3) = "", "0,00", Trim(grdGrilla.TextMatrix(I, 3))) 'bonoficacion
                Imprimir 17.5, Renglon, False, Trim(grdGrilla.TextMatrix(I, 5)) 'importe
                Renglon = Renglon + 0.5
            End If
        Next I
        '-----OBSERVACIONES---------------------
        If txtObservaciones.Text <> "" Then
            Imprimir 1.2, Renglon + 1, False, "Observaciones: " & Trim(txtObservaciones.Text)
        End If
        '---------PORCUENTA Y ORDEN----------------------
        Imprimir 1.5, 19.5, True, "Por Cuenta y Orden de: " & Trim(cboRep.List(cboRep.ListIndex))
        '-------------IMPRIMO TOTALES--------------------
        Printer.FontSize = 10
        Imprimir 14, 22, False, "Sub-Total"
        Imprimir 17, 22, True, Trim(txtSubtotal.Text)
        'Imprimir 0.3, 18.9, True, txtSubtotal.Text
        If txtPorcentajeBoni.Text <> "" Then
            If chkBonificaEnPesos.Value = Checked Then
                Imprimir 14, 22.5, False, "Bonificación ($)"
                Imprimir 17, 22.5, True, txtPorcentajeBoni.Text
                Imprimir 14, 23, False, "Imp. Bonif."
                Imprimir 17, 23, True, txtImporteBoni.Text
            Else
                Imprimir 14, 22.5, False, "Bonificación (%)"
                Imprimir 17, 22.5, True, txtPorcentajeBoni.Text
                Imprimir 14, 23, False, "Imp. Bonif."
                Imprimir 17, 23, True, txtImporteBoni.Text
            End If
            Imprimir 14, 23.5, False, "Sub-Total"
            Imprimir 17, 23.5, True, txtSubTotalBoni.Text
            
            Imprimir 14, 24, False, "% I.V.A."
            Imprimir 17, 24, True, txtPorcentajeIva.Text
            Imprimir 14, 24.5, False, "I.V.A."
            Imprimir 17, 24.5, True, txtImporteIva.Text
            Imprimir 14, 25, True, "Total"
            Imprimir 17, 25, True, txtTotal.Text
        Else
            Imprimir 14, 22.5, False, "% I.V.A."
            Imprimir 17, 22.5, True, txtPorcentajeIva.Text
            Imprimir 14, 23, False, "I.V.A."
            Imprimir 17, 23, True, txtImporteIva.Text
            Imprimir 14, 23.5, True, "Total"
            Imprimir 17, 23.5, True, txtTotal.Text
        End If
        Printer.EndDoc
    Next W
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
End Sub

Public Sub ImprimirEncabezado()
 '-----------IMPRIME EL ENCABEZADO DE LA FACTURA-------------------
    Printer.FontSize = 10
    Imprimir 13.6, 0.6, True, Trim(cboFactura.List(cboFactura.ListIndex)) & "   Nº " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroFactura.Text)
    Imprimir 15.5, 2.1, False, Format(FechaFactura, "dd/mm/yyyy")

    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI,C.CLI_CUIT,C.CLI_INGBRU, L.LOC_DESCRI"
    sql = sql & ", P.PRO_DESCRI,CI.IVA_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L,"
    sql = sql & " PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE"
    sql = sql & " C.CLI_CODIGO=" & XN(txtCodCliente.Text)
    sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
    sql = sql & " AND C.PRO_CODIGO=L.PRO_CODIGO"
    sql = sql & " AND C.PAI_CODIGO=L.PAI_CODIGO"
    sql = sql & " AND L.PRO_CODIGO=P.PRO_CODIGO"
    sql = sql & " AND C.IVA_CODIGO=CI.IVA_CODIGO"

    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        Imprimir 1.3, 4.6, True, "(" & Trim(Rec1!CLI_CODIGO) & ") " & Trim(Rec1!CLI_RAZSOC) & _
                               "  - Vend: (" & CboVend.ItemData(CboVend.ListIndex) & ")" & Trim(CboVend.List(CboVend.ListIndex))
        
        Imprimir 1.3, 5, False, Trim(Rec1!CLI_DOMICI)
        'REMITO
        Imprimir 1.3, 5.4, False, "Loc: " & Trim(Rec1!LOC_DESCRI) & " -- Prov: " & Trim(Rec1!PRO_DESCRI)
        Imprimir 1.7, 6.2, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 7.9, 6.2, False, IIf(IsNull(Rec1!CLI_CUIT), "NO INFORMADO", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 15.7, 6.2, False, IIf(IsNull(Rec1!CLI_INGBRU), "NO INFORMADO", Format(Rec1!CLI_INGBRU, "###-#####-##"))
    End If
    
    Rec1.Close
    Imprimir 4.8, 7.5, False, cboCondicion.Text
    'Imprimir 1.1, 9.2, False, "Código"
    Imprimir 1.2, 9.2, False, "Descripción"
    Imprimir 12, 9.2, False, "Cant."
    Imprimir 13.5, 9.2, False, "P.Unit."
    Imprimir 15.5, 9.2, False, "Bonif."
    Imprimir 17.5, 9.2, False, "Importe"
End Sub

Private Sub CmdNuevo_Click()
    'significa que no estoy buacando
    VBanderaBuscar = False
    
    'LIMPIO FACTURA
    Limpiar_Factura
    cboRep.SetFocus
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
    VEstadoFactura = 1
    tabDatos.Tab = 0
   
End Sub

Private Sub Limpiar_Factura()
    For I = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(I, 0) = ""
        grdGrilla.TextMatrix(I, 1) = ""
        grdGrilla.TextMatrix(I, 2) = ""
        grdGrilla.TextMatrix(I, 3) = ""
        grdGrilla.TextMatrix(I, 4) = ""
        grdGrilla.TextMatrix(I, 5) = ""
        grdGrilla.TextMatrix(I, 6) = I
    Next
    txtCodCliente.Text = ""
    txtNroFactura.Text = ""
    txtNroSucursal.Text = ""
    FechaFactura.Text = Date
    lblEstadoFactura.Caption = ""
    txtSubtotal.Text = ""
    txtTotal.Text = ""
    txtPorcentajeBoni.Text = ""
    txtPorcentajeIva.Text = ""
    txtImporteBoni.Text = ""
    txtSubTotalBoni.Text = ""
    txtImporteIva.Text = ""
    txtObservaciones.Text = ""
    cboCondicion.ListIndex = 0
    lblEstado.Caption = ""
    cmdGrabar.Enabled = True
    chkBonificaEnPorsentaje.Value = Unchecked
    chkBonificaEnPesos.Value = Unchecked
    FrameFactura.Enabled = True
    FrameCliente.Enabled = True
    FechaFactura.Text = Date
    CboVend.ListIndex = 0
    cboFactura.ListIndex = 0
    cboRep.ListIndex = 0
End Sub

Private Sub cmdNuevoRubro_Click()
    ABMFormaPago.Show vbModal
    cboCondicion.Clear
    LlenarComboFormaPago
    cboCondicion.SetFocus
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmFacturaClienteConcepto = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
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
    
    
    grdGrilla.FormatString = "Descripción|Cantidad|Precio|Bonif.|Pre.Bonif.|Importe|Orden"
    grdGrilla.ColWidth(0) = 5600 'DESCRIPCION
    grdGrilla.ColWidth(1) = 900  'CANTIDAD
    grdGrilla.ColWidth(2) = 1100 'PRECIO
    grdGrilla.ColWidth(3) = 900  'BONOFICACION
    grdGrilla.ColWidth(4) = 1000 'PRE BONIFICACION
    grdGrilla.ColWidth(5) = 1100 'IMPORTE
    grdGrilla.ColWidth(6) = 0    'ORDEN
    grdGrilla.Cols = 7
    grdGrilla.Rows = 1
    
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For I = 0 To 6
        grdGrilla.Col = I
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    
    For I = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & (I - 1)
    Next
    
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Comp.|^Número|^Fecha|Cliente|Cod_Estado|" _
                              & "PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE|COD CONDICION VENTA|" _
                              & "BONIFICA EN PESOS|REPRESENTADA|COD CLIENTE|VENDEDOR"
    GrdModulos.ColWidth(0) = 1000  'TIPO FACTURA
    GrdModulos.ColWidth(1) = 1600 'NUMERO
    GrdModulos.ColWidth(2) = 1200 'FECHA
    GrdModulos.ColWidth(3) = 6000 'CLIENTE
    GrdModulos.ColWidth(4) = 0    'COD_ESTADO
    GrdModulos.ColWidth(5) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(6) = 0    'PORCENTAJE IVA
    GrdModulos.ColWidth(7) = 0    'OBSERVACIONES
    GrdModulos.ColWidth(8) = 0    'COD TIPO COMPROBANTE
    GrdModulos.ColWidth(9) = 0    'COD CONDICION VENTA
    GrdModulos.ColWidth(10) = 0   'BONIFICA EN PESOS
    GrdModulos.ColWidth(11) = 0   'REPRESENTADA
    GrdModulos.ColWidth(12) = 0   'CODIGO CLIENTE
    GrdModulos.ColWidth(13) = 0   'VENDEDOR
    GrdModulos.Rows = 1
    '------------------------------------
    GrdModulos.BorderStyle = flexBorderNone
    
    GrdModulos.row = 0
    For I = 0 To 3
        GrdModulos.Col = I
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    '------------------------------------
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE FACTURA
    LlenarComboFactura
    'CARGO COMBO CON LAS CONDICIONES DE VENTA
    LlenarComboFormaPago
    
    'CARGO COMBO VENDEDOR
    Call CargoComboBox(CboVend, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE")
    CboVend.ListIndex = 0
    
    'CRAGO COMBO REPRESENTADA
    Call CargoComboBox(cboRep, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboRep.ListIndex = 0
    
    Call CargoComboBox(cboBuscaRep, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboBuscaRep.AddItem "(Todas)"
    cboBuscaRep.ListIndex = cboBuscaRep.ListCount - 1
    Call BuscaCodigoProxItemData(CInt(VRepresentada), cboRep)
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoFactura) 'ESTADO PENDIENTE
    VEstadoFactura = 1
    FechaFactura.Text = Date
    tabDatos.Tab = 0
    
    'BUSCO IVA
    sql = "SELECT IVA FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        VIva = IIf(IsNull(rec!iva), "", Format(rec!iva, "0.00"))
        VIvaCalculo = (CDbl(VIva) / 100) + 1
    End If
    rec.Close
    If cboFactura.ItemData(cboFactura.ListIndex) <> 2 Then
        txtPorcentajeIva.Text = VIva
    Else
        txtPorcentajeIva.Text = ""
    End If
    'significa que no estoy buacando
    VBanderaBuscar = False
End Sub

Private Sub LlenarComboFactura()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'FACT%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboFactura1.AddItem "(Todas)"
        Do While rec.EOF = False
            cboFactura.AddItem rec!TCO_DESCRI
            cboFactura.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            cboFactura1.AddItem rec!TCO_DESCRI
            cboFactura1.ItemData(cboFactura.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboFactura.ListIndex = 0
        cboFactura1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboFormaPago()
    sql = "SELECT * FROM FORMA_PAGO"
    sql = sql & " ORDER BY FPG_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboCondicion.AddItem rec!FPG_DESCRI
            cboCondicion.ItemData(cboCondicion.NewIndex) = rec!FPG_CODIGO
            rec.MoveNext
        Loop
        cboCondicion.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0
            If MsgBox("Seguro que desea quitar", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = grdGrilla.RowSel
                grdGrilla.Col = 0
            
                txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                txtTotal.Text = txtSubtotal.Text
                txtPorcentajeBoni_LostFocus
                txtPorcentajeIva_LostFocus
            End If
        
        Case 1, 2
            VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(txtEdit.Text)) / 100)
            VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) - VBonificacion)
            grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Valido_Importe(CStr(VBonificacion))
            grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            txtPorcentajeBoni_LostFocus
            txtPorcentajeIva_LostFocus
            
        Case 3
            VBonificacion = 0
            grdGrilla.Text = ""
            grdGrilla.Col = 4
            grdGrilla.Text = ""
            grdGrilla.Col = 5
            grdGrilla.Text = ""
            
            VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)))
            grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            txtPorcentajeBoni_LostFocus
            txtPorcentajeIva_LostFocus
            grdGrilla.Col = 3
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 1
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" Then
                    chkBonificaEnPorsentaje.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Then
       
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 3 Then
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
            If grdGrilla.Col = 2 Or grdGrilla.Col = 3 Then
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
    grdGrilla = txtEdit.Text
    txtEdit.Visible = False
End Sub

Private Sub grdGrilla_GotFocus()
    If grdGrilla.Rows > 1 Then
        If txtEdit.Visible = False Then
            grdGrilla.Col = 0
            Exit Sub
        End If
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
        
        Set Rec1 = New ADODB.Recordset
        Limpiar_Factura
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
        
        'significa que estoy buacando
        VBanderaBuscar = True
        
        txtCodCliente.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 12)
        txtCodCliente_LostFocus
        
        'CABEZA FACTURA
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 8)), cboFactura)
        'BUSCO REPRESENTADA
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 11)), cboRep)
        
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        txtNroFactura.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        FechaFactura.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)

        Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), lblEstadoFactura)
        VEstadoFactura = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4))
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) <> "" Then
            txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
        End If

        'CONDICION VENTA
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 9)), cboCondicion)
        'VENDEDOR
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 13)), CboVend)
        '----BUSCO DETALLE DE LA FACTURA------------------
        sql = "SELECT DFC.* "
        sql = sql & " FROM DETALLE_FACTURA_CLIENTE DFC"
        sql = sql & " WHERE DFC.FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND DFC.FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND DFC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 8))
        sql = sql & " AND DFC.REP_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 11))
        sql = sql & " ORDER BY DFC.DFC_NROITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            I = 1
            Do While Rec1.EOF = False
                grdGrilla.TextMatrix(I, 0) = Rec1!DFC_CONCEPTO
                grdGrilla.TextMatrix(I, 1) = ChkNull(Rec1!DFC_CANTIDAD)
                If Not IsNull(Rec1!DFC_PRECIO) Then
                    grdGrilla.TextMatrix(I, 2) = Valido_Importe((Rec1!DFC_PRECIO))
                Else
                    grdGrilla.TextMatrix(I, 2) = ChkNull(Rec1!DFC_PRECIO)
                End If
                If IsNull(Rec1!DFC_BONIFICA) Then
                    grdGrilla.TextMatrix(I, 3) = ""
                Else
                    grdGrilla.TextMatrix(I, 3) = Valido_Importe(Rec1!DFC_BONIFICA)
                End If
                VBonificacion = 0
                If Not IsNull(Rec1!DFC_BONIFICA) Then
                    VBonificacion = (((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) * CDbl(Rec1!DFC_BONIFICA)) / 100)
                    VBonificacion = ((CDbl(Rec1!DFC_CANTIDAD) * CDbl(Rec1!DFC_PRECIO)) - VBonificacion)
                    grdGrilla.TextMatrix(I, 4) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
                Else
                    If Not IsNull(Rec1!DFC_CANTIDAD) And Not IsNull(Rec1!DFC_PRECIO) Then
                        VBonificacion = (CDbl(ChkNull(Rec1!DFC_CANTIDAD)) * CDbl(ChkNull(Rec1!DFC_PRECIO)))
                        grdGrilla.TextMatrix(I, 4) = ""
                        grdGrilla.TextMatrix(I, 5) = Valido_Importe(CStr(VBonificacion))
                    End If
                End If
                grdGrilla.TextMatrix(I, 6) = Rec1!DFC_NROITEM
                I = I + 1
                Rec1.MoveNext
            Loop
            VBonificacion = 0
        End If
        Rec1.Close
        '--CARGO LOS TOTALES----
        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
        txtTotal.Text = txtSubtotal.Text
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 10) = "S" Then
            chkBonificaEnPesos.Value = Checked
        ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 10) = "N" Then
            chkBonificaEnPorsentaje.Value = Checked
        Else
            chkBonificaEnPesos.Value = Unchecked
            chkBonificaEnPorsentaje.Value = Unchecked
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) <> "" Then
            txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
            txtPorcentajeBoni_LostFocus
        End If
        
        If cboFactura.ItemData(cboFactura.ListIndex) <> 2 And GrdModulos.TextMatrix(GrdModulos.RowSel, 6) <> "" Then
            txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            txtPorcentajeIva_LostFocus
        Else
            txtPorcentajeIva.Text = ""
        End If
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        '--------------
        FrameFactura.Enabled = False
        FrameCliente.Enabled = False
        '--------------
        'significa que estoy buacando
        VBanderaBuscar = True
        tabDatos.Tab = 0
        cboCondicion.SetFocus
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then txtCliente.SetFocus
    frameBuscar.Caption = "Buscar Factura por..."
  Else
    If VEstadoFactura = 1 Then
        cmdGrabar.Enabled = True
    Else
        cmdGrabar.Enabled = False
    End If
  End If
End Sub

Private Sub LimpiarBusqueda()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    cboFactura1.ListIndex = 0
    cboBuscaRep.ListIndex = cboBuscaRep.ListCount - 1
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

Private Sub txtCliRazSoc_Change()
    If txtCliRazSoc.Text = "" Then
        txtCodCliente.Text = ""
        txtProvincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCliRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtCliRazSoc_LostFocus()
    If txtCodCliente.Text = "" And txtCliRazSoc.Text <> "" Then
        rec.Open BuscoCliente(txtCliRazSoc), DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                frmBuscar.TipoBusqueda = 1
                frmBuscar.txtDescriB.Text = txtCliRazSoc.Text
                frmBuscar.Show vbModal
                If frmBuscar.grdBuscar.Text <> "" Then
                    frmBuscar.grdBuscar.Col = 0
                    txtCodCliente.Text = frmBuscar.grdBuscar.Text
                    frmBuscar.grdBuscar.Col = 1
                    txtCliRazSoc.Text = frmBuscar.grdBuscar.Text
                    txtCodCliente_LostFocus
                    FechaDesde.SetFocus
                Else
                    txtCodCliente.SetFocus
                End If
            Else
                txtCodCliente.Text = rec!CLI_CODIGO
                txtCliRazSoc.Text = rec!CLI_RAZSOC
                txtCodCliente_LostFocus
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        rec.Close
    ElseIf txtCodCliente.Text = "" And txtCliRazSoc.Text = "" Then
        MsgBox "Debe elegir un cliente", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
    End If
End Sub

Private Sub txtCodCliente_Change()
    If txtCodCliente.Text = "" Then
        txtCliRazSoc.Text = ""
        txtProvincia.Text = ""
        txtCliLocalidad.Text = ""
        txtDomici.Text = ""
    End If
End Sub

Private Sub txtCodCliente_GotFocus()
    SelecTexto txtCodCliente
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCliente_LostFocus()
    If txtCodCliente.Text <> "" Then
        rec.Open BuscoCliente(txtCodCliente), DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then
            txtCliRazSoc.Text = rec!CLI_RAZSOC
            txtProvincia.Text = rec!PRO_DESCRI
            txtCliLocalidad.Text = rec!LOC_DESCRI
            txtDomici.Text = rec!CLI_DOMICI
            
            If rec!IVA_CODIGO = 1 Then 'RESPONSABLE INSCRIPTO
                   Call BuscaProx("FACTURA A", cboFactura)
                If CInt(VRepresentada) = cboRep.ItemData(cboRep.ListIndex) Then
                   'or CInt(VRepresentada2) = cboRep.ItemData(cboRep.ListIndex) Then
                    txtNroSucursal.Text = ""
                    txtNroSucursal_LostFocus
                    txtNroFactura.Text = ""
                    txtNroFactura_LostFocus
                Else
                    txtNroSucursal.Text = VSucursal
                    txtNroFactura.Text = VNroFc
                    txtPorcentajeIva.Text = VIva
                End If
            ElseIf rec!IVA_CODIGO = 2 Or rec!IVA_CODIGO = 3 Then
                Call BuscaProx("FACTURA B", cboFactura)
                If CInt(VRepresentada) = cboRep.ItemData(cboRep.ListIndex) Then
                   'or CInt(VRepresentada2) = cboRep.ItemData(cboRep.ListIndex) Then
                    txtNroSucursal.Text = ""
                    txtNroSucursal_LostFocus
                    txtNroFactura.Text = ""
                    txtNroFactura_LostFocus
                Else
                    txtNroSucursal.Text = VSucursal
                    txtNroFactura.Text = VNroFc
                    txtPorcentajeIva.Text = ""
                End If
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Function BuscoCliente(Codigo As String) As String
        sql = "SELECT C.CLI_CODIGO, C.IVA_CODIGO, C.CLI_RAZSOC, C.CLI_DOMICI, P.PRO_DESCRI, L.LOC_DESCRI"
        sql = sql & " FROM CLIENTE C,  PROVINCIA P, LOCALIDAD L"
        sql = sql & " WHERE"
        If txtCodCliente.Text <> "" Then
            sql = sql & " C.CLI_CODIGO=" & XN(Codigo)
        Else
            sql = sql & " C.CLI_RAZSOC LIKE '" & Trim(Codigo) & "%'"
        End If
        sql = sql & " AND C.PRO_CODIGO=P.PRO_CODIGO"
        sql = sql & " AND C.LOC_CODIGO=L.LOC_CODIGO"
        BuscoCliente = sql
End Function

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0
    If grdGrilla.Col = 1 Then KeyAscii = CarNumeroEntero(KeyAscii)
    If grdGrilla.Col = 2 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    If grdGrilla.Col = 3 Then KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
    CarTexto (KeyAscii)
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    VBonificacion = 0
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            Case 1 'CANTIDAD
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) = "" Then
                        grdGrilla.Text = "1"
                        txtEdit.Text = "1"
                    End If
                End If
            
            Case 2 'PRECIO
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) <> "" Then
                        txtEdit.Text = Valido_Importe(txtEdit.Text)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = txtEdit.Text
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(CInt(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2))))
                        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                        txtTotal.Text = txtSubtotal.Text
                        txtPorcentajeBoni_LostFocus
                        txtPorcentajeIva_LostFocus
                    Else
                        MsgBox "Debe ingresar el Precio", vbExclamation, TIT_MSGBOX
                        grdGrilla.Col = 1
                    End If
                Else
                    txtEdit.Text = ""
                End If
                
            Case 3 'BONIFICACION
                If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) <> "" Then
                    If Trim(txtEdit) <> "" Then
                        If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                            Exit Sub
                        End If
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) * CDbl(txtEdit.Text)) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 1)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 2)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = Valido_Importe(CStr(VBonificacion))
                        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                        txtTotal.Text = txtSubtotal.Text
                        txtPorcentajeBoni_LostFocus
                        txtPorcentajeIva_LostFocus
                    Else
                        MsgBox "Debe ingresar el Porcentaje", vbExclamation, TIT_MSGBOX
                        grdGrilla.Col = 3
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

Private Function BuscoRepetetidos(Codigo As Long, Linea As Integer) As Boolean
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 0) <> "" Then
            If Codigo = CLng(grdGrilla.TextMatrix(I, 0)) And (I <> Linea) Then
                MsgBox "El producto ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

Private Sub txtNroFactura_GotFocus()
    SelecTexto txtNroFactura
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroFactura_LostFocus()
    If VBanderaBuscar = False Then
        If txtNroFactura.Text = "" Then
            'BUSCO EL NUMERO DE FACTURA QUE CORRESPONDE
            txtNroFactura.Text = BuscoUltimoNumeroComprobante(cboRep.ItemData(cboRep.ListIndex), cboFactura.ItemData(cboFactura.ListIndex))
        Else
            txtNroFactura.Text = Format(txtNroFactura.Text, "00000000")
        End If
        If cboFactura.ItemData(cboFactura.ListIndex) <> 2 Then
            txtPorcentajeIva.Text = VIva
        Else
            txtPorcentajeIva.Text = ""
        End If
        VSucursal = txtNroSucursal.Text
        VNroFc = txtNroFactura.Text
    End If
End Sub

Private Function SumaTotal() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 6) <> "" Then
            VTotal = VTotal + (CInt(grdGrilla.TextMatrix(I, 2)) * CDbl(grdGrilla.TextMatrix(I, 3)))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 5) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(I, 5))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Sub txtNroSucursal_GotFocus()
    SelecTexto txtNroSucursal
End Sub

Private Sub txtNroSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroSucursal_LostFocus()
    If txtNroSucursal.Text = "" Then
'        Select Case cboRep.ItemData(cboRep.ListIndex)
'        Case CInt(VRepresentada2) 'ESTA ES DE SERVIPACK S.R.L.
'            txtNroSucursal.Text = Sucursal2
'        Case Else 'ESTA ES DE ESTILO S.R.L. O CUALQUIER OTRA REPRESENTADA
            txtNroSucursal.Text = Sucursal
'        End Select
    Else
        txtNroSucursal.Text = Format(txtNroSucursal.Text, "0000")
    End If
End Sub

Private Sub txtObservaciones_GotFocus()
    SelecTexto txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtPorcentajeBoni_GotFocus()
    SelecTexto txtPorcentajeBoni
End Sub

Private Sub txtPorcentajeBoni_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeBoni, KeyAscii)
End Sub

Private Sub txtPorcentajeBoni_LostFocus()
    If txtPorcentajeBoni.Text <> "" And txtSubtotal.Text <> "" Then
        If chkBonificaEnPorsentaje.Value = Checked Then
            If ValidarPorcentaje(txtPorcentajeBoni) = False Then
                txtPorcentajeBoni.SetFocus
                Exit Sub
            End If
            txtImporteBoni.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeBoni.Text)) / 100
            txtImporteBoni.Text = Valido_Importe(txtImporteBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
            txtPorcentajeIva_LostFocus
        ElseIf chkBonificaEnPesos.Value = Checked Then
            txtPorcentajeBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtImporteBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
            txtPorcentajeIva_LostFocus
        Else
            txtPorcentajeBoni.Text = ""
            txtImporteBoni.Text = ""
            MsgBox "Debe elegir como bonifica", vbExclamation, TIT_MSGBOX
            chkBonificaEnPorsentaje.SetFocus
        End If
    End If
End Sub

Private Sub txtPorcentajeIva_GotFocus()
    SelecTexto txtPorcentajeIva
End Sub

Private Sub txtPorcentajeIva_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtPorcentajeIva, KeyAscii)
End Sub

Private Sub txtPorcentajeIva_LostFocus()
    If txtSubtotal.Text <> "" Then
        If txtPorcentajeIva.Text <> "" Then
            If ValidarPorcentaje(txtPorcentajeIva) = False Then
                txtPorcentajeIva.SetFocus
                Exit Sub
            End If
            If txtImporteBoni.Text <> "" Then
                txtImporteIva.Text = (CDbl(txtSubTotalBoni.Text) * CDbl(txtPorcentajeIva.Text)) / 100
                txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
                txtTotal.Text = CDbl(txtSubTotalBoni.Text) + CDbl(txtImporteIva.Text)
                txtTotal.Text = Valido_Importe(txtTotal.Text)
            Else
                txtImporteIva.Text = (CDbl(txtSubtotal.Text) * CDbl(txtPorcentajeIva.Text)) / 100
                txtImporteIva.Text = Valido_Importe(txtImporteIva.Text)
                txtTotal.Text = CDbl(txtSubtotal.Text) + CDbl(txtImporteIva.Text)
                txtTotal.Text = Valido_Importe(txtTotal.Text)
            End If
        End If
    End If
End Sub

