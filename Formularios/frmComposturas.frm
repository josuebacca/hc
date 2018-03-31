VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmComposturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Composturas de Clientes..."
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
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
   ScaleHeight     =   8055
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   690
      Left            =   9060
      Picture         =   "frmComposturas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7350
      Width           =   765
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Sobre"
      Height          =   690
      Left            =   6705
      Picture         =   "frmComposturas.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7350
      Width           =   765
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nueva"
      Height          =   690
      Left            =   8280
      Picture         =   "frmComposturas.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7350
      Width           =   765
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Facturar"
      Height          =   690
      Left            =   7485
      Picture         =   "frmComposturas.frx":1DD6
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7350
      Width           =   765
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7290
      Left            =   30
      TabIndex        =   37
      Top             =   60
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   12859
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
      TabPicture(0)   =   "frmComposturas.frx":26A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freRemito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameCliente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAviso"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmComposturas.frx":26BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtAviso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   4965
         Width           =   8600
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   6120
         TabIndex        =   83
         Top             =   5520
         Width           =   3555
         Begin VB.OptionButton OptSComp 
            Caption         =   "S/Comprobante"
            Height          =   195
            Left            =   1920
            TabIndex        =   85
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptCComp 
            Caption         =   "C/Comprobante"
            Height          =   195
            Left            =   360
            TabIndex        =   84
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Estados..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   120
         TabIndex        =   68
         Top             =   6240
         Width           =   9600
         Begin VB.ComboBox cboEntrego 
            Height          =   315
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   570
            Width           =   3015
         End
         Begin VB.TextBox txtRetiro 
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
            Left            =   6450
            TabIndex        =   21
            Top             =   240
            Width           =   2940
         End
         Begin VB.CheckBox chkEntregadoSArreglar 
            Caption         =   "Entregado S/Arreglar"
            Height          =   255
            Left            =   1155
            TabIndex        =   18
            Top             =   585
            Width           =   1950
         End
         Begin VB.CheckBox chkEntregado 
            Caption         =   "Entregado"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   585
            Width           =   1065
         End
         Begin VB.CheckBox chkNoArreglado 
            Caption         =   "No Arreglado"
            Height          =   255
            Left            =   1155
            TabIndex        =   16
            Top             =   225
            Width           =   1590
         End
         Begin VB.CheckBox chkArreglado 
            Caption         =   "Arreglado"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   225
            Width           =   1065
         End
         Begin FechaCtl.Fecha FecArreglo 
            Height          =   285
            Left            =   4275
            TabIndex        =   19
            Top             =   225
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FecEntrega 
            Height          =   285
            Left            =   4275
            TabIndex        =   20
            Top             =   555
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Entrego:"
            Height          =   195
            Left            =   5640
            TabIndex        =   81
            Top             =   615
            Width           =   630
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Retiro:"
            Height          =   195
            Left            =   5640
            TabIndex        =   80
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Entrega:"
            Height          =   195
            Left            =   3165
            TabIndex        =   70
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Arreglo:"
            Height          =   195
            Left            =   3165
            TabIndex        =   69
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Presupuesto..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   105
         TabIndex        =   54
         Top             =   4170
         Width           =   9600
         Begin VB.CheckBox chkSaldo 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   9120
            TabIndex        =   79
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox txtSaldo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7710
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   270
            Width           =   1335
         End
         Begin VB.CheckBox chkTotal 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2400
            TabIndex        =   77
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   990
            TabIndex        =   11
            Top             =   270
            Width           =   1335
         End
         Begin VB.CheckBox chkSena 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   5640
            TabIndex        =   76
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox txtSena 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4230
            TabIndex        =   12
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   420
            TabIndex        =   57
            Top             =   330
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Seña:"
            Height          =   195
            Left            =   3660
            TabIndex        =   56
            Top             =   330
            Width           =   420
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Saldo:"
            Height          =   195
            Left            =   7140
            TabIndex        =   55
            Top             =   330
            Width           =   450
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
         Height          =   1470
         Left            =   -74805
         TabIndex        =   72
         Top             =   435
         Width           =   9330
         Begin VB.TextBox txtBusDesApa 
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
            Left            =   3180
            TabIndex        =   26
            Top             =   600
            Width           =   4150
         End
         Begin VB.TextBox txtBusCodApa 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2490
            TabIndex        =   25
            Top             =   600
            Width           =   675
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   915
            Left            =   8160
            MaskColor       =   &H000000FF&
            TabIndex        =   29
            ToolTipText     =   "Buscar "
            Top             =   315
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox txtBuscaCliente 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2490
            MaxLength       =   40
            TabIndex        =   23
            Top             =   210
            Width           =   675
         End
         Begin VB.TextBox txtBuscarCliDescri 
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
            Left            =   3180
            MaxLength       =   50
            TabIndex        =   24
            Tag             =   "Descripción"
            Top             =   210
            Width           =   4155
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   6210
            TabIndex        =   28
            Top             =   1035
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   2490
            TabIndex        =   27
            Top             =   1035
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Aparato:"
            Height          =   195
            Left            =   1680
            TabIndex        =   87
            Top             =   660
            Width           =   645
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1395
            TabIndex        =   75
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   5130
            TabIndex        =   74
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   1755
            TabIndex        =   73
            Top             =   255
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Orden de Arreglo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   58
         Top             =   5520
         Width           =   6000
         Begin VB.CheckBox chkOrden 
            Caption         =   "Orden de Arreglo"
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin FechaCtl.Fecha fechaOrden 
            Height          =   285
            Left            =   3840
            TabIndex        =   14
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Arreglo:"
            Height          =   195
            Left            =   2760
            TabIndex        =   82
            Top             =   285
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Recepción..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   120
         TabIndex        =   59
         Top             =   2385
         Width           =   9600
         Begin VB.TextBox txtProbReal 
            Height          =   630
            Left            =   6300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1035
            Width           =   3135
         End
         Begin VB.TextBox txtProbCliente 
            Height          =   630
            Left            =   990
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1035
            Width           =   3135
         End
         Begin VB.TextBox txtEstadoActual 
            Height          =   630
            Left            =   6300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cboDestino 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cboVendedor 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   570
            Width           =   3135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Problema Real"
            Height          =   195
            Left            =   5115
            TabIndex        =   67
            Top             =   1035
            Width           =   1020
         End
         Begin VB.Label Label8 
            Caption         =   "Problema Según Cliente"
            Height          =   615
            Left            =   180
            TabIndex        =   66
            Top             =   1035
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Estado Actual"
            Height          =   195
            Left            =   5115
            TabIndex        =   65
            Top             =   330
            Width           =   990
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Recibio:"
            Height          =   195
            Left            =   180
            TabIndex        =   64
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Destino:"
            Height          =   195
            Left            =   180
            TabIndex        =   63
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aparato..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   60
         Top             =   1725
         Width           =   9600
         Begin VB.TextBox txtCodMarca 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   5535
            TabIndex        =   4
            Top             =   255
            Width           =   675
         End
         Begin VB.TextBox txtDesMarca 
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
            Left            =   6225
            TabIndex        =   5
            Top             =   255
            Width           =   2985
         End
         Begin VB.TextBox txtCodAparato 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   990
            TabIndex        =   2
            Top             =   255
            Width           =   675
         End
         Begin VB.TextBox txtDesAparato 
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
            Left            =   1680
            TabIndex        =   3
            Top             =   255
            Width           =   2985
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Aparato:"
            Height          =   195
            Left            =   180
            TabIndex        =   62
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
            Height          =   195
            Left            =   4965
            TabIndex        =   61
            Top             =   315
            Width           =   495
         End
      End
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
         Height          =   1395
         Left            =   120
         TabIndex        =   46
         Top             =   345
         Width           =   6675
         Begin VB.CommandButton cmdbuscaComp 
            Height          =   350
            Left            =   6120
            Picture         =   "frmComposturas.frx":26D8
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Buscar Composturas del Cliente"
            Top             =   280
            Width           =   375
         End
         Begin VB.TextBox txtCiva 
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            TabIndex        =   48
            Top             =   990
            Width           =   2745
         End
         Begin VB.TextBox txtDomici 
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            TabIndex        =   47
            Top             =   645
            Width           =   5520
         End
         Begin VB.TextBox txtRazSoc 
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
            Left            =   1890
            TabIndex        =   1
            Top             =   300
            Width           =   4140
         End
         Begin VB.TextBox txtcodCli 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   990
            TabIndex        =   0
            Top             =   300
            Width           =   870
         End
         Begin MSMask.MaskEdBox txtCuit 
            Height          =   315
            Left            =   5025
            TabIndex        =   49
            Top             =   990
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   13
            Mask            =   "##-########-#"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   " I.V.A.:"
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   1050
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   180
            TabIndex        =   52
            Top             =   690
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "C.U.I.T.:"
            Height          =   195
            Index           =   10
            Left            =   4305
            TabIndex        =   50
            Top             =   1050
            Width           =   660
         End
      End
      Begin VB.Frame freRemito 
         Caption         =   "Compostura..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   6810
         TabIndex        =   38
         Top             =   345
         Width           =   2835
         Begin VB.TextBox txtNroCompostura 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1005
            MaxLength       =   8
            TabIndex        =   39
            Top             =   300
            Width           =   1125
         End
         Begin FechaCtl.Fecha FechaCompostura 
            Height          =   285
            Left            =   1005
            TabIndex        =   40
            Top             =   660
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   330
            TabIndex        =   44
            Top             =   1035
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   330
            TabIndex        =   43
            Top             =   345
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   450
            TabIndex        =   42
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblEstadoCompostura 
            AutoSize        =   -1  'True
            Caption         =   "EST. COMPOSTURA"
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
            Left            =   1005
            TabIndex        =   41
            Top             =   1050
            Width           =   1560
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   5265
         Left            =   -74820
         TabIndex        =   30
         Top             =   1980
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   12
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
      Begin VB.Label Label24 
         Caption         =   "Aviso o Promoción"
         Height          =   435
         Left            =   270
         TabIndex        =   71
         Top             =   5070
         Width           =   750
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
         TabIndex        =   45
         Top             =   570
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   4200
      Top             =   7560
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
      Left            =   120
      TabIndex        =   36
      Top             =   7635
      Width           =   660
   End
End
Attribute VB_Name = "frmComposturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function ValidarCompostura() As Boolean
    If FechaCompostura.Text = "" Then
        MsgBox "La Fecha de la Compostura es requerida", vbExclamation, TIT_MSGBOX
        FechaCompostura.SetFocus
        ValidarCompostura = False
        Exit Function
    End If
    If txtcodCli.Text = "" Then
        MsgBox "Debe ingresar un Cliente", vbExclamation, TIT_MSGBOX
        txtcodCli.SetFocus
        ValidarCompostura = False
        Exit Function
    End If
    If txtCodAparato.Text = "" Then
        MsgBox "Debe ingresar un Aparato", vbExclamation, TIT_MSGBOX
        txtCodAparato.SetFocus
        ValidarCompostura = False
        Exit Function
    End If
    If chkOrden.Value = Checked Then
        If fechaOrden.Text = "" Then
            MsgBox "Debe ingresar fecha de la Orden de Arreglo", vbExclamation, TIT_MSGBOX
            fechaOrden.SetFocus
            ValidarCompostura = False
            Exit Function
        End If
    End If
        
    
    ValidarCompostura = True
End Function

Private Sub Check1_Click()

End Sub

Private Sub chkEntregado_Click()
    If chkArreglado.Value = Unchecked Then
        MsgBox "Imposible tildar como entregado si no se ha tildado la opcion arreglado.", vbInformation, TIT_MSGBOX
        chkArreglado.Value = Unchecked
    End If
    If chkArreglado.Value = Checked Then
        If !file("G:\archivos\compost.dbf") Then  ', PREGUNTA SI EXISTE, LA COMPOSTURA EN ESE DISCO
            chkArreglado.Value = Unchecked
        Else
            FecEntrega.Value = Date
            chkEntregadoSArreglar.Enabled = False
        End If
    Else
        FecEntrega.Value = ""
        chkEntregadoSArreglar.Enabled = True
    End If
End Sub

Private Sub chkEntregadoSArreglar_Click()
    If chkEntregado.Value = Checked Then
        MsgBox "Imposible tildar como entregado sin arreglar si esta tildado la opcion entregado", vbInformation, TIT_MSGBOX
        chkEntregadoSArreglar.Value = Unchecked
    End If
    If chkArreglado.Value = Checked Then
        MsgBox "Imposible tildar como entregado sin arreglar si esta tildado la opcion arreglado", vbInformation, TIT_MSGBOX
        chkEntregadoSArreglar.Value = Unchecked
    End If
    If chkEntregadoSArreglar.Value = Checked Then
        FecEntrega.Text = Date
        chkEntregado.Enabled = Unchecked
    Else
        FecEntrega.Text = Date
        chkEntregado.Enabled = Checked
    End If
End Sub

Private Sub chkNoArreglado_Click()
    If chkArreglado.Value = Checked Then
        MsgBox "Imposible tildar como no arreglado si ha tildado la opcion arreglado.", vbInformation
        chkNoArreglado.Value = Unchecked
    Else
        If chkNoArreglado.Value = Checked Then
            FecArreglo.Text = Date
        Else
            FecArreglo.Text = ""
        End If
    End If
End Sub

Private Sub cmdbuscaComp_Click()
    If txtcodCli.Text <> "" Then
        tabDatos.Tab = 1
        txtBuscaCliente.Text = txtcodCli.Text
        txtBuscarCliDescri.Text = txtRazSoc.Text
        CmdBuscAprox_Click
    End If
End Sub

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    lblEstado.Caption = "Buscando..."
    Screen.MousePointer = vbHourglass
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT CO.CO_NUMERO, CO.CO_FECHA,"
    sql = sql & " C.CLI_RAZSOC, A.APT_DESCRI, D.DES_DESCRI, V.VEN_NOMBRE, E.EST_DESCRI,CO.CO_TOTAL "
    sql = sql & " FROM COMPOSTURAS CO, CLIENTE C,  APARATO A,DESTINOS D,VENDEDOR V,ESTADO_DOCUMENTO E"
    sql = sql & " WHERE CO.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & "   AND CO.APT_CODIGO=A.APT_CODIGO"
    sql = sql & "   AND CO.DES_CODIGO=D.DES_CODIGO"
    sql = sql & "   AND CO.VEN_CODIGO=V.VEN_CODIGO"
    sql = sql & "   AND CO.EST_CODIGO=E.EST_CODIGO"
    If txtBuscaCliente.Text <> "" Then sql = sql & " AND CO.CLI_CODIGO=" & XN(txtBuscaCliente.Text)
    If txtBusCodApa.Text <> "" Then sql = sql & " AND CO.APT_CODIGO=" & XN(txtBusCodApa.Text)
    If FechaDesde.Text <> "" Then sql = sql & " AND CO.CO_FECHA>=" & XDQ(FechaDesde.Text)
    If FechaHasta.Text <> "" Then sql = sql & " AND CO.CO_FECHA<=" & XDQ(FechaHasta.Text)
    
    'If cboRecibo1.List(cboRecibo1.ListIndex) <> "(Todos)" Then sql = sql & " AND RC.TCO_CODIGO=" & XN(cboRecibo1.ItemData(cboRecibo1.ListIndex))
    sql = sql & " ORDER BY CO.CO_NUMERO, CO.CO_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!CO_NUMERO & Chr(9) & rec!CO_FECHA _
                               & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!APT_DESCRI _
                               & Chr(9) & rec!DES_DESCRI & Chr(9) & rec!VEN_NOMBRE _
                               & Chr(9) & rec!EST_DESCRI & Chr(9) & Valido_Importe(rec!CO_TOTAL)
            rec.MoveNext
        Loop
        GrdModulos.HighLight = flexHighlightAlways
        GrdModulos.SetFocus
        GrdModulos.Col = 0
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron datos... ", vbExclamation, TIT_MSGBOX
        'txtCliente.SetFocus
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdGrabar_Click()
'reglas de negocio
If chkEntregado.Value = Checked Then
    If chkSena.Value = Unchecked And chkTotal.Value = Unchecked Then
        MsgBox "Imposible facturar esta compostura, parametro erroneo", 48, "Error"
        Exit Sub
    End If
End If
If chkArreglado.Value = Unchecked Then
    If MsgBox("Cuidado la compostura no ha sido arreglada, solo se facturará la seña, ¿Desea facturar la seña?", vbYesNo, TIT_MSGBOX) = 6 Then
        MsgBox "Se ha facturado la seña"
    Else
        MsgBox "La Compostura se ha actualizado"
        GrabarCompostura
    End If
End If

If chkTotal.Value = Checked Then
    MsgBox "Imposible facturar esta compostura, ¡Ya fue facturada!", vbInformation, TIT_MSGBOX
End If

If txtSena.Text <> "" And chkSena.Value = Checked Then
    MsgBox "Factura Saldo", vbInformation, TIT_MSGBOX
     'Esto es cuando indica si factura la seña, saldo o el total
'    thisform.container3.optiongroup2.option1.enabled = .f.
'    thisform.container3.optiongroup2.option2.enabled = .t.
'    thisform.container3.optiongroup2.option3.enabled = .f.
Else
    If txtSena.Text <> "" And chkSena.Value = Unchecked Then
        MsgBox "Factura Seña / Total", vbInformation, TIT_MSGBOX
    'Esto es cuando indica si factura la seña, saldo o el total
    '    thisform.container3.optiongroup2.option1.enabled = .f.
    '    thisform.container3.optiongroup2.option2.enabled = .t.
    '    thisform.container3.optiongroup2.option3.enabled = .f.
    Else
        MsgBox "Factura Total", vbInformation, TIT_MSGBOX
    End If
End If

 
End Sub
Private Function GrabarCompostura()
'Grabar Compostura, armar funcion
 
 On Error GoTo HayErrorCompostura
    
    If MsgBox("¿Graba la Compostura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Function
    If ValidarCompostura = False Then Exit Function
    
    DBConn.BeginTrans
    sql = "SELECT * FROM COMPOSTURAS"
    sql = sql & " WHERE CO_NUMERO=" & XN(txtNroCompostura)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            
'    MsgBox "Codigo" & rec!co_numero
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then 'NUEVA COMPOSTURA
        sql = "INSERT INTO COMPOSTURAS"
        sql = sql & " (CO_NUMERO,CO_FECHA,EST_CODIGO,CLI_CODIGO,APT_CODIGO,"
        sql = sql & "MAR_CODIGO,DES_CODIGO,"
        sql = sql & "VEN_CODIGO,CO_ESTACT,CO_PROSCLI,CO_PROREAL,"
        sql = sql & "CO_TOTAL,CO_SENIA,CO_SALDO,CO_ARREGLO,CO_NOARREGLO,"
        sql = sql & "CO_ENTREGO,CO_ENTREGOSIN,CO_FECARR,CO_FECENT,"
        sql = sql & "VEN_CODIGOE,CO_AVISO,CO_CONCBTE,CO_RETIRADO,CO_ORDENARR,CO_FECHAARR)"
        sql = sql & " VALUES ("
        sql = sql & XN(txtNroCompostura) & ","
        sql = sql & XDQ(FechaCompostura) & ","
        If chkOrden.Value = Checked Then 'SI TIENE ORDEN DE ARREGLO, HAY QUE CAMBIARLE EL ESTADO
            sql = sql & XN("4") & ","
        Else
            sql = sql & XN("1") & ","
        End If
        sql = sql & XN(txtcodCli) & ","
        sql = sql & XN(txtCodAparato) & ","
        sql = sql & XN(txtCodMarca) & ","
        sql = sql & cboDestino.ItemData(cboDestino.ListIndex) & ","
        sql = sql & cboVendedor.ItemData(cboVendedor.ListIndex) & ","
        sql = sql & XS(txtEstadoActual) & ","
        sql = sql & XS(txtProbCliente) & ","
        sql = sql & XS(txtProbReal) & ","
        sql = sql & XN(txtImporte) & ","
        sql = sql & XN(txtSena) & ","
        sql = sql & XN(txtSaldo) & ","
        If chkArreglado.Value = Checked Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        If chkNoArreglado.Value = Checked Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        If chkEntregado.Value = Checked Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        If chkEntregadoSArreglar.Value = Checked Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        sql = sql & XDQ(FecArreglo) & ","
        sql = sql & XDQ(FecEntrega) & ","
        sql = sql & cboEntrego.ItemData(cboEntrego.ListIndex) & ","
        sql = sql & XS(txtAviso) & ","
        If OptCComp.Value = True Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        sql = sql & XS(txtRetiro.Text) & ","
        If chkOrden.Value = Checked Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        sql = sql & XDQ(fechaOrden) & ")"
        
        DBConn.Execute sql
    Else 'Modificar Compostura
        
        sql = "UPDATE COMPOSTURAS SET "
        sql = sql & " CO_NUMERO =" & XN(txtNroCompostura)
        sql = sql & " ,CO_FECHA =" & XDQ(FechaCompostura)
        
        If chkOrden.Value = Checked Then 'SI TIENE ORDEN DE ARREGLO, HAY QUE CAMBIARLE EL ESTADO
            sql = sql & " ,EST_CODIGO = 4"
        Else
            sql = sql & " ,EST_CODIGO = 1"
        End If
        sql = sql & " ,CLI_CODIGO =" & XN(txtcodCli)
        sql = sql & " ,APT_CODIGO =" & XN(txtCodAparato)
        sql = sql & " ,MAR_CODIGO =" & XN(txtCodMarca)
        sql = sql & " ,DES_CODIGO =" & cboDestino.ItemData(cboDestino.ListIndex)
        sql = sql & " ,VEN_CODIGO =" & cboVendedor.ItemData(cboVendedor.ListIndex)
        sql = sql & " ,CO_ESTACT =" & XS(txtEstadoActual)
        sql = sql & " ,CO_PROSCLI =" & XS(txtProbCliente)
        sql = sql & " ,CO_PROREAL =" & XS(txtProbReal)
        sql = sql & " ,CO_TOTAL =" & XN(txtImporte)
        sql = sql & " ,CO_SENIA =" & XN(txtSena)
        sql = sql & ",CO_SALDO =" & XN(txtSaldo)
        If chkArreglado.Value = Checked Then
            sql = sql & " ,CO_ARREGLO = 1"
        Else
            sql = sql & " ,CO_ARREGLO = 0"
        End If
        If chkNoArreglado.Value = Checked Then
            sql = sql & " ,CO_NOARREGLO = 1"
        Else
            sql = sql & " ,CO_NOARREGLO = 0"
        End If
        If chkEntregadoSArreglar.Value = Checked Then
            sql = sql & " ,CO_ENTREGOSIN = 1"
        Else
            sql = sql & " ,CO_ENTREGOSIN = 0"
        End If
        sql = sql & " ,CO_FECARR =" & XDQ(FecArreglo)
        sql = sql & " ,CO_FECENT    = " & XDQ(FecEntrega)
        sql = sql & " ,VEN_CODIGOE =" & cboEntrego.ItemData(cboEntrego.ListIndex)
        sql = sql & " ,CO_AVISO =" & XS(txtAviso)
        If OptCComp.Value = True Then
            sql = sql & " ,CO_CONCBTE = 1"
        Else
            sql = sql & " ,CO_CONCBTE = 0"
        End If
        sql = sql & " ,CO_RETIRADO =" & XS(txtRetiro.Text)
        If chkOrden.Value = Checked Then
            sql = sql & " ,CO_ORDENARR = 1"
        Else
            sql = sql & " ,CO_ORDENARR = 0"
        End If
        sql = sql & " ,CO_FECHAARR =" & XDQ(fechaOrden)
        DBConn.Execute sql
    End If
        
    rec.Close
    DBConn.CommitTrans
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    
    'CmdNuevo_Click
    Exit Function
    
HayErrorCompostura:
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Function

Private Sub cmdImprimir_Click()
    
    If MsgBox("¿Confirma la Impresion del sobre?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptNormal
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    
    If txtNroCompostura.Text <> "" Then
        Rep.SelectionFormula = "{COMPOSTURAS.CO_NUMERO}=" & XN(txtNroCompostura.Text)
    Else
        Exit Sub
    End If
    
    'Rep.WindowTitle = "Lista de Precios..."
       
    Rep.ReportFileName = DRIVE & DirReport & "rptCompostura.rpt"
    
    Rep.Destination = crptToPrinter
    Rep.Action = 1
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    lblEstado.Caption = ""
End Sub

Private Sub CmdNuevo_Click()
    txtcodCli.Text = ""
    txtCodCli_LostFocus
    txtCodAparato.Text = ""
    txtCodAparato_LostFocus
    txtCodMarca.Text = ""
    txtCodMarca_LostFocus
    cboDestino.ListIndex = 0
    txtEstadoActual.Text = ""
    txtProbCliente.Text = ""
    txtProbReal.Text = ""
    txtImporte.Text = ""
    chkTotal.Value = Unchecked
    txtSena.Text = ""
    chkSena.Value = Unchecked
    txtSaldo.Text = ""
    chkSaldo.Value = Unchecked
    chkArreglado.Value = Unchecked
    chkNoArreglado.Value = Unchecked
    chkEntregado.Value = Unchecked
    chkEntregadoSArreglar.Value = Unchecked
    FecArreglo.Text = ""
    FecEntrega.Text = ""
    txtRetiro.Text = ""
    OptCComp.Value = True
    cboEntrego.ListIndex = 0
    NroNuevaCompostura
    cmdbuscaComp.Enabled = False
    chkOrden.Value = Unchecked
    
    
End Sub

Private Sub CmdSalir_Click()
    
    If chkEntregado.Value = Checked Then
        If (chkSena.Value = Unchecked) And (chkTotal.Value = Unchecked) And (chkSaldo.Value = Unchecked) Then
            'If MsgBox("¿Graba la Compostura?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Function
            If MsgBox("Existe Disco secundario?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
                ' Aca hace un insert en la tabla Factura de la BD Negro
                ' Aca hace un insert en la tabla Compostura de la BD Negro
                ' y el update si fuera necesario
            End If
        End If
    Else
        GrabarCompostura
    End If
    'If ValidarCompostura = True Then
        If MsgBox("¿Confirma Salir?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
            Set frmComposturas = Nothing
            Unload Me
    'End If
' ESTE ES EL CODIGO DEL SISTEMA EN VFOX
'    oBarra.grabar.click()
'do case
'case thisform.chkentregado.value = .t.
'    if thisform.chkestseña.value = .f. and ;
'        thisform.chkestpresup.value = .f. and;
'        thisform.chkestsaldo.value = .f.
'        if file("G:\archivos\compost.dbf")
'            local quenumero
'            quenumero = THISFORM.txtNumero.Value
'            SELECT MOVIMIENTO
'            locate for movimiento.numero = quenumero
'            INSERT INTO G:\ARCHIVOS\MOVIMIE.DBF (DNI,FECHA,NUMAPA,APARATO,MARCA,ESTADO,PROBLEMAC,PROBLEMAD,PRESUPUEST,ESTPRESUP,SEÑA,ESTSEÑA,SALDO,ESTSALDO,DESTINO,ORDENARREG,FECHAORDEN,ARREGLADO,NOARREGLAD,FEC_ARRE,ENTREGADO,ENTREGADO_,FECHAENTRE,RETIRO,ESTADO1);
'            VALUES (MOVIMIENTO.DNI,MOVIMIENTO.FECHA,MOVIMIENTO.NUMAPA,MOVIMIENTO.APARATO,MOVIMIENTO.MARCA,MOVIMIENTO.ESTADO,MOVIMIENTO.PROBLEMAC,MOVIMIENTO.PROBLEMAD,MOVIMIENTO.PRESUPUESTO,MOVIMIENTO.ESTPRESUP,MOVIMIENTO.SEÑA,MOVIMIENTO.ESTSEÑA,MOVIMIENTO.SALDO,MOVIMIENTO.ESTSALDO,MOVIMIENTO.DESTINO,MOVIMIENTO.ORDENARREGLO,MOVIMIENTO.FECHAORDEN,MOVIMIENTO.ARREGLADO,MOVIMIENTO.NOARREGLADO,MOVIMIENTO.FEC_ARRE,MOVIMIENTO.ENTREGADO,MOVIMIENTO.ENTREGADO_SA,MOVIMIENTO.FECHAENTREGA,MOVIMIENTO.RETIRO,MOVIMIENTO.ESTADO1)
'            SELECT MOVIMIENTO
'            GO TOP
'            locate for movimiento.numero = quenumero
'            IF FOUND()
'                IF LOCK()
'                    =RLOCK()
'                    Delete
'                End If
'            End If
'            SELECT composturas
'            LOCATE FOR COMPOSTURAS.NUMERO = QUENUMERO
'            INSERT INTO G:\ARCHIVOS\COMPOST.DBF (FECHA,CLIENTE,APELLIDO,NOMBRE,DNI,DIRECCION,TE,CUIT,IVA,ESTADO);
'            VALUES (COMPOSTURAS.FECHA,COMPOSTURAS.CLIENTE,COMPOSTURAS.APELLIDO,COMPOSTURAS.NOMBRE,COMPOSTURAS.DNI,;
'            COMPOSTURAS.DIRECCION,COMPOSTURAS.TE,COMPOSTURAS.CUIT,COMPOSTURAS.IVA,COMPOSTURAS.ESTADO)
'            SELECT COMPOSTURAS
'            LOCATE FOR COMPOSTURAS.NUMERO = QUENUMERO
'            Delete
'            USE IN MOVIMIE &&G:\ARCHIVOS\MOVIMIE.DBF
'            USE IN COMPOST &&G:\ARCHIVOS\COMPOST.DBF
'            USE IN MOVIMIENTO
'            USE IN COMPOSTURAS
'            USE MOVIMIENTO EXCLUSIVE
'            locate for movimiento.numero = quenumero
'            IF FOUND()
'                Delete
'            End If
'        End If
'    End If
'ENDCASE
'USE IN MOVIMIENTO
'USE MOVIMIENTO EXCLUSIVE
'PACK
'USE COMPOSTURAS EXCLUSIVE
'PACK
'LOCAL PAPA
'do case
'case thisform.chkentregado.value = .t.
'    if thisform.chkestseña.value = .f. and ;
'        thisform.chkestpresup.value = .f. and;
'        thisform.chkestsaldo.value = .f.
'        if file("G:\archivos\compost.dbf")
'            use G:\archivos\compost.dbf
'            GO BOTTOM
'            Skip -1
'            PAPA = PADL(ALLTRIM(Str(Val(Numero) + 1)), 5, "0")
'            Skip
'            REPLACE NUMERO WITH PAPA
'            use G:\archivos\MOVIMIE.dbf
'            GO BOTTOM
'            Skip -1
'            PAPA = PADL(ALLTRIM(Str(Val(Numero) + 1)), 5, "0")
'            Skip
'            REPLACE NUMERO WITH PAPA
'        End If
'    End If
'ENDCASE
'THISFORM.Release
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MySendKeys Chr(9)
    End If
    If KeyAscii = vbKeyEscape Then
        CmdSalir_Click
    End If
End Sub

Private Sub Form_Load()
    Centrar_pantalla Me
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    FechaCompostura.Text = Date
    tabDatos.Tab = 0
    
    CargoComboBox cboDestino, "DESTINOS", "DES_CODIGO", "DES_DESCRI", "DES_DESCRI"
    If cboDestino.ListCount > 0 Then cboDestino.ListIndex = 0
    
    CargoComboBox cboVendedor, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE", "VEN_NOMBRE"
    If cboVendedor.ListCount > 0 Then cboVendedor.ListIndex = 0
    
    sql = "SELECT VEN_CODIGO, VEN_NOMBRE"
    sql = sql & " FROM VENDEDOR"
    sql = sql & " ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboEntrego.AddItem ""
        Do While rec.EOF = False
            cboEntrego.AddItem Trim(rec!VEN_NOMBRE)
            cboEntrego.ItemData(cboEntrego.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
    End If
    rec.Close
    If cboEntrego.ListCount > 0 Then cboEntrego.ListIndex = 0
    
    Call BuscoEstado(1, lblEstadoCompostura) 'ESTADO PENDIENTE
    
    NroNuevaCompostura
    configurogrillas
    
End Sub
Private Function NroNuevaCompostura()
    'Proximo nro de Compostura
    FechaCompostura.Text = Date
    txtNroCompostura.Text = "1"
    sql = "SELECT MAX(CO_NUMERO) AS MAYOR FROM COMPOSTURAS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtNroCompostura.Text = rec!MAYOR + 1
        lblEstado.Caption = "Ingresando Nueva Compostura...."
    End If
    rec.Close
    'Aviso o Promocion
    sql = "SELECT AVISO FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtAviso.Text = rec!AVISO
    End If
    rec.Close
End Function

Private Function configurogrillas()
'GRILLA BUSQUEDA
    GrdModulos.FormatString = "^Nro Comp|^Fecha|^Cliente|^Aparato|^Destino|^Recibio|^Estado|^Total"
    GrdModulos.ColWidth(0) = 1000 'Nro Comp
    GrdModulos.ColWidth(1) = 1200 'Fecha
    GrdModulos.ColWidth(2) = 1600 'Cliente
    GrdModulos.ColWidth(3) = 1600 'Aparato
    GrdModulos.ColWidth(4) = 1600 'Destino
    GrdModulos.ColWidth(5) = 1600 'Recibio
    GrdModulos.ColWidth(6) = 1600 'Estado
    GrdModulos.ColWidth(7) = 1200 'Total
    GrdModulos.Cols = 8
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
End Function
Private Sub GrdModulos_dblClick()
tabDatos.Tab = 0
lblEstado.Caption = "Actualizando la Compostura...."
sql = "SELECT * FROM COMPOSTURAS "
',CLIENTES CL, MARCAS M, APARATO A,DESTINOS D,VENDEDOR V"
sql = sql & "WHERE "
sql = sql & "CO_NUMERO = " & GrdModulos.TextMatrix(GrdModulos.RowSel, 0)
'sql = sql & "C.CLI_CODIGO = CL.CLI_CODIGO AND"
'sql = sql & "C.MAR_CODIGO = M.MAR_CODIGO AND"
'sql = sql & "C.APT_CODIGO = A.APT_CODIGO AND"
'sql = sql & "C.DES_CODIGO = D.DES_CODIGO AND"
'sql = sql & "C.DES_CODIGO = D.DES_CODIGO AND"
Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
If Rec1.EOF = False Then
    txtNroCompostura.Text = Rec1!CO_NUMERO
    'ESTADO DE LA COMPOSTURA
    Call BuscoEstado(Rec1!EST_CODIGO, lblEstadoCompostura)
    txtcodCli.Text = Rec1!CLI_CODIGO
    txtCodCli_LostFocus
    txtCodAparato.Text = Rec1!APT_CODIGO
    txtCodAparato_LostFocus
    txtCodMarca.Text = IIf(IsNull(Rec1!MAR_CODIGO), "", Rec1!MAR_CODIGO)
    txtCodMarca_LostFocus
    If IsNull(Rec1!DES_CODIGO) Then
        cboDestino.ListIndex = 0
    Else
        Call BuscaCodigoProxItemData(CInt(Rec1!DES_CODIGO), cboDestino)
    End If

    txtEstadoActual.Text = IIf(IsNull(CO_ESTACT), "", CO_ESTACT)
    txtProbCliente.Text = IIf(IsNull(Rec1!CO_PROSCLI), "", Rec1!CO_PROSCLI)
    txtProbReal.Text = IIf(IsNull(Rec1!CO_PROREAL), "", Rec1!CO_PROSCLI)
    txtImporte.Text = IIf(IsNull(Rec1!CO_TOTAL), "", Rec1!CO_TOTAL)
    If txtImporte.Text <> "" Then
        chkTotal.Value = Checked
    Else
        chkTotal.Value = Unchecked
    End If
    txtSena.Text = IIf(IsNull(Rec1!CO_SENIA), "", Rec1!CO_SENIA)
    If txtSena.Text <> "" Then
        chkSena.Value = Checked
    Else
        chkSena.Value = Unchecked
    End If
    txtSaldo.Text = IIf(IsNull(Rec1!CO_SALDO), "", Rec1!CO_SALDO)
    If txtSaldo.Text <> "" Then
        chkSena.Value = Checked
    Else
        chkSena.Value = Unchecked
    End If
    
    chkArreglado.Value = IIf(Rec1!CO_ARREGLO = 1, Checked, Unchecked)
    chkNoArreglado.Value = IIf(Rec1!CO_NOARREGLO = 1, Checked, Unchecked)
    chkEntregado.Value = IIf(Rec1!CO_ENTREGO = 1, Checked, Unchecked)
    chkEntregadoSArreglar.Value = IIf(Rec1!CO_ENTREGOSIN = 1, Checked, Unchecked)
    FecArreglo.Text = IIf(IsNull(Rec1!CO_FECARR), "", Rec1!CO_FECARR)
    FecEntrega.Text = IIf(IsNull(Rec1!CO_FECENT), "", Rec1!CO_FECENT)
    txtRetiro.Text = IIf(IsNull(Rec1!CO_RETIRADO), "", Rec1!CO_RETIRADO)
    If Rec1!CO_CONCBTE = 1 Then
        OptCComp.Value = True
    Else
        OptSComp.Value = True
    End If
    
    If IsNull(Rec1!VEN_CODIGOE) Then
        cboEntrego.ListIndex = 0
    Else
        Call BuscaCodigoProxItemData(CInt(Rec1!VEN_CODIGOE), cboEntrego)
    End If
    chkOrden.Value = IIf(Rec1!CO_ORDENARR = 1, Checked, Unchecked)
    fechaOrden.Text = IIf(IsNull(Rec1!CO_FECHAARR), "", Rec1!CO_FECHAARR)
    
'    cboEntrego.ListIndex = IIf(IsNull(Rec1!CO_RETIRADO), "", Rec1!CO_RETIRADO)
      
End If
Rec1.Close

End Sub

Private Sub txtBusCodApa_Change()
    If txtCodAparato.Text = "" Then
        txtBusDesApa.Text = ""
    End If
End Sub

Private Sub txtBusCodApa_GotFocus()
    SelecTexto txtCodAparato
End Sub

Private Sub txtBusCodApa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarAparato "txtBusCodApa", "CODIGO"
    End If
End Sub

Private Sub txtBusCodApa_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtBusCodApa_LostFocus()
    If txtBusCodApa.Text <> "" Then
        sql = "SELECT APT_CODIGO,APT_DESCRI"
        sql = sql & " FROM APARATO"
        sql = sql & " WHERE APT_CODIGO =" & XN(txtBusCodApa.Text)
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtBusDesApa.Text = ChkNull(rec!APT_DESCRI)
        Else
            MsgBox "El Código no existe", vbInformation
            txtBusDesApa.Text = ""
            txtBusCodApa.Text = ""
            txtBusCodApa.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtCodAparato_Change()
    If txtCodAparato.Text = "" Then
        txtDesAparato.Text = ""
    End If
End Sub

Private Sub txtCodAparato_GotFocus()
    SelecTexto txtCodAparato
End Sub

Private Sub txtCodAparato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarAparato "txtCodAparato", "CODIGO"
    End If
End Sub

Private Sub txtCodAparato_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodAparato_LostFocus()
    If txtCodAparato.Text <> "" Then
        sql = "SELECT APT_CODIGO,APT_DESCRI"
        sql = sql & " FROM APARATO"
        sql = sql & " WHERE APT_CODIGO =" & XN(txtCodAparato.Text)
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesAparato.Text = ChkNull(rec!APT_DESCRI)
        Else
            MsgBox "El Código no existe", vbInformation
            txtDesAparato.Text = ""
            txtCodAparato.Text = ""
            txtCodAparato.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtcodCli_Change()
    If txtcodCli.Text = "" Then
        txtRazSoc.Text = ""
        txtDomici.Text = ""
        txtCuit.Text = ""
        txtCiva.Text = ""
    End If
End Sub

Private Sub txtcodCli_GotFocus()
    SelecTexto txtcodCli
End Sub

Private Sub txtcodCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtcodCli", "CODIGO"
    End If
End Sub

Private Sub txtcodCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodCli_LostFocus()
    If txtcodCli.Text <> "" Then
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI,I.IVA_CODIGO,I.IVA_DESCRI,"
        sql = sql & "C.CLI_TELEFONO,C.CLI_CUIT,C.CLI_INGBRU"
        sql = sql & " FROM CLIENTE C, CONDICION_IVA I"
        sql = sql & " WHERE I.IVA_CODIGO = C.IVA_CODIGO"
        sql = sql & " AND CLI_CODIGO =" & XN(txtcodCli.Text)
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtRazSoc.Text = ChkNull(rec!CLI_RAZSOC)
            txtDomici.Text = ChkNull(rec!CLI_DOMICI)
            txtCiva.Text = ChkNull(rec!IVA_DESCRI)
            txtCuit.Text = ChkNull(rec!CLI_CUIT)
            cmdbuscaComp.Enabled = True
        Else
            MsgBox "El Código no existe", vbInformation
            txtRazSoc.Text = ""
            txtcodCli.Text = ""
            txtcodCli.SetFocus
            cmdbuscaComp.Enabled = False
        End If
        If rec.State = 1 Then rec.Close
    Else
        cmdbuscaComp.Enabled = False
    End If
End Sub

Private Sub txtCodMarca_Change()
    If txtCodMarca.Text = "" Then
        txtDesMarca.Text = ""
    End If
End Sub

Private Sub txtCodMarca_GotFocus()
    SelecTexto txtCodMarca
End Sub

Private Sub txtCodMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarMarca "txtCodMarca", "CODIGO"
    End If
End Sub

Private Sub txtCodMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCodMarca_LostFocus()
    If txtCodMarca.Text <> "" Then
        sql = "SELECT MAR_CODIGO,MAR_DESCRI"
        sql = sql & " FROM MARCAS"
        sql = sql & " WHERE MAR_CODIGO =" & XN(txtCodMarca.Text)
        If rec.State = 1 Then rec.Close
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesMarca.Text = ChkNull(rec!MAR_DESCRI)
        Else
            MsgBox "El Código no existe", vbInformation
            txtDesMarca.Text = ""
            txtCodMarca.Text = ""
            txtCodMarca.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDesAparato_Change()
    If txtDesAparato.Text = "" Then
        txtCodAparato.Text = ""
    End If
End Sub

Private Sub txtDesAparato_GotFocus()
    SelecTexto txtDesAparato
End Sub

Private Sub txtDesAparato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarAparato "txtCodAparato", "CODIGO"
    End If
End Sub

Private Sub txtDesAparato_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesAparato_LostFocus()
    If txtCodAparato.Text = "" And txtDesAparato.Text <> "" Then
        sql = "SELECT APT_CODIGO,APT_DESCRI"
        sql = sql & " FROM APARATO"
        sql = sql & " WHERE APT_DESCRI LIKE '" & XN(Trim(txtDesAparato.Text)) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarAparato "txtCodAparato", "CADENA", Trim(txtDesAparato.Text)
                If rec.State = 1 Then rec.Close
                txtDesAparato.SetFocus
            Else
                txtCodAparato.Text = rec!APT_CODIGO
                txtDesAparato.Text = rec!APT_DESCRI
            End If
        Else
            lblEstado.Caption = ""
            MsgBox "El Aparato no existe", vbExclamation, TIT_MSGBOX
            txtCodAparato.Text = ""
            txtDesAparato.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDesMarca_Change()
    If txtDesMarca.Text = "" Then
        txtCodMarca.Text = ""
    End If
End Sub

Private Sub txtDesMarca_GotFocus()
    SelecTexto txtDesMarca
End Sub

Private Sub txtDesMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarMarca "txtCodMarca", "CODIGO"
    End If
End Sub

Private Sub txtDesMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesMarca_LostFocus()
    If txtCodMarca.Text = "" And txtDesMarca.Text <> "" Then
        sql = "SELECT MAR_CODIGO,MAR_DESCRI"
        sql = sql & " FROM MARCAS"
        sql = sql & " WHERE MAR_DESCRI LIKE '" & XN(Trim(txtDesMarca.Text)) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarMarca "txtCodMarca", "CADENA", Trim(txtDesMarca.Text)
                If rec.State = 1 Then rec.Close
                txtDesMarca.SetFocus
            Else
                txtCodMarca.Text = rec!MAR_CODIGO
                txtDesMarca.Text = rec!MAR_DESCRI
            End If
        Else
            lblEstado.Caption = ""
            MsgBox "La Marca no existe", vbExclamation, TIT_MSGBOX
            txtCodMarca.Text = ""
            txtDesMarca.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtImporte_GotFocus()
    SelecTexto txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtImporte, KeyAscii)
End Sub

Private Sub txtImporte_LostFocus()
    If txtImporte.Text <> "" Then
        txtImporte.Text = Valido_Importe(txtImporte.Text)
        'chkTotal.Value = Checked
    Else
        txtImporte.Text = "0,00"
        'chkTotal.Value = Unchecked
    End If
    txtSaldo.Text = CDbl(txtImporte.Text) - CDbl(Chk0(txtSena.Text))
    txtSaldo.Text = Valido_Importe(txtSaldo.Text)
End Sub

Private Sub txtRazSoc_Change()
    If txtRazSoc.Text = "" Then
        txtcodCli.Text = ""
        txtDomici.Text = ""
        txtCuit.Text = ""
        txtCiva.Text = ""
    End If
End Sub
Private Sub txtRazSoc_GotFocus()
    SelecTexto txtRazSoc
End Sub

Private Sub txtRazSoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtcodCli", "CODIGO"
    End If
End Sub

Private Sub txtRazSoc_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtRazSoc_LostFocus()
    If txtcodCli.Text = "" And txtRazSoc.Text <> "" Then
        sql = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC,C.CLI_DOMICI,I.IVA_DESCRI, C.CLI_CUIT"
        sql = sql & " FROM CLIENTE C, CONDICION_IVA I"
        sql = sql & " WHERE I.IVA_CODIGO = C.IVA_CODIGO"
        sql = sql & " AND CLI_RAZSOC LIKE '" & XN(Trim(txtRazSoc.Text)) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtcodCli", "CADENA", Trim(txtRazSoc.Text)
                If rec.State = 1 Then rec.Close
                txtRazSoc.SetFocus
            Else
                txtcodCli.Text = rec!CLI_CODIGO
                txtRazSoc.Text = rec!CLI_RAZSOC
                txtDomici.Text = ChkNull(rec!CLI_DOMICI)
                txtCiva.Text = ChkNull(rec!IVA_DESCRI)
                txtCuit.Text = ChkNull(rec!CLI_CUIT)
            End If
        Else
            lblEstado.Caption = ""
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtcodCli.Text = ""
            txtRazSoc.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtRetiro_GotFocus()
    SelecTexto txtRetiro
End Sub

Private Sub txtRetiro_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtSena_GotFocus()
    SelecTexto txtSena
End Sub

Private Sub txtSena_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtSena, KeyAscii)
End Sub

Private Sub txtSena_LostFocus()
    If txtSena.Text <> "" Then
        txtSena.Text = Valido_Importe(txtSena.Text)
    Else
        txtSena.Text = "0,00"
    End If
    If CDbl(txtSena.Text) > CDbl(txtImporte.Text) Then
        MsgBox "La Seña no puede ser Mayor al Importe de la Compostura", vbExclamation, TIT_MSGBOX
        txtSena.SetFocus
        Exit Sub
    End If
    txtSaldo.Text = CDbl(txtImporte.Text) - CDbl(Chk0(txtSena.Text))
    txtSaldo.Text = Valido_Importe(txtSaldo.Text)
End Sub

Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtcodCli" Then
                txtcodCli.Text = .ResultFields(2)
                txtCodCli_LostFocus
'            Else
'                txtBuscaCliente.Text = .ResultFields(2)
'                txtBuscaCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Public Sub BuscarAparato(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT APT_DESCRI, APT_CODIGO"
        cSQL = cSQL & " FROM APARATO"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE APT_DESCRI LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Descripción, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "APT_DESCRI"
        campo1 = .Field
        .Field = "APT_CODIGO"
        campo2 = .Field
        .OrderBy = "APT_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Aparatos :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCodAparato.Text = .ResultFields(2)
            txtCodAparato_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

Public Sub BuscarMarca(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT MAR_DESCRI, MAR_CODIGO"
        cSQL = cSQL & " FROM MARCAS"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE MAR_DESCRI LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Descripción, Código"
        .sql = cSQL
        .Headers = hSQL
        .Field = "MAR_DESCRI"
        campo1 = .Field
        .Field = "MAR_CODIGO"
        campo2 = .Field
        .OrderBy = "MAR_DESCRI"
        camponumerico = False
        .Titulo = "Busqueda de Marcas :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCodMarca.Text = .ResultFields(2)
            txtCodMarca_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub txtBusDesApa_Change()
    If txtBusDesApa.Text = "" Then
        txtBusCodApa.Text = ""
    End If
End Sub

Private Sub txtBusDesApa_GotFocus()
    SelecTexto txtBusDesApa
End Sub

Private Sub txtBusDesApa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarAparato "txtBusCodApa", "CODIGO"
    End If
End Sub

Private Sub txtBusDesApa_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtBusDesApa_LostFocus()
    If txtBusCodApa.Text = "" And txtBusDesApa.Text <> "" Then
        sql = "SELECT APT_CODIGO,APT_DESCRI"
        sql = sql & " FROM APARATO"
        sql = sql & " WHERE APT_DESCRI LIKE '" & XN(Trim(txtBusDesApa.Text)) & "%'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarAparato "txtBusCodApa", "CADENA", Trim(txtBusDesApa.Text)
                If rec.State = 1 Then rec.Close
                txtBusDesApa.SetFocus
            Else
                txtBusCodApa.Text = rec!APT_CODIGO
                txtBusDesApa.Text = rec!APT_DESCRI
            End If
        Else
            lblEstado.Caption = ""
            MsgBox "El Aparato no existe", vbExclamation, TIT_MSGBOX
            txtBusCodApa.Text = ""
            txtBusDesApa.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub
