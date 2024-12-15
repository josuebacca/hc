VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form frmNotaDeditoClienteCheques 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Dédito Clientes por Cheques..."
   ClientHeight    =   7665
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
   ScaleHeight     =   7665
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   450
      Left            =   8535
      TabIndex        =   15
      Top             =   7185
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   450
      Left            =   10305
      TabIndex        =   17
      Top             =   7185
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   7650
      TabIndex        =   14
      Top             =   7185
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   450
      Left            =   9420
      TabIndex        =   16
      Top             =   7185
      Width           =   870
   End
   Begin TabDlg.SSTab tabDatos 
      Height          =   7125
      Left            =   60
      TabIndex        =   27
      Top             =   15
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   12568
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
      TabPicture(0)   =   "frmNotaDeditoClienteCheques.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameNotaDebito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FramePara"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Buscar"
      TabPicture(1)   =   "frmNotaDeditoClienteCheques.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameBuscar"
      Tab(1).Control(1)=   "GrdModulos"
      Tab(1).ControlCount=   2
      Begin VB.Frame FramePara 
         Caption         =   "Para..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   4455
         TabIndex        =   47
         Top             =   375
         Width           =   6585
         Begin VB.ComboBox CboVend 
            Height          =   315
            Left            =   870
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1530
            Width           =   3495
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
            Height          =   315
            Left            =   1890
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Buscar Cliente"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   405
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
            Left            =   870
            MaxLength       =   50
            TabIndex        =   67
            Top             =   1095
            Width           =   5640
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
            Left            =   870
            TabIndex        =   66
            Top             =   765
            Width           =   2175
         End
         Begin VB.TextBox txtCliLocalidad 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3525
            TabIndex        =   65
            Top             =   765
            Width           =   2985
         End
         Begin VB.TextBox txtCliRazSoc 
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
            Left            =   2325
            MaxLength       =   50
            TabIndex        =   6
            Tag             =   "Descripción"
            Top             =   420
            Width           =   4185
         End
         Begin VB.TextBox txtCodCliente 
            Height          =   315
            Left            =   870
            MaxLength       =   40
            TabIndex        =   5
            Top             =   420
            Width           =   975
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
            Left            =   90
            TabIndex        =   76
            Top             =   1575
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Loc.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3120
            TabIndex        =   70
            Top             =   810
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   90
            TabIndex        =   69
            Top             =   1125
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia:"
            Height          =   195
            Left            =   90
            TabIndex        =   68
            Top             =   810
            Width           =   705
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   64
            Top             =   480
            Width           =   555
         End
      End
      Begin VB.Frame FrameNotaDebito 
         Caption         =   "Nota de Débito..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   105
         TabIndex        =   29
         Top             =   375
         Width           =   4350
         Begin VB.TextBox txtNroSucursal 
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
            Left            =   1260
            MaxLength       =   4
            TabIndex        =   2
            Top             =   975
            Width           =   555
         End
         Begin VB.TextBox txtNroNotaDebito 
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
            Left            =   1845
            MaxLength       =   8
            TabIndex        =   3
            Top             =   975
            Width           =   1065
         End
         Begin VB.ComboBox cboRep 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   285
            Width           =   3015
         End
         Begin VB.ComboBox cboNotaDebito 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   630
            Width           =   2400
         End
         Begin FechaCtl.Fecha FechaNotaDebito 
            Height          =   285
            Left            =   1260
            TabIndex        =   4
            Top             =   1335
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   73
            Top             =   315
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   135
            TabIndex        =   50
            Top             =   660
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   135
            TabIndex        =   48
            Top             =   1365
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   135
            TabIndex        =   46
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   135
            TabIndex        =   45
            Top             =   1665
            Width           =   555
         End
         Begin VB.Label lblEstadoNotaDebito 
            AutoSize        =   -1  'True
            Caption         =   "EST. NOTA DEBITO"
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
            Left            =   1260
            TabIndex        =   44
            Top             =   1680
            Width           =   1500
         End
      End
      Begin VB.Frame frameBuscar 
         Caption         =   "Buscar Nota de Dédito por..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   -74610
         TabIndex        =   32
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
            Left            =   6645
            TabIndex        =   77
            Text            =   "A"
            Top             =   1305
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.ComboBox cboBuscaRep 
            Height          =   315
            Left            =   2775
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1950
            Width           =   3090
         End
         Begin VB.CommandButton cmdBuscarVen 
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
            Left            =   3825
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Buscar Vendedor"
            Top             =   945
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.ComboBox cboNotaDebito1 
            Height          =   315
            Left            =   2775
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1605
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
            Left            =   3825
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   42
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
            Left            =   3825
            MaskColor       =   &H000000FF&
            Picture         =   "frmNotaDeditoClienteCheques.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Buscar Sucursal"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   405
         End
         Begin VB.TextBox txtVendedor 
            Height          =   300
            Left            =   2775
            TabIndex        =   20
            Top             =   945
            Width           =   990
         End
         Begin VB.TextBox txtDesVen 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4260
            TabIndex        =   39
            Top             =   960
            Width           =   4620
         End
         Begin VB.CommandButton CmdBuscAprox 
            Caption         =   "Buscar"
            Height          =   360
            Left            =   7215
            MaskColor       =   &H000000FF&
            TabIndex        =   25
            ToolTipText     =   "Buscar "
            Top             =   1905
            UseMaskColor    =   -1  'True
            Width           =   1665
         End
         Begin FechaCtl.Fecha FechaHasta 
            Height          =   285
            Left            =   5280
            TabIndex        =   22
            Top             =   1290
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Separador       =   "/"
            Text            =   ""
            MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
         End
         Begin FechaCtl.Fecha FechaDesde 
            Height          =   330
            Left            =   2775
            TabIndex        =   21
            Top             =   1290
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4260
            MaxLength       =   50
            TabIndex        =   34
            Tag             =   "Descripción"
            Top             =   255
            Width           =   4620
         End
         Begin VB.TextBox txtCliente 
            Height          =   300
            Left            =   2775
            MaxLength       =   40
            TabIndex        =   18
            Top             =   255
            Width           =   975
         End
         Begin VB.TextBox txtSucursal 
            Height          =   300
            Left            =   2775
            MaxLength       =   40
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtDesSuc 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4260
            MaxLength       =   50
            TabIndex        =   33
            Tag             =   "Descripción"
            Top             =   600
            Width           =   4620
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Representada:"
            Height          =   195
            Left            =   1635
            TabIndex        =   74
            Top             =   1980
            Width           =   1080
         End
         Begin VB.Label lbltipoFac 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   1635
            TabIndex        =   63
            Top             =   1650
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   0
            Left            =   1635
            TabIndex        =   40
            Top             =   975
            Width           =   750
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Left            =   4230
            TabIndex        =   38
            Top             =   1335
            Width           =   960
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   1635
            TabIndex        =   37
            Top             =   1320
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
            Left            =   1635
            TabIndex        =   36
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
            Left            =   1635
            TabIndex        =   35
            Top             =   645
            Width           =   660
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdModulos 
         Height          =   3930
         Left            =   -74625
         TabIndex        =   26
         Top             =   2985
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6932
         _Version        =   393216
         Cols            =   13
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
         Height          =   4800
         Left            =   105
         TabIndex        =   30
         Top             =   2235
         Width           =   10950
         Begin VB.CheckBox chkBonificaEnPesos 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en $"
            Height          =   285
            Left            =   390
            TabIndex        =   10
            Top             =   4035
            Width           =   1290
         End
         Begin VB.CheckBox chkBonificaEnPorsentaje 
            Alignment       =   1  'Right Justify
            Caption         =   "Bonifica en % "
            Height          =   285
            Left            =   390
            TabIndex        =   9
            Top             =   3735
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
            TabIndex        =   61
            Top             =   4065
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
            Left            =   6900
            TabIndex        =   58
            Top             =   4065
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeIva 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6900
            TabIndex        =   12
            Top             =   3735
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
            TabIndex        =   55
            Top             =   4065
            Width           =   1155
         End
         Begin VB.TextBox txtPorcentajeBoni 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2850
            TabIndex        =   11
            Top             =   3735
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
            Left            =   8970
            TabIndex        =   52
            Top             =   4065
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
            Left            =   8970
            TabIndex        =   51
            Top             =   3735
            Width           =   1350
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Left            =   1470
            MaxLength       =   60
            TabIndex        =   13
            Top             =   4410
            Width           =   8865
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   945
            TabIndex        =   31
            Top             =   480
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSFlexGridLib.MSFlexGrid grdGrilla 
            Height          =   3495
            Left            =   210
            TabIndex        =   8
            Top             =   255
            Width           =   10365
            _ExtentX        =   18283
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   3
            Cols            =   12
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
            TabIndex        =   62
            Top             =   4125
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   6150
            TabIndex        =   60
            Top             =   4110
            Width           =   630
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "% I.V.A.:"
            Height          =   195
            Left            =   6150
            TabIndex        =   59
            Top             =   3765
            Width           =   705
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   1920
            TabIndex        =   57
            Top             =   4110
            Width           =   630
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bonificación:"
            Height          =   195
            Left            =   1920
            TabIndex        =   56
            Top             =   3765
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   54
            Top             =   4110
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub-Total:"
            Height          =   195
            Left            =   8175
            TabIndex        =   53
            Top             =   3765
            Width           =   750
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones:"
            Height          =   195
            Left            =   210
            TabIndex        =   49
            Top             =   4455
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
         TabIndex        =   28
         Top             =   570
         Width           =   1065
      End
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "<F1> Buscar Nota de Débito"
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
      Left            =   4260
      TabIndex        =   75
      Top             =   7260
      Width           =   2760
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
      TabIndex        =   43
      Top             =   7260
      Width           =   660
   End
End
Attribute VB_Name = "frmNotaDeditoClienteCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim W As Integer
Dim VBonificacion As Double
Dim VTotal As Double
Dim VEstadoNotaDebito As Integer
Dim VIva As String
Dim VIvaCalculo As Double
Dim VSucursal As String
Dim VNroNd As String
Dim VBanderaBuscar  As Boolean

Private Sub cboNotaDebito_Click()
    txtNroSucursal.Text = ""
    txtNroNotaDebito.Text = ""
End Sub

Private Sub cboRep_Click()
    txtNroSucursal.Text = ""
    txtNroNotaDebito.Text = ""
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
            
    sql = "SELECT ND.*, C.CLI_RAZSOC , TC.TCO_ABREVIA"
    sql = sql & " FROM NOTA_DEBITO_CLIENTE ND,"
    sql = sql & " TIPO_COMPROBANTE TC , CLIENTE C"
    sql = sql & " WHERE ND.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND ND.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND ND.NDC_SERVICHEQUE='C'" 'PARA QUE BUSQUE CHEQUES
    If txtCliente.Text <> "" Then sql = sql & " AND ND.CLI_CODIGO=" & XN(txtCliente)
    If FechaDesde <> "" Then sql = sql & " AND ND.NDC_FECHA>=" & XDQ(FechaDesde)
    If FechaHasta <> "" Then sql = sql & " AND ND.NDC_FECHA<=" & XDQ(FechaHasta)
    If cboNotaDebito1.List(cboNotaDebito1.ListIndex) <> "(Todas)" Then sql = sql & " AND ND.TCO_CODIGO=" & XN(cboNotaDebito1.ItemData(cboNotaDebito1.ListIndex))
    If cboBuscaRep.List(cboBuscaRep.ListIndex) <> "(Todas)" Then sql = sql & " AND ND.REP_CODIGO=" & XN(cboBuscaRep.ItemData(cboBuscaRep.ListIndex))
    sql = sql & " ORDER BY ND.NDC_SUCURSAL,ND.NDC_NUMERO"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!NDC_SUCURSAL, "0000") & "-" & Format(rec!NDC_NUMERO, "00000000") _
                            & Chr(9) & rec!NDC_FECHA & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!NDC_BONIFICA & Chr(9) & rec!NDC_IVA & Chr(9) & rec!NDC_OBSERVACION _
                            & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!NDC_BONIPESOS _
                            & Chr(9) & rec!CLI_CODIGO & Chr(9) & rec!REP_CODIGO & Chr(9) & ChkNull(rec!VEN_CODIGO)
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
        GrdModulos.HighLight = flexHighlightAlways
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
    frmBuscar.TxtDescriB = ""
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
    Dim VStockFisico As String
    
    If ValidarNotaBebito = False Then Exit Sub
    If MsgBox("¿Confirma Nota de Débito?" & Chr(13) & Chr(13) & _
            "Representada: " & cboRep.List(cboRep.ListIndex) & Chr(13) & _
            "Tipo ND:  " & cboNotaDebito.List(cboNotaDebito.ListIndex) & Chr(13) & _
            "Número:   " & txtNroSucursal.Text & "-" & txtNroNotaDebito.Text, vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo HayErrorFactura
    
    DBConn.BeginTrans
    sql = "SELECT * FROM NOTA_DEBITO_CLIENTE"
    sql = sql & " WHERE TCO_CODIGO=" & XN(cboNotaDebito.ItemData(cboNotaDebito.ListIndex))
    sql = sql & " AND NDC_NUMERO= " & XN(txtNroNotaDebito)
    sql = sql & " AND NDC_SUCURSAL=" & XN(txtNroSucursal)
    sql = sql & " AND REP_CODIGO=" & XN(cboRep.ItemData(cboRep.ListIndex))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    
    If rec.EOF = True Then
        sql = "INSERT INTO NOTA_DEBITO_CLIENTE"
        sql = sql & " (TCO_CODIGO, NDC_NUMERO, NDC_SUCURSAL, NDC_FECHA, REP_CODIGO,"
        sql = sql & " CLI_CODIGO, VEN_CODIGO, NDC_BONIFICA, NDC_IVA, NDC_SERVICHEQUE,"
        sql = sql & " NDC_OBSERVACION,NDC_NUMEROTXT,NDC_SUBTOTAL,NDC_TOTAL,NDC_BONIPESOS,"
        sql = sql & " NDC_SALDO,EST_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & XN(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)) & ","
        sql = sql & XN(txtNroNotaDebito) & ","
        sql = sql & XN(txtNroSucursal) & ","
        sql = sql & XDQ(FechaNotaDebito) & ","
        sql = sql & XN(cboRep.ItemData(cboRep.ListIndex)) & ","
        sql = sql & XN(txtCodCliente) & ","
        If CboVend.List(CboVend.ListIndex) <> "(Ninguno)" Then
            sql = sql & XN(CboVend.ItemData(CboVend.ListIndex)) & ","
        Else
            sql = sql & "NULL,"
        End If
        sql = sql & XN(txtPorcentajeBoni) & ","
        sql = sql & XN(txtPorcentajeIva.Text) & ","
        sql = sql & "'C'" & "," 'SE TRATA CHEQUES
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & XS(Format(txtNroNotaDebito.Text, "00000000")) & ","
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
        If chkBonificaEnPesos.Value = Checked Then
            sql = sql & "'S'" & "," 'BONIFICA EN PESOS
        ElseIf chkBonificaEnPorsentaje.Value = Checked Then
            sql = sql & "'N'" & "," 'BONIFICA EN PORCENTAJE
        Else
            sql = sql & "NULL" & "," 'NO HAY BONIFICACION
        End If
        sql = sql & XN(txtTotal) & "," 'SALDO NOTA DEBITO
        sql = sql & "3)" 'ESTADO DEFINITIVO
        DBConn.Execute sql
           
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                sql = "INSERT INTO DETALLE_NOTA_DEBITO_CLIENTE"
                sql = sql & " (TCO_CODIGO, NDC_NUMERO, NDC_SUCURSAL, NDC_FECHA, REP_CODIGO,"
                sql = sql & " DND_NROITEM, BAN_CODINT,"
                sql = sql & " CHE_NUMERO, DND_PRECIO, DND_BONIFICA)"
                sql = sql & " VALUES ("
                sql = sql & XN(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)) & ","
                sql = sql & XN(txtNroNotaDebito) & ","
                sql = sql & XN(txtNroSucursal) & ","
                sql = sql & XDQ(FechaNotaDebito) & ","
                sql = sql & XN(cboRep.ItemData(cboRep.ListIndex)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 11)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 10)) & ","
                sql = sql & XS(grdGrilla.TextMatrix(i, 0)) & ","
                sql = sql & XN(grdGrilla.TextMatrix(i, 6)) & "," 'IMPORTE CHEQUE
                sql = sql & XN(grdGrilla.TextMatrix(i, 7)) & ")"
                DBConn.Execute sql
            End If
        Next
        
        'ACTUALIZO LA TABLA PARAMENTROS Y LE SUMO UNO A LA NOTA DE DEBITO QUE CORRESPONDA
        Call ActualizoNumeroComprobantes(cboRep.ItemData(cboRep.ListIndex), cboNotaDebito.ItemData(cboNotaDebito.ListIndex), txtNroNotaDebito.Text)
        
'        'ACTUALIZO LA CUENTA CORRIENTE DEL CLIENTE
'        DBConn.Execute AgregoCtaCteCliente(txtCodCliente, CStr(cboNotaDebito.ItemData(cboNotaDebito.ListIndex)) _
'                                            , txtNroNotaDebito, txtNroSucursal, CStr(cboRep.ItemData(cboRep.ListIndex)), _
'                                            FechaNotaDebito, txtTotal, "D", CStr(Date))
        
        DBConn.CommitTrans
    Else
        MsgBox "La Nota de Débito ya fue Registrada", vbCritical, TIT_MSGBOX
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

Private Function ValidarNotaBebito() As Boolean
    If FechaNotaDebito.Text = "" Then
        MsgBox "La Fecha de la Nota de Débito es requerida", vbExclamation, TIT_MSGBOX
        FechaNotaDebito.SetFocus
        ValidarNotaBebito = False
        Exit Function
    End If
    If txtCodCliente.Text = "" Then
        MsgBox "El Cliente es requerido", vbExclamation, TIT_MSGBOX
        txtCodCliente.SetFocus
        ValidarNotaBebito = False
        Exit Function
    End If
    If chkBonificaEnPesos.Value = Checked Or chkBonificaEnPorsentaje.Value = Checked Then
        If txtPorcentajeBoni.Text = "" Then
            MsgBox "Debe ingresar la Bonificación", vbExclamation, TIT_MSGBOX
            txtPorcentajeBoni.SetFocus
            ValidarNotaBebito = False
            Exit Function
        End If
    End If
    ValidarNotaBebito = True
End Function

Private Sub cmdImprimir_Click()
    If MsgBox("¿Confirma Impresión Nota de Débito?" & Chr(13) & Chr(13) & _
            "Representada: " & cboRep.List(cboRep.ListIndex) & Chr(13) & _
            "Tipo ND:  " & cboNotaDebito.List(cboNotaDebito.ListIndex) & Chr(13) & _
            "Número:   " & txtNroSucursal.Text & "-" & txtNroNotaDebito.Text, vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
    ImprimirNotaDebito
End Sub

Public Sub ImprimirNotaDebito()
    Dim Renglon As Double
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Imprimiendo..."
    
    For W = 1 To 3 'SE IMPRIME POR TRIPLICADO
      '-----IMPRESION DEL ENCABEZADO------------------
        ImprimirEncabezado
        
      '---- IMPRESION DE LA NOTA DE DEBITO ------------------
        Renglon = 9.9
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 0) <> "" Then
                Printer.FontSize = 8
                Imprimir 1.1, Renglon, False, grdGrilla.TextMatrix(i, 0) 'NRO CHEQUE
                If Len(grdGrilla.TextMatrix(i, 1)) < 45 Then
                    Imprimir 2.8, Renglon, False, Trim(grdGrilla.TextMatrix(i, 1)) 'DESCRIPCION CHEQUE
                Else
                    Imprimir 2.8, Renglon, False, Trim(Left(grdGrilla.TextMatrix(i, 1), 44)) & "..." 'DESCRIPCION CHEQUE
                End If
                Printer.FontSize = 9
                Imprimir 13.5, Renglon, False, grdGrilla.TextMatrix(i, 6) 'IMPORTE CHEQUE
                Imprimir 15.5, Renglon, False, IIf(grdGrilla.TextMatrix(i, 7) = "", "0,00", grdGrilla.TextMatrix(i, 4)) 'bonoficacion
                Imprimir 17.5, Renglon, False, grdGrilla.TextMatrix(i, 9) 'importe
                Renglon = Renglon + 0.5
            End If
        Next i
        '-----OBSERVACIONES---------------------
        If txtObservaciones.Text <> "" Then
            Imprimir 1.2, Renglon + 1, False, "Observaciones: " & Trim(txtObservaciones.Text)
        End If
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
 '-----------IMPRIME EL ENCABEZADO DE LA NOTA DE DEBITO-------------------
    Printer.FontSize = 8
    Imprimir 13.4, 0.6, True, Trim(cboNotaDebito.List(cboNotaDebito.ListIndex)) & "   Nº " & Trim(txtNroSucursal.Text) & "-" & Trim(txtNroNotaDebito.Text)
    Printer.FontSize = 10
    Imprimir 15.5, 2.1, False, Format(FechaNotaDebito.Text, "dd/mm/yyyy")
    
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT C.CLI_CODIGO, C.CLI_RAZSOC, C.CLI_DOMICI, C.CLI_CUIT, C.CLI_INGBRU,"
    sql = sql & " L.LOC_DESCRI , P.PRO_DESCRI,CI.IVA_DESCRI"
    sql = sql & " FROM CLIENTE C, LOCALIDAD L,"
    sql = sql & " PROVINCIA P, CONDICION_IVA CI"
    sql = sql & " WHERE"
    sql = sql & " C.CLI_CODIGO=" & XN(txtCodCliente)
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
        Imprimir 1.3, 5.4, False, "Loc: " & Trim(Rec1!LOC_DESCRI) & " -- Prov: " & Trim(Rec1!PRO_DESCRI)
        Imprimir 1.7, 6.2, False, Trim(Rec1!IVA_DESCRI)
        Imprimir 7.9, 6.2, False, IIf(IsNull(Rec1!CLI_CUIT), "NO INFORMADO", Format(Rec1!CLI_CUIT, "##-########-#"))
        Imprimir 15.7, 6.2, False, IIf(IsNull(Rec1!CLI_INGBRU), "NO INFORMADO", Format(Rec1!CLI_INGBRU, "###-#####-##"))
    End If
    Rec1.Close
    Imprimir 4.8, 7.5, False, "CHEQUES DEVUELTOS"
    Imprimir 1.1, 9.2, False, "Nro.Cheque"
    Imprimir 2.8, 9.2, False, "Banco"
    Imprimir 13.5, 9.2, False, "Importe Che."
    Imprimir 15.5, 9.2, False, "Bonif."
    Imprimir 17.5, 9.2, False, "Importe"
End Sub

Private Sub CmdNuevo_Click()
    'significa que no estoy buacando
    VBanderaBuscar = False
   For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 0) = ""
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.TextMatrix(i, 5) = ""
        grdGrilla.TextMatrix(i, 6) = ""
        grdGrilla.TextMatrix(i, 7) = ""
        grdGrilla.TextMatrix(i, 8) = ""
        grdGrilla.TextMatrix(i, 9) = ""
        grdGrilla.TextMatrix(i, 10) = ""
        grdGrilla.TextMatrix(i, 11) = i
   Next
   txtCodCliente.Text = ""
   txtNroNotaDebito.Text = ""
   txtNroSucursal.Text = ""
   FechaNotaDebito.Text = Date
   lblEstadoNotaDebito.Caption = ""
   txtSubtotal.Text = ""
   txtTotal.Text = ""
   txtPorcentajeBoni.Text = ""
   txtPorcentajeIva.Text = ""
   txtImporteBoni.Text = ""
   txtSubTotalBoni.Text = ""
   txtImporteIva.Text = ""
   txtObservaciones.Text = ""
   lblEstado.Caption = ""
   cmdGrabar.Enabled = True
   'CARGO ESTADO
   Call BuscoEstado(1, lblEstadoNotaDebito) 'ESTADO PENDIENTE
   VEstadoNotaDebito = 1
   '--------------
   chkBonificaEnPorsentaje.Value = Unchecked
   chkBonificaEnPesos.Value = Unchecked
   FrameNotaDebito.Enabled = True
   FramePara.Enabled = True
   tabDatos.Tab = 0
   cboNotaDebito.ListIndex = 0
   CboVend.ListIndex = 0
   cboRep.ListIndex = 0
   cboRep.SetFocus
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmNotaDeditoClienteCheques = Nothing
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

    grdGrilla.FormatString = "Nro Cheque|Bco|Loc|Suc|Código|Banco|Imp Che|" _
                             & "Bonif.|Pre.Bonif.|Importe|COD INT BANCO|Orden"
    grdGrilla.ColWidth(0) = 1200 'NRO CHEQUE
    grdGrilla.ColWidth(1) = 500  'BCO
    grdGrilla.ColWidth(2) = 500  'LOC
    grdGrilla.ColWidth(3) = 500  'SUC
    grdGrilla.ColWidth(4) = 800  'CODIGO
    grdGrilla.ColWidth(5) = 2900 'BANCO
    grdGrilla.ColWidth(6) = 1000 'IMPORTE CHEQUE
    grdGrilla.ColWidth(7) = 700  'BONOFICACION
    grdGrilla.ColWidth(8) = 1000 'PRE BONIFICACION
    grdGrilla.ColWidth(9) = 1000 'IMPORTE
    grdGrilla.ColWidth(10) = 0   'CODIGO INTERNO BANCO
    grdGrilla.ColWidth(11) = 0   'ORDEN
    grdGrilla.Cols = 12
    grdGrilla.Rows = 1
    
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To 11
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    For i = 2 To 14
        grdGrilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" _
                             & Chr(9) & (i - 1)
    Next
    'GRILLA (GrdModulos) PARA LA BUSQUEDA
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|Cod_Estado|" _
                              & "PORCENTAJE BONIFICA|PORCENTAJE IVA|" _
                              & "OBSERVACIONES|COD TIPO COMPROBANTE NOTA DEBITO|" _
                              & "BONIFICA EN PESOS|COD CLIENTE|REPRESENTADA|VENDEDOR"
    GrdModulos.ColWidth(0) = 900 'TIPO NOTA DEBITO
    GrdModulos.ColWidth(1) = 1300 'NUMERO
    GrdModulos.ColWidth(2) = 1200 'FECHA
    GrdModulos.ColWidth(3) = 5500 'CLIENTE
    GrdModulos.ColWidth(4) = 0    'COD_ESTADO
    GrdModulos.ColWidth(5) = 0    'PORCENTAJE BONIFICA
    GrdModulos.ColWidth(6) = 0    'PORCENTAJE IVA
    GrdModulos.ColWidth(7) = 0    'OBSERVACIONES
    GrdModulos.ColWidth(8) = 0    'COD TIPO COMPROBANTE NOTA DEBITO
    GrdModulos.ColWidth(9) = 0    'BONIFICA EN PESOS
    GrdModulos.ColWidth(10) = 0   'COD CLIENTE
    GrdModulos.ColWidth(11) = 0   'REPRESENTADA
    GrdModulos.ColWidth(12) = 0   'VENDEDOR
    GrdModulos.Cols = 13
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
    lblEstado.Caption = ""
    'CARGO COMBO CON LOS TIPOS DE NOTA DE DEBITO
    LlenarComboNotaDebito
    
    'CARGO COMBO VENDEDOR
    Call CargoComboBox(CboVend, "VENDEDOR", "VEN_CODIGO", "VEN_NOMBRE")
    CboVend.ListIndex = 0
    
    'CRAGO COMBO REPRESENTADA
    Call CargoComboBox(cboRep, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboRep.ListIndex = 0
    
    Call CargoComboBox(cboBuscaRep, "REPRESENTADA", "REP_CODIGO", "REP_RAZSOC")
    cboBuscaRep.AddItem "(Todas)"
    cboBuscaRep.ListIndex = cboBuscaRep.ListCount - 1
    
    'CARGO ESTADO
    Call BuscoEstado(1, lblEstadoNotaDebito) 'ESTADO PENDIENTE
    VEstadoNotaDebito = 1
    FechaNotaDebito.Text = Date
    tabDatos.Tab = 0
    'BUSCO IVA
    sql = "SELECT IVA FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        VIva = 0 'IIf(IsNull(rec!iva), "", Format(rec!iva, "0.00"))
        VIvaCalculo = 1 '(CDbl(VIva) / 100) + 1
    End If
    rec.Close
    If cboNotaDebito.ItemData(cboNotaDebito.ListIndex) <> 8 Then
        txtPorcentajeIva.Text = VIva
    Else
        txtPorcentajeIva.Text = ""
    End If
    'significa que no estoy buacando
    VBanderaBuscar = False
End Sub

Private Sub LlenarComboNotaDebito()
    sql = "SELECT * FROM TIPO_COMPROBANTE"
    sql = sql & " WHERE TCO_DESCRI LIKE 'NOTA DE DEB%'"
    sql = sql & " ORDER BY TCO_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboNotaDebito1.AddItem "(Todas)"
        Do While rec.EOF = False
            cboNotaDebito.AddItem rec!TCO_DESCRI
            cboNotaDebito.ItemData(cboNotaDebito.NewIndex) = rec!TCO_CODIGO
            cboNotaDebito1.AddItem rec!TCO_DESCRI
            cboNotaDebito1.ItemData(cboNotaDebito1.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboNotaDebito.ListIndex = 0
        cboNotaDebito1.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Select Case grdGrilla.Col
        Case 0, 1, 2, 3, 4
            LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
            grdGrilla.TextMatrix(grdGrilla.RowSel, 11) = grdGrilla.RowSel
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 0
        Case 7
            VBonificacion = 0
            grdGrilla.Text = ""
            grdGrilla.Col = 8
            grdGrilla.Text = ""
            VBonificacion = CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 6))
            grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
            txtSubtotal.Text = Valido_Importe(SumaBonificacion)
            txtTotal.Text = txtSubtotal.Text
            grdGrilla.Col = 0
        End Select
    End If
    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
        Case 4
            If grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = "" _
               And grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = "" And grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = "" _
               And grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = "" Then
                chkBonificaEnPorsentaje.SetFocus
            End If
        End Select
    End If
End Sub

Private Sub grdGrilla_KeyPress(KeyAscii As Integer)
    If (grdGrilla.Col = 0) Or (grdGrilla.Col = 1) Or _
       (grdGrilla.Col = 2) Or (grdGrilla.Col = 3) Or _
       (grdGrilla.Col = 4) Or (grdGrilla.Col = 6) Or (grdGrilla.Col = 7) Then
        If KeyAscii = vbKeyReturn Then
            If grdGrilla.Col = 7 Then
                If grdGrilla.row < grdGrilla.Rows - 1 Then
                    grdGrilla.row = grdGrilla.row + 1
                    grdGrilla.Col = 0
                Else
                    MySendKeys Chr(9)
                End If
            ElseIf grdGrilla.Col = 4 Then
                grdGrilla.Col = 6
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
        Set Rec1 = New ADODB.Recordset
        lblEstado.Caption = "Buscando..."
        Screen.MousePointer = vbHourglass
        'CLIENTE
        txtCodCliente.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 10)
        txtCodCliente_LostFocus
        
        'CABEZA NOTA DEBITO
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 8)), cboNotaDebito)
        'BUSCO LA REPRESENTADA
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 11)), cboRep)
        'BUSCO VENDEDOR
        Call BuscaCodigoProxItemData(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 12)), CboVend)
        
        txtNroNotaDebito.Text = Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8)
        txtNroSucursal.Text = Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4)
        FechaNotaDebito.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 2)
        Call BuscoEstado(CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4)), lblEstadoNotaDebito)
        VEstadoNotaDebito = CInt(GrdModulos.TextMatrix(GrdModulos.RowSel, 4))
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 7) <> "" Then
            txtObservaciones.Text = Trim(GrdModulos.TextMatrix(GrdModulos.RowSel, 7))
        End If
        
        '----BUSCO DETALLE DE LA NOTA DE DEBITO------------------
        sql = "SELECT NDC.*, B.BAN_BANCO, B.BAN_LOCALIDAD, B.BAN_SUCURSAL, B.BAN_CODIGO, B.BAN_DESCRI"
        sql = sql & " FROM DETALLE_NOTA_DEBITO_CLIENTE NDC, BANCO B"
        sql = sql & " WHERE NDC.NDC_NUMERO=" & XN(Right(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 8))
        sql = sql & " AND NDC.NDC_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(GrdModulos.RowSel, 1), 4))
        sql = sql & " AND NDC.REP_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 11))
        sql = sql & " AND NDC.NDC_FECHA=" & XDQ(GrdModulos.TextMatrix(GrdModulos.RowSel, 2))
        sql = sql & " AND NDC.TCO_CODIGO=" & XN(GrdModulos.TextMatrix(GrdModulos.RowSel, 8))
        sql = sql & " AND NDC.BAN_CODINT=B.BAN_CODINT"
        sql = sql & " ORDER BY NDC.DND_NROITEM"
        Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec1.EOF = False Then
            i = 1
            Do While Rec1.EOF = False
                grdGrilla.TextMatrix(i, 0) = Rec1!CHE_NUMERO
                grdGrilla.TextMatrix(i, 1) = Rec1!BAN_BANCO
                grdGrilla.TextMatrix(i, 2) = Rec1!BAN_LOCALIDAD
                grdGrilla.TextMatrix(i, 3) = Rec1!BAN_SUCURSAL
                grdGrilla.TextMatrix(i, 4) = Rec1!BAN_CODIGO
                grdGrilla.TextMatrix(i, 5) = Rec1!BAN_DESCRI
                grdGrilla.TextMatrix(i, 6) = Valido_Importe(Rec1!DND_PRECIO)  'IMPORTE CHEQUE
                If IsNull(Rec1!DND_BONIFICA) Then
                    grdGrilla.TextMatrix(i, 7) = ""
                Else
                    grdGrilla.TextMatrix(i, 7) = Valido_Importe(Rec1!DND_BONIFICA)
                End If
                VBonificacion = 0
                If Not IsNull(Rec1!DND_BONIFICA) Then
                    VBonificacion = ((CDbl(Rec1!DND_PRECIO) * CDbl(Rec1!DND_BONIFICA)) / 100)
                    VBonificacion = (CDbl(Rec1!DND_PRECIO) - VBonificacion)
                    grdGrilla.TextMatrix(i, 8) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(i, 9) = Valido_Importe(CStr(VBonificacion))
                Else
                    VBonificacion = (CDbl(Rec1!DND_PRECIO))
                    grdGrilla.TextMatrix(i, 8) = ""
                    grdGrilla.TextMatrix(i, 9) = Valido_Importe(CStr(VBonificacion))
                End If
                grdGrilla.TextMatrix(i, 10) = Rec1!BAN_CODINT
                grdGrilla.TextMatrix(i, 11) = Rec1!DND_NROITEM
                i = i + 1
                Rec1.MoveNext
            Loop
            VBonificacion = 0
        End If
        Rec1.Close
        '--CARGO LOS TOTALES----
        txtSubtotal.Text = Valido_Importe(SumaBonificacion)
        txtTotal.Text = txtSubtotal.Text
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 9) = "S" Then 'SI BONOFICA EN PESOS
            chkBonificaEnPesos.Value = Checked
        ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 9) = "N" Then 'SI BONIFICA EN PORCENTAJE
            chkBonificaEnPorsentaje.Value = Checked
        Else
            chkBonificaEnPesos.Value = Unchecked
            chkBonificaEnPorsentaje.Value = Unchecked
        End If
        If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) <> "" Then 'PORCENTAJE DE BONIFICACION
            txtPorcentajeBoni.Text = GrdModulos.TextMatrix(GrdModulos.RowSel, 5)
            txtPorcentajeBoni_LostFocus
        End If
        If cboNotaDebito.ItemData(cboNotaDebito.ListIndex) <> 8 And GrdModulos.TextMatrix(GrdModulos.RowSel, 6) <> "" Then 'PORCENTAJE IVA
            txtPorcentajeIva = GrdModulos.TextMatrix(GrdModulos.RowSel, 6)
            txtPorcentajeIva_LostFocus
        Else
            txtPorcentajeIva.Text = ""
        End If
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        '--------------
        FrameNotaDebito.Enabled = False
        FramePara.Enabled = False
        '--------------
        'significa que estoy buacando
        VBanderaBuscar = True
        tabDatos.Tab = 0
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdModulos_dblClick
End Sub

Private Sub tabDatos_Click(PreviousTab As Integer)
  If tabDatos.Tab = 1 Then
    GrdModulos.Rows = 1
    cmdGrabar.Enabled = False
    LimpiarBusqueda
    If Me.Visible = True Then txtCliente.SetFocus
  Else
    If VEstadoNotaDebito = 1 Then
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
    GrdModulos.HighLight = flexHighlightNever
    cboNotaDebito1.ListIndex = 0
    cboBuscaRep.ListIndex = cboBuscaRep.ListCount - 1
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
                frmBuscar.TxtDescriB.Text = txtCliRazSoc.Text
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
        Set Rec1 = New ADODB.Recordset
        Rec1.Open BuscoCliente(txtCodCliente), DBConn, adOpenStatic, adLockOptimistic
        
        If Rec1.EOF = False Then
            txtCliRazSoc.Text = Rec1!CLI_RAZSOC
            txtProvincia.Text = Rec1!PRO_DESCRI
            txtCliLocalidad.Text = Rec1!LOC_DESCRI
            txtDomici.Text = Rec1!CLI_DOMICI
            
            If Rec1!IVA_CODIGO = 1 Then 'RESPONSABLE INSCRIPTO
                Call BuscaProx("NOTA DE DEBITO A", cboNotaDebito)
                If CInt(VRepresentada) = cboRep.ItemData(cboRep.ListIndex) Then
                   'or CInt(VRepresentada2) = cboRep.ItemData(cboRep.ListIndex) Then
                    txtNroSucursal.Text = ""
                    txtNroSucursal_LostFocus
                    txtNroNotaDebito.Text = ""
                    txtNroNotaDebito_LostFocus
                Else
                    txtNroSucursal.Text = VSucursal
                    txtNroNotaDebito.Text = VNroNd
                    txtPorcentajeIva.Text = VIva
                End If
            ElseIf Rec1!IVA_CODIGO = 2 Or Rec1!IVA_CODIGO = 3 Then
                Call BuscaProx("NOTA DE DEBITO B", cboNotaDebito)
                If CInt(VRepresentada) = cboRep.ItemData(cboRep.ListIndex) Then
                   'or CInt(VRepresentada2) = cboRep.ItemData(cboRep.ListIndex) Then
                    txtNroSucursal.Text = ""
                    txtNroSucursal_LostFocus
                    txtNroNotaDebito.Text = ""
                    txtNroNotaDebito_LostFocus
                Else
                    txtNroSucursal.Text = VSucursal
                    txtNroNotaDebito.Text = VNroNd
                    txtPorcentajeIva.Text = ""
                End If
            End If
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCodCliente.SetFocus
        End If
        Rec1.Close
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
    If grdGrilla.Col = 0 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 10
    End If
    If grdGrilla.Col = 1 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 3
    End If
    If grdGrilla.Col = 2 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 3
    End If
    If grdGrilla.Col = 3 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 3
    End If
    If grdGrilla.Col = 4 Then
        KeyAscii = CarNumeroEntero(KeyAscii)
        txtEdit.MaxLength = 6
    End If
    If grdGrilla.Col = 6 Or grdGrilla.Col = 7 Then
        KeyAscii = CarNumeroDecimal(txtEdit, KeyAscii)
        txtEdit.MaxLength = 15
    End If
    CarTexto KeyAscii
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    VBonificacion = 0
    If KeyCode = vbKeyF1 Then
        frmBuscar.TipoBusqueda = 2
        frmBuscar.CodListaPrecio = 0
        grdGrilla.Col = 0
        EDITAR grdGrilla, txtEdit, 13
        frmBuscar.Show vbModal
    End If

    If KeyCode = vbKeyReturn Then
        Select Case grdGrilla.Col
            
            Case 0 'NUMERO CHEQUE
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "0000000000")
                    grdGrilla.Col = 1
                End If
                
            Case 1 'BCO
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000")
                    grdGrilla.Col = 2
                End If
                
            Case 2 'LOC
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000")
                    grdGrilla.Col = 3
                End If
            
            Case 3 'SUC
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000")
                    grdGrilla.Col = 4
                End If
            
            Case 4 'CODIGO
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Format(txtEdit.Text, "000000")
                    grdGrilla.Col = 6
                End If
                'BUSCO EL BANCO-------------------------------------
                If grdGrilla.TextMatrix(grdGrilla.row, 0) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 1) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 2) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 3) <> "" And _
                    txtEdit.Text <> "" Then
                    
                    'BUSCO EL CODIGO INTERNO
                    sql = "SELECT BAN_CODINT, BAN_DESCRI"
                    sql = sql & " FROM BANCO"
                    sql = sql & " WHERE BAN_BANCO = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 1))
                    sql = sql & " AND BAN_LOCALIDAD = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 2))
                    sql = sql & " AND BAN_SUCURSAL = " & XS(grdGrilla.TextMatrix(grdGrilla.row, 3))
                    sql = sql & " AND BAN_CODIGO = " & XS(txtEdit.Text)
                    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
                    If rec.EOF = False Then 'EXITE
                       grdGrilla.TextMatrix(grdGrilla.row, 10) = rec!BAN_CODINT
                       grdGrilla.TextMatrix(grdGrilla.row, 5) = rec!BAN_DESCRI
                       rec.Close
                    Else
                       If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
                         MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
                         grdGrilla.Col = 1
                         grdGrilla.SetFocus
                       End If
                       rec.Close
                       Exit Sub
                    End If
                Else
                    MsgBox "Faltan Datos", vbCritical, TIT_MSGBOX
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                End If
                
            Case 6 'IMPORTE CHEQUE
                If Trim(txtEdit) = "" Then
                    txtEdit = ""
                    LimpiarFilasDeGrilla grdGrilla, grdGrilla.row
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                Else
                    txtEdit.Text = Valido_Importe(CStr(CDbl(txtEdit.Text) * VIvaCalculo))
                    grdGrilla.Col = 7
                End If
                
                If grdGrilla.TextMatrix(grdGrilla.row, 0) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 1) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 2) <> "" And _
                    grdGrilla.TextMatrix(grdGrilla.row, 3) <> "" And _
                    txtEdit.Text <> "" Then
                
                    VBonificacion = CDbl(txtEdit.Text)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                    If grdGrilla.TextMatrix(grdGrilla.RowSel, 7) <> "" Then
                        VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) * CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 7))) / 100)
                        VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) - VBonificacion)
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(CStr(VBonificacion))
                        grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                    End If
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                    txtPorcentajeIva_LostFocus
                Else
                    MsgBox "No puede ingresar el importe del cheque, Faltan datos!!", vbExclamation, TIT_MSGBOX
                    grdGrilla.TextMatrix(grdGrilla.row, 6) = ""
                    grdGrilla.Col = 0
                    grdGrilla.SetFocus
                    Exit Sub
                End If
                
            Case 7 'BONIFICACION
                If Trim(txtEdit) <> "" Then
                    If txtEdit.Text = ValidarPorcentaje(txtEdit) = False Then
                        Exit Sub
                    End If
                    VBonificacion = ((CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) * CDbl(txtEdit.Text)) / 100)
                    VBonificacion = (CDbl(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) - VBonificacion)
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = Valido_Importe(CStr(VBonificacion))
                    grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = Valido_Importe(CStr(VBonificacion))
                    txtSubtotal.Text = Valido_Importe(SumaBonificacion)
                    txtTotal.Text = txtSubtotal.Text
                    txtPorcentajeIva_LostFocus
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
                MsgBox "El Servicio ya fue elegido anteriormente", vbExclamation, TIT_MSGBOX
                BuscoRepetetidos = False
                Exit Function
            End If
        End If
    Next
    BuscoRepetetidos = True
End Function

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

Private Function SumaTotal() As Double
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 9) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 9))
        End If
    Next
    SumaTotal = Valido_Importe(CStr(VTotal))
End Function

Private Function SumaBonificacion() As Double
    VTotal = 0
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 9) <> "" Then
            VTotal = VTotal + CDbl(grdGrilla.TextMatrix(i, 9))
        End If
    Next
    SumaBonificacion = Valido_Importe(CStr(VTotal))
End Function

Private Sub txtNroNotaDebito_GotFocus()
    SelecTexto txtNroNotaDebito
End Sub

Private Sub txtNroNotaDebito_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtNroNotaDebito_LostFocus()
    If VBanderaBuscar = False Then
        If txtNroNotaDebito.Text = "" Then
            'BUSCO EL NUMERO DE NOTA DE DEBITO QUE CORRESPONDE
            txtNroNotaDebito.Text = BuscoUltimoNumeroComprobante(cboRep.ItemData(cboRep.ListIndex), cboNotaDebito.ItemData(cboNotaDebito.ListIndex))
        Else
            txtNroNotaDebito.Text = Format(txtNroNotaDebito.Text, "00000000")
        End If
        If cboNotaDebito.ItemData(cboNotaDebito.ListIndex) <> 8 Then
            txtPorcentajeIva.Text = VIva
        Else
            txtPorcentajeIva.Text = ""
        End If
        VSucursal = txtNroSucursal.Text
        VNroNd = txtNroNotaDebito.Text
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
        ElseIf chkBonificaEnPesos.Value = Checked Then
            txtPorcentajeBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtImporteBoni.Text = Valido_Importe(txtPorcentajeBoni.Text)
            txtTotal.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtTotal.Text = Valido_Importe(txtTotal.Text)
            txtSubTotalBoni.Text = CDbl(txtSubtotal.Text) - CDbl(txtImporteBoni.Text)
            txtSubTotalBoni.Text = Valido_Importe(txtSubTotalBoni.Text)
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
        Else
            txtPorcentajeIva.Text = Format(VIva, "0.00")
            txtPorcentajeIva_LostFocus
        End If
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
