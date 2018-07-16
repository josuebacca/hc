VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmhistoriaclinica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia Clinica"
   ClientHeight    =   10935
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   16725
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Rep 
      Left            =   600
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Paciente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16455
      Begin VB.TextBox txtEdad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "Descripción"
         Top             =   360
         Width           =   555
      End
      Begin VB.TextBox txtNAfil 
         Height          =   285
         Left            =   5160
         TabIndex        =   90
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txthorad 
         Height          =   285
         Left            =   3960
         TabIndex        =   88
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Top             =   -120
         Visible         =   0   'False
         Width           =   495
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
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   17
         Top             =   360
         Width           =   1395
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
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   15
         Tag             =   "Descripción"
         Top             =   360
         Width           =   3435
      End
      Begin VB.TextBox txtTelefono 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   15000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "Descripción"
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtOSocial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   9780
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "Descripción"
         Top             =   360
         Width           =   3795
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Edad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7230
         TabIndex        =   91
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2940
         TabIndex        =   21
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero/DNI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13680
         TabIndex        =   19
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Obra Social:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8460
         TabIndex        =   18
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdAgregarPedido 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   13080
      TabIndex        =   25
      Top             =   10080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   15240
      TabIndex        =   24
      Top             =   10080
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   14160
      TabIndex        =   23
      Top             =   10080
      Width           =   1095
   End
   Begin TabDlg.SSTab tabhc 
      Height          =   8535
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Curso Clinico"
      TabPicture(0)   =   "frmhistoriaclinica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ecografias"
      TabPicture(1)   =   "frmhistoriaclinica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdabrirdoc"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "cmdVerEstudio"
      Tab(1).Control(3)=   "cmdVer"
      Tab(1).Control(4)=   "cmdAgregarEco"
      Tab(1).Control(5)=   "cmdEliminarEco"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Laboratorio"
      TabPicture(2)   =   "frmhistoriaclinica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Ginecologia"
      TabPicture(3)   =   "frmhistoriaclinica.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "cmdImprimirEstGine"
      Tab(3).Control(2)=   "cmdEliminarEstGine"
      Tab(3).Control(3)=   "cmdAgregarEstGine"
      Tab(3).Control(4)=   "cboTipoEstGine"
      Tab(3).Control(5)=   "optFechaGine"
      Tab(3).Control(6)=   "optTipoEst"
      Tab(3).Control(7)=   "Label8"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Pedidos"
      TabPicture(4)   =   "frmhistoriaclinica.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(1)=   "Frame8"
      Tab(4).ControlCount=   2
      Begin VB.CommandButton cmdabrirdoc 
         Height          =   495
         Left            =   -67200
         Picture         =   "frmhistoriaclinica.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Abrir desde un archivo existente"
         Top             =   2520
         Width           =   495
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   120
         TabIndex        =   45
         Top             =   7260
         Width           =   8055
         Begin VB.CommandButton cmdGineco 
            Caption         =   "&Ginecologia"
            Height          =   855
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdSiguiente 
            Caption         =   "&Siguiente Paciente"
            Height          =   855
            Left            =   480
            TabIndex        =   59
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdLabora 
            Caption         =   "&Laboratorio"
            Height          =   855
            Left            =   4080
            TabIndex        =   52
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdEcogra 
            Caption         =   "&Ecografias"
            Height          =   855
            Left            =   2880
            TabIndex        =   51
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAnterior 
            Caption         =   "&Anterior Paciente"
            Height          =   855
            Left            =   1680
            TabIndex        =   49
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdPedidos 
            Caption         =   "&Pedidos"
            Height          =   855
            Left            =   6480
            TabIndex        =   48
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Imágenes anteriores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   -66480
         TabIndex        =   116
         Top             =   480
         Width           =   8055
         Begin VB.ComboBox cboImgAnt 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   360
            Width           =   3735
         End
         Begin VB.ComboBox cboDocImgAnt 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   375
            Width           =   2295
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Filtro"
            Height          =   375
            Left            =   6960
            TabIndex        =   117
            Top             =   720
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesdeImg 
            Height          =   315
            Left            =   1335
            TabIndex        =   119
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHastaImg 
            Height          =   315
            Left            =   4575
            TabIndex        =   120
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdImagenes 
            Height          =   6630
            Left            =   120
            TabIndex        =   121
            Top             =   1200
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   11695
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label39 
            Caption         =   "Imágen:"
            Height          =   255
            Left            =   3480
            TabIndex        =   128
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   240
            TabIndex        =   124
            Top             =   435
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   123
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   240
            TabIndex        =   122
            Top             =   795
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ecografia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   -74040
         TabIndex        =   92
         Top             =   720
         Width           =   8295
         Begin VB.TextBox Text3 
            Height          =   5115
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   98
            Top             =   960
            Width           =   6375
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   96
            Top             =   6360
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   95
            Top             =   6360
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4515
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   600
            Width           =   3180
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   2880
            TabIndex        =   93
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   7680
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1305
            TabIndex        =   97
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   690
            TabIndex        =   101
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Indicaciones:"
            Height          =   675
            Left            =   240
            TabIndex        =   100
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   3840
            TabIndex        =   99
            Top             =   660
            Width           =   540
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Pedidos anteriores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8175
         Left            =   -66720
         TabIndex        =   69
         Top             =   480
         Width           =   8055
         Begin VB.ComboBox cboDocPedidos 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2055
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   375
            Width           =   3975
         End
         Begin VB.CommandButton cmdFiltroPedidos 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   87
            Top             =   360
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesdePedido 
            Height          =   315
            Left            =   2025
            TabIndex        =   85
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHastaPedido 
            Height          =   315
            Left            =   4575
            TabIndex        =   86
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdPedidos 
            Height          =   6375
            Left            =   120
            TabIndex        =   70
            Top             =   1200
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   11245
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   1290
            TabIndex        =   73
            Top             =   435
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   2
            Left            =   3600
            TabIndex        =   72
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   840
            TabIndex        =   71
            Top             =   795
            Width           =   990
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Pedido Medico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8175
         Left            =   -74880
         TabIndex        =   60
         Top             =   480
         Width           =   8055
         Begin VB.TextBox txtnroPedido 
            Height          =   315
            Left            =   2880
            TabIndex        =   75
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cboEspecPedido 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmhistoriaclinica.frx":6322
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":6324
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   1920
            Width           =   2220
         End
         Begin VB.CommandButton cmdImprimirPedido 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   4560
            TabIndex        =   81
            Top             =   7680
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarPedido 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   82
            Top             =   7680
            Width           =   1095
         End
         Begin VB.TextBox txtMotivoPedido 
            Height          =   315
            Left            =   1305
            TabIndex        =   78
            Top             =   1440
            Width           =   6375
         End
         Begin VB.TextBox txtDescPedido 
            Height          =   5115
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   80
            Top             =   2400
            Width           =   6375
         End
         Begin VB.CommandButton cmdCancelarPedido 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   83
            Top             =   7680
            Width           =   1095
         End
         Begin VB.ComboBox cboDocPedido 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmhistoriaclinica.frx":6326
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":6328
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   1080
            Width           =   2220
         End
         Begin VB.TextBox txtConsultorioPedido 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   62
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtProfesionPedido 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   61
            Top             =   1080
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker FechaPed 
            Height          =   315
            Left            =   1305
            TabIndex        =   76
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   77
            Top             =   660
            Width           =   495
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Especialidad:"
            Height          =   195
            Left            =   240
            TabIndex        =   74
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   240
            TabIndex        =   68
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   67
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   1140
            Width           =   540
         End
         Begin VB.Label Label20 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3720
            TabIndex        =   65
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6240
            TabIndex        =   64
            Top             =   1110
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Consulta Medica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   120
         TabIndex        =   33
         Top             =   420
         Width           =   8055
         Begin VB.TextBox txtProfesion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4680
            TabIndex        =   41
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtConsultorio 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   42
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtnrocon 
            Height          =   315
            Left            =   2880
            TabIndex        =   56
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComCtl2.DTPicker FechaProx 
            Height          =   315
            Left            =   2040
            TabIndex        =   47
            Top             =   6360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   43205
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   38
            Top             =   6360
            Width           =   1095
         End
         Begin VB.TextBox txtIndicaciones 
            Height          =   4035
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   2040
            Width           =   6375
         End
         Begin VB.TextBox txtMotivo 
            Height          =   315
            Left            =   1305
            TabIndex        =   44
            Top             =   1440
            Width           =   6375
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   1305
            TabIndex        =   35
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   39
            Top             =   6360
            Width           =   1095
         End
         Begin VB.ComboBox cboDocCon 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmhistoriaclinica.frx":632A
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":632C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1080
            Width           =   2580
         End
         Begin VB.Label Label18 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6360
            TabIndex        =   58
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3960
            TabIndex        =   57
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Próxima Consulta:"
            Height          =   375
            Left            =   360
            TabIndex        =   55
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   600
            TabIndex        =   43
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Indicaciones:"
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   600
            TabIndex        =   36
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   1560
            Width           =   525
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Consultas anteriores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   8280
         TabIndex        =   26
         Top             =   420
         Width           =   8055
         Begin VB.CommandButton cmdFiltro 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   54
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboDocAnt 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2055
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   375
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2025
            TabIndex        =   29
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   4575
            TabIndex        =   30
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdConsultas 
            Height          =   6630
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   11695
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   840
            TabIndex        =   32
            Top             =   795
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   0
            Left            =   3600
            TabIndex        =   31
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   1290
            TabIndex        =   28
            Top             =   435
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdImprimirEstGine 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -57480
         TabIndex        =   12
         Top             =   9060
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEstGine 
         Caption         =   "Eliminar Estudio"
         Height          =   375
         Left            =   -59400
         TabIndex        =   11
         Top             =   9060
         Width           =   1575
      End
      Begin VB.CommandButton cmdAgregarEstGine 
         Caption         =   "Agregar Estudio"
         Height          =   375
         Left            =   -61080
         TabIndex        =   10
         Top             =   9060
         Width           =   1335
      End
      Begin VB.ComboBox cboTipoEstGine 
         Height          =   315
         Left            =   -73320
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   1380
         Width           =   2775
      End
      Begin VB.OptionButton optFechaGine 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   -68400
         TabIndex        =   8
         Top             =   900
         Width           =   2055
      End
      Begin VB.OptionButton optTipoEst 
         Caption         =   "Tipo Estudio"
         Height          =   375
         Left            =   -73320
         TabIndex        =   7
         Top             =   900
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdVerEstudio 
         Caption         =   "Ver Estudio(VA ACA?)"
         Height          =   375
         Left            =   -60360
         TabIndex        =   5
         Top             =   8580
         Width           =   1335
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -61800
         TabIndex        =   4
         Top             =   8580
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregarEco 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -58800
         TabIndex        =   3
         Top             =   9060
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEco 
         Caption         =   "Elinimar"
         Height          =   375
         Left            =   -57000
         TabIndex        =   2
         Top             =   9060
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Imágen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   -74880
         TabIndex        =   102
         Top             =   480
         Width           =   8295
         Begin VB.CommandButton cmdImprimirEco 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   4560
            TabIndex        =   131
            Top             =   7440
            Width           =   1095
         End
         Begin VB.ComboBox cboImg 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmhistoriaclinica.frx":632E
            Left            =   1320
            List            =   "frmhistoriaclinica.frx":6330
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   1080
            Width           =   4980
         End
         Begin VB.ComboBox cboDocImg 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmhistoriaclinica.frx":6332
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":6334
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   1560
            Width           =   2580
         End
         Begin VB.CommandButton cmdAceptarImg 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   109
            Top             =   7440
            Width           =   1095
         End
         Begin VB.TextBox txtImgDescri 
            Height          =   5355
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   107
            Top             =   2040
            Width           =   6375
         End
         Begin VB.CommandButton cmdCancelarImg 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   106
            Top             =   7440
            Width           =   1095
         End
         Begin VB.TextBox txtNroImg 
            Height          =   315
            Left            =   2880
            TabIndex        =   105
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtConsulImg 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   104
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtProfImg 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4680
            TabIndex        =   103
            Top             =   1560
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker FechaImg 
            Height          =   315
            Left            =   1305
            TabIndex        =   108
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   110821377
            CurrentDate     =   41098
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Imágen:"
            Height          =   195
            Left            =   600
            TabIndex        =   126
            Top             =   1080
            Width           =   570
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   600
            TabIndex        =   115
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   114
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   600
            TabIndex        =   113
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3960
            TabIndex        =   112
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6360
            TabIndex        =   111
            Top             =   1560
            Width           =   855
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Buscar por:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   900
         Width           =   1095
      End
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Doctor:"
      Height          =   195
      Left            =   720
      TabIndex        =   125
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Nro Carnet:"
      Height          =   195
      Left            =   5640
      TabIndex        =   89
      Top             =   2160
      Width           =   810
   End
End
Attribute VB_Name = "frmhistoriaclinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rec2 As New ADODB.Recordset
Dim edad As Integer
Dim años As Integer
Public NroAfil As String
Public TurOSocial As String

Private Function BuscarProxPaciente(codven, DIA) As Integer
    Dim CodPac As Integer
    CodPac = 0
    sql = " SELECT TOP 1 C.CLI_CODIGO,C.CLI_NROAFIL,CLI_CUMPLE, T.TUR_HORAD, T.TUR_OSOCIAL "
    sql = sql & " FROM TURNOS T, CLIENTE C "
    sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND T.VEN_CODIGO = " & codven
    sql = sql & " AND T.TUR_FECHA = " & DIA
    sql = sql & " ORDER BY  T.TUR_HORAD ASC"
    Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec2.EOF = False Then
        CodPac = Rec2!CLI_CODIGO
        txthorad = Format(Rec2!TUR_HORAD, "hh:mm")
        NroAfil = ChkNull(Rec2!CLI_NROAFIL)
        TurOSocial = ChkNull(Rec2!TUR_OSOCIAL)
        Calculo_Edad IIf(IsNull(Rec2!CLI_CUMPLE), Date, Rec2!CLI_CUMPLE)
    End If
    Rec2.Close
    BuscarProxPaciente = CodPac
End Function
Private Function Calculo_Edad(cumple As Date)
    'calculo de edad
    If Not (IsNull(cumple)) Then
        años = Year(Date) - Year(cumple)
        If Month(Fecha) < Month(cumple) Then años = años - 1 'todavía no ha llegado el mes de su cumple
         If Month(Now) = Month(cumple) And Day(Fecha) < Day(cumple) Then años = años - 1 'es el mes pero no ha llegado el día de su cumple
        edad = años
    Else
        edad = 0
    End If
    txtEdad.Text = edad
End Function
Private Function validarcclinico() As Boolean
    If txtBuscaCliente.Text = "" Then
        MsgBox "No ha ingresado el paciente", vbCritical, TIT_MSGBOX
        txtBuscaCliente.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If cboDocCon.ListIndex = -1 Then
        MsgBox "No ha ingresado el doctor", vbCritical, TIT_MSGBOX
        cboDesde.SetFocus
        validarcclinico = False
        Exit Function
    End If
    If IsNull(Fecha.Value) Then
        MsgBox "No ha ingresado la fecha de la consulta", vbCritical, TIT_MSGBOX
        Fecha.SetFocus
        validarcclinico = False
        Exit Function
    End If
    
        If txtMotivo.Text = "" Then
        MsgBox "No ha ingresado el motivo", vbCritical, TIT_MSGBOX
        txtMotivo.SetFocus
        validarcclinico = False
        Exit Function
    End If

    If txtIndicaciones.Text = "" Then
        MsgBox "No ha ingresado la indicación", vbCritical, TIT_MSGBOX
        txtMotivo.SetFocus
        validarcclinico = False
        Exit Function
    End If
    
    validarcclinico = True
End Function
Private Function validarImagen()
    If txtBuscaCliente.Text = "" Then
        MsgBox "No ha ingresado el paciente", vbCritical, TIT_MSGBOX
        txtBuscaCliente.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If txtImgDescri.Text = "" Then
        MsgBox "No ha ingresado la descripción de la Imagen", vbCritical, TIT_MSGBOX
        txtMotivo.SetFocus
        validarImagen = False
        Exit Function
    End If

    If cboDocCon.ListIndex = -1 Then
        MsgBox "No ha ingresado el doctor", vbCritical, TIT_MSGBOX
        cboDesde.SetFocus
        validarImagen = False
        Exit Function
    End If
    If IsNull(FechaImg.Value) Then
        MsgBox "No ha ingresado la fecha de la Imágen", vbCritical, TIT_MSGBOX
        FechaImg.SetFocus
        validarImagen = False
        Exit Function
    End If
 
    validarImagen = True
End Function
Private Function ImprimirTurno()
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
            
    Rep.SelectionFormula = " {TURNOS.TUR_FECHA}= " & XDQ(MViewFecha.Value)
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.VEN_CODIGO}= " & cboDoctor.ItemData(cboDoctor.ListIndex)
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.CLI_CODIGO}= " & XN(txtCodigo.Text)
    'Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.TUR_HORAD}= #" & TRIM(cboDesde.Text) & "#"
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    'Rep.Connect = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & SERVIDOR & ";"
    Rep.WindowTitle = "Impresion del Turno"
    Rep.ReportFileName = DirReport & "rptTurno.rpt"
    Rep.Action = 1
End Function
Private Function LimpiarImagen()
    FechaImg.Value = Date
    txtImgDescri = ""
    cboImg.ListIndex = -1
    txtNroImg = ""
End Function
Private Sub cmdAgregar_Click()
    
    
End Sub



Private Sub b_Click()

End Sub

Private Sub cboDocCon_Change()
sql = "SELECT PR_CODIGO, VEN_CONSULTORIO FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO = "
    sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    txtConsultorio.Text = rec!VEN_CONSULTORIO
    txtProfesion.Text = rec!PR_CODIGO
    rec.Close
  
End Sub

Private Sub cboDocCon_Click()
    Dim pro As String
    LimpiarConsulta
    LimpiarPedido
    
    sql = "SELECT VEN_CODIGO,PR_CODIGO, VEN_CONSULTORIO FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO = "
    sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtConsultorio.Text = Chk0(rec!VEN_CONSULTORIO)
        pro = Chk0(rec!PR_CODIGO)
        'defino profesion y consultorio en pedido
        txtConsultorioPedido.Text = Chk0(rec!VEN_CONSULTORIO)
    End If
    codven = rec!VEN_CODIGO
    rec.Close
    
    sql = "SELECT PR_DESCRI FROM PROFESION"
    sql = sql & " WHERE PR_CODIGO ="
    sql = sql & pro
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtProfesion.Text = rec!PR_DESCRI
        txtProfesionPedido = rec!PR_DESCRI
    End If
    rec.Close
    
    BuscaCodigoProxItemData cboDocCon.ItemData(cboDocCon.ListIndex), cboDocPedido
    If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then 'VEERR
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
    
End Sub

Private Sub cboDocImg_Change()
   sql = "SELECT PR_CODIGO, VEN_CONSULTORIO FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO = "
    sql = sql & cboDocImg.ItemData(cboDocImg.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    txtConsulImg.Text = rec!VEN_CONSULTORIO
    txtProfImg.Text = rec!PR_CODIGO
    rec.Close
End Sub

Private Sub cboDocImg_Click()
     Dim pro As String
    LimpiarImagen
    'LimpiarPedido
    sql = "SELECT VEN_CODIGO,PR_CODIGO, VEN_CONSULTORIO FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO = "
    sql = sql & cboDocImg.ItemData(cboDocImg.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtConsulImg.Text = Chk0(rec!VEN_CONSULTORIO)
        pro = Chk0(rec!PR_CODIGO)
    End If
    codven = rec!VEN_CODIGO
    rec.Close
    
    sql = "SELECT PR_DESCRI FROM PROFESION"
    sql = sql & " WHERE PR_CODIGO ="
    sql = sql & pro
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtProfImg.Text = rec!PR_DESCRI
    End If
    rec.Close
    If cboDocImg.ItemData(cboDocImg.ListIndex) = Int(Doc) Then 'VEERR
        cmdAceptarImg.Enabled = True
    Else
        cmdAceptarImg.Enabled = False
    End If
    
    'BuscaCodigoProxItemData cboDocImg.ItemData(cboDocImg.ListIndex), cboDocPedido
End Sub

Private Sub cboDocPedido_Click()
sql = "SELECT PR_CODIGO, VEN_CONSULTORIO FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO = "
    sql = sql & cboDocPedido.ItemData(cboDocPedido.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtConsultorioPedido.Text = Chk0(rec!VEN_CONSULTORIO)
        pro = rec!PR_CODIGO

    End If
    rec.Close
    
    sql = "SELECT PR_DESCRI FROM PROFESION"
    sql = sql & " WHERE PR_CODIGO ="
    sql = sql & pro
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    txtProfesionPedido = rec!PR_DESCRI
    rec.Close
End Sub

Private Sub cmdAbrir_Click()
    Set word = CreateObject("word.Basic")
    cmdAceptar.Enabled = True
    On Error Resume Next
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Seleccione un nombre de archivo"
    CommonDialog1.Filter = "Documents(*.doc;*.docx)"
    
    CommonDialog1.ShowOpen
    If Err.Number = 0 Then
        'If CommonDialog1.FileName Like "*.bmp" _
        'Or CommonDialog1.FileName Like "*.gif" _
        'Or CommonDialog1.FileName Like "*.jpg" Then
            
            'Image1.Picture = LoadPicture(CommonDialog1.FileName)
            'txtimagen.Text = CommonDialog1.FileName
            word.FileOpen (CommonDialog1.FileName)
            word.AppShow
            'word.filePrintDefault
            On Error GoTo 0
        'Else
        '    MsgBox "El Archivo seleccionado no es válido", vbExclamation, Me.Caption
        'End If
        
    End If
    'word.AppClose
End Sub

Private Sub cmdabrirdoc_Click()
    Set word = CreateObject("word.Basic")
    cmdAceptar.Enabled = True
    On Error Resume Next
    CommonDialog2.CancelError = True
    CommonDialog2.DialogTitle = "Seleccione un nombre de archivo"
    CommonDialog2.Filter = "Documents(*.doc;*.docx)"
    
    CommonDialog2.ShowOpen
    If Err.Number = 0 Then
        'If CommonDialog1.FileName Like "*.bmp" _
        'Or CommonDialog1.FileName Like "*.gif" _
        'Or CommonDialog1.FileName Like "*.jpg" Then
            
            'Image1.Picture = LoadPicture(CommonDialog1.FileName)
            'txtimagen.Text = CommonDialog1.FileName
            word.FileOpen (CommonDialog2.FileName)
            word.AppShow
            'word.filePrintDefault
            On Error GoTo 0
        'Else
        '    MsgBox "El Archivo seleccionado no es válido", vbExclamation, Me.Caption
        'End If
        
    End If
    'word.AppClose
End Sub

Private Sub cmdAceptar_Click()
    Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    Dim i As Integer
    Dim Num As Integer
    
    'Validar los campos requeridos
    If validarcclinico = False Then Exit Sub
    If txtnrocon.Text = "" Then
        If MsgBox("¿Desea cargar la Consulta Medica?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        'agregar teniendo en cuentas loc combos de horas
        'On Error GoTo HayErrorTurno

        i = 0
        sql = "SELECT MAX(CCL_NUMERO) as ultimo FROM CCLINICO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Num = Chk0(rec!Ultimo) + 1 ' guardo el proxuimo turno
        Else
            Num = 1
        End If
        rec.Close
        
        'NUEVA CONSULTA
        sql = "INSERT INTO CCLINICO"
        sql = sql & " (CCL_NUMERO,CCL_FECHA,"
        sql = sql & " CLI_CODIGO,VEN_CODIGO,CCL_MOTIVO,CCL_INDICA,CCL_FECPC)"
        sql = sql & " VALUES ("
        sql = sql & Num & ","
        sql = sql & XDQ(Fecha.Value) & ","
        sql = sql & XN(txtCodigo.Text) & ","
        sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex) & ","
        sql = sql & XS(txtMotivo.Text) & ","
        sql = sql & XS(txtIndicaciones.Text) & ","
        sql = sql & XDQ(ChkNull(FechaProx.Value)) & ")"
'        If optSI2 = True Then
'            sql = sql & XS("SI") & ")"
'        Else
'            sql = sql & XS("NO") & ")"
'        End If
        DBConn.Execute sql
    Else
        If MsgBox("¿Desea Modificar la Consulta Medica?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        sql = "UPDATE CCLINICO SET "
        sql = sql & " CCL_FECHA = " & XDQ(Fecha.Value)
        sql = sql & " ,CLI_CODIGO=" & XN(txtCodigo.Text)
        sql = sql & " ,VEN_CODIGO=" & cboDocCon.ItemData(cboDocCon.ListIndex)
        sql = sql & " ,CCL_MOTIVO=" & XS(txtMotivo.Text)
        sql = sql & " ,CCL_INDICA=" & XS(txtIndicaciones.Text)
        sql = sql & " ,CCL_FECPC=" & XDQ(ChkNull(FechaProx.Value))
'        If optSI2 = True Then
'            sql = sql & " ,CCL_CONMUTUAL=" & XS("SI")
'        Else
'            sql = sql & " ,CCL_CONMUTUAL=" & XS("NO")
'        End If
        sql = sql & " WHERE CCL_NUMERO = " & txtnrocon.Text
        DBConn.Execute sql
    End If
       
    'DBConn.CommitTrans
        
    'cboDesde.ListIndex = cboDesde.ListIndex + 1
    'Next
    'cboDesde.Text = sHoraDAux
    'If MsgBox("¿Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'ImprimirTurno
    
    LimpiarConsulta
    
    CargarConsultasAnteriores
            
    'Exit Sub
    
'HayErrorTurno:
    'Screen.MousePointer = vbNormal
    'If rec.State = 1 Then rec.Close
    'If Rec1.State = 1 Then Rec1.Close
    'DBConn.RollbackTrans
    'MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
    'agregar columnas en la grilla, para guardar el codigo de doctor, paciente
End Sub

Private Sub cmdAceptarImg_Click()
    Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    Dim i As Integer
    Dim Num As Integer
    
    'Validar los campos requeridos
    If validarImagen = False Then Exit Sub
    If txtNroImg.Text = "" Then
        If MsgBox("¿Desea cargar los Datos de la Imágen?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        'agregar teniendo en cuentas loc combos de horas
        'On Error GoTo HayErrorTurno

        i = 0
        sql = "SELECT MAX(IMG_CODIGO) as ultimo FROM IMAGEN"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Num = Chk0(rec!Ultimo) + 1 ' guardo el proxuimo turno
        Else
            Num = 1
        End If
        rec.Close
        
        'NUEVA CONSULTA
        sql = "INSERT INTO IMAGEN"
        sql = sql & " (IMG_CODIGO,IMG_FECHA,"
        sql = sql & " CLI_CODIGO,VEN_CODIGO,TIP_CODIGO,IMG_DESCRI)"
        sql = sql & " VALUES ("
        sql = sql & Num & ","
        sql = sql & XDQ(FechaImg.Value) & ","
        sql = sql & XN(txtCodigo.Text) & ","
        sql = sql & cboDocImg.ItemData(cboDocImg.ListIndex) & ","
        sql = sql & cboImg.ItemData(cboImg.ListIndex) & ","
        sql = sql & XS(txtImgDescri.Text) & ")"
        DBConn.Execute sql
    Else
        If MsgBox("¿Desea Modificar la Imagen?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        sql = "UPDATE IMAGEN SET "
        sql = sql & " IMG_FECHA = " & XDQ(FechaImg.Value)
        sql = sql & " ,CLI_CODIGO=" & XN(txtCodigo.Text)
        sql = sql & " ,VEN_CODIGO=" & cboDocCon.ItemData(cboDocCon.ListIndex)
        sql = sql & " ,TIP_CODIGO=" & cboImg.ItemData(cboImg.ListIndex)
        sql = sql & " ,IMG_DESCRI=" & XS(txtImgDescri.Text)
        sql = sql & " WHERE IMG_CODIGO = " & txtNroImg.Text
        DBConn.Execute sql
    End If
       
    'DBConn.CommitTrans
        
    'cboDesde.ListIndex = cboDesde.ListIndex + 1
    'Next
    'cboDesde.Text = sHoraDAux
    'If MsgBox("¿Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'ImprimirTurno
    
    LimpiarImagen
    
    CargarImagenesAnteriores
End Sub

Private Sub cmdAceptarPedido_Click()
    Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    Dim i As Integer
    Dim Num As Integer
    
    'Validar los campos requeridos
    If validarPedido = False Then Exit Sub
    If txtnroPedido.Text = "" Then
        If MsgBox("¿Desea cargar el Pedido?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        'agregar teniendo en cuentas loc combos de horas
        'On Error GoTo HayErrorTurno

        i = 0
        sql = "SELECT MAX(PED_NUMERO) as ultimo FROM PEDIDO"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Num = rec!Ultimo + 1 ' guardo el proxuimo turno
        Else
            Num = 1
        End If
        rec.Close
        
        'NUEVO PEDIDO
        sql = "INSERT INTO PEDIDO"
        sql = sql & " (PED_NUMERO,PED_FECHA,"
        sql = sql & " ESP_CODIGO,CLI_CODIGO,PED_MOTIVO,PED_DESCRI,VEN_CODIGO)"
        sql = sql & " VALUES ("
        sql = sql & Num & ","
        sql = sql & XDQ(FechaPed.Value) & ","
        sql = sql & cboEspecPedido.ItemData(cboEspecPedido.ListIndex) & ","
        sql = sql & XN(txtCodigo.Text) & ","
        sql = sql & XS(txtMotivoPedido.Text) & ","
        sql = sql & XS(txtDescPedido.Text) & ","
        sql = sql & cboDocPedido.ItemData(cboDocPedido.ListIndex) & ")"
        DBConn.Execute sql
    Else
        If MsgBox("¿Desea Modificar el Pedido?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        sql = "UPDATE PEDIDO SET "
        sql = sql & " PED_FECHA = " & XDQ(FechaPed.Value)
        sql = sql & " ,CLI_CODIGO=" & XN(txtCodigo.Text)
        sql = sql & " ,ESP_CODIGO=" & cboEspecPedido.ItemData(cboEspecPedido.ListIndex)
        sql = sql & " ,PED_MOTIVO=" & XS(txtMotivoPedido.Text)
        sql = sql & " ,PED_DESCRI=" & XS(txtDescPedido.Text)
        sql = sql & "  WHERE PED_NUMERO = " & txtnroPedido.Text
        sql = sql & "  AND VEN_CODIGO=" & cboDocPedido.ItemData(cboDocPedido.ListIndex)
        DBConn.Execute sql
    End If
       
    'DBConn.CommitTrans
        
    'cboDesde.ListIndex = cboDesde.ListIndex + 1
    'Next
    'cboDesde.Text = sHoraDAux
    'If MsgBox("¿Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'ImprimirTurno
    
    LimpiarPedido
    CargarPedidosAnteriores
End Sub
Private Function LimpiarPedido()
    txtMotivoPedido = ""
    txtDescPedido = ""
    txtnroPedido = ""
    FechaPed.Value = Date
    'cboEspecPedido.ListIndex = 0
End Function
Private Function validarPedido() As Boolean
    validarPedido = True
    If txtBuscaCliente.Text = "" Then
        MsgBox "No ha ingresado el paciente", vbCritical, TIT_MSGBOX
        txtBuscaCliente.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If txtMotivoPedido.Text = "" Then
        MsgBox "No ha ingresado el motivo de pedido", vbCritical, TIT_MSGBOX
        txtMotivoPedido.SetFocus
        validarPedido = False
        Exit Function
    End If
    If txtDescPedido.Text = "" Then
        MsgBox "No ha ingresado la Descripción del pedido", vbCritical, TIT_MSGBOX
        txtDescPedido.SetFocus
        validarPedido = False
        Exit Function
    End If
    If cboEspecPedido.ListIndex < 1 Then
        MsgBox "No ha ingresado la especialidad", vbCritical, TIT_MSGBOX
        cboEspecPedido.SetFocus
        validarPedido = False
        Exit Function
    End If
    'If cboDocPedido.ListIndex = -1 Then
     '   MsgBox "No ha ingresado el doctor", vbCritical, TIT_MSGBOX
      '  cboDocPedido.SetFocus
      '  validarPedido = False
    'End If
End Function
Private Function CargarPedidosAnteriores()
    Dim sColor As String
    Dim USUARIO As String
    Dim Rec1 As New ADODB.Recordset
    sql = "SELECT P.*,V.VEN_NOMBRE,C.CLI_RAZSOC, E.ESP_DESCRI"
    sql = sql & " FROM PEDIDO P, CLIENTE C, VENDEDOR V, ESPECIALIDAD E "
    sql = sql & " WHERE P.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND P.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND P.ESP_CODIGO = E.ESP_CODIGO"
    If txtBuscaCliente.Text <> "" Then
        sql = sql & " AND P.CLI_CODIGO = " & XN(txtCodigo)
    End If
    If cboDocPedidos.ListIndex > 0 Then
        sql = sql & " AND P.VEN_CODIGO = " & cboDocPedidos.ItemData(cboDocPedidos.ListIndex)
    End If
    'PED_NUMERO|Cod cliente|VEN_CODIGO|Especialidad
    If FechaDesdePedido.Value <> "" Then sql = sql & " AND P.PED_FECHA>=" & XDQ(FechaDesdePedido.Value)
    If FechaHastaPedido.Value <> "" Then sql = sql & " AND P.PED_FECHA<=" & XDQ(FechaHastaPedido.Value)
    sql = sql & " ORDER BY PED_FECHA DESC"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdPedidos.Rows = 1
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            grdPedidos.AddItem Rec1!PED_FECHA & Chr(9) & Rec1!PED_DESCRI & Chr(9) & Rec1!PED_MOTIVO & Chr(9) & _
                                 Rec1!PED_NUMERO & Chr(9) & Rec1!CLI_CODIGO & Chr(9) & Rec1!VEN_CODIGO & Chr(9) & _
                                 Rec1!ESP_CODIGO & Chr(9) & Rec1!ESP_DESCRI

                                     
            
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
End Function
        

Private Sub cmdAnterior_Click()
''llamada aplicacion que de pacientes
'    Dim actual As Integer '1 paciente actuial 0 no es paciente actual
'    Dim horaactual As String '1 paciente actuial 0 no es paciente actual
'    If cboDocCon.ListIndex <> -1 Then
'        actual = 0
'
'        'Buscar en turno el anterior paciente
'        sql = "SELECT T.*, C.CLI_NROAFIL FROM CLIENTE C,TURNOS T "
'        sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO AND "
'        sql = sql & " T.VEN_CODIGO = " & cboDocCon.ItemData(cboDocCon.ListIndex)
'        sql = sql & " AND T.TUR_FECHA = " & XDQ(Fecha.Value)
'        sql = sql & " ORDER BY T.TUR_HORAD Desc "
'        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
'        If Rec2.EOF = False Then
'            Do While Rec2.EOF = False
'                If Rec2!CLI_CODIGO = txtCodigo Then
'                    'PACIENTE ACTUAL
'                    actual = 1
'                End If
'                If actual = 1 And Rec2!CLI_CODIGO <> txtCodigo Then
'                    txtCodigo = Rec2!CLI_CODIGO
'                    'txtMotivo.Text = Rec2!TUR_MOTIVO
'                    txthorad = Format(Rec2!TUR_HORAD, "hh:mm")
'                    TxtCodigo_LostFocus
'                    If Rec2!TUR_OSOCIAL = "PARTICULAR" Or IsNull(Rec2!TUR_OSOCIAL) Then 'para turnos cargados antes
'                        txtOSocial.Text = "PARTICULAR"
'                    Else
'                        txtOSocial.Text = Rec2!TUR_OSOCIAL & " - " & ChkNull(Rec2!CLI_NROAFIL)
'                    End If
'                    Rec2.Close
'                    Exit Sub
'                Else
'                    If actual = 1 And Rec2!CLI_CODIGO = txtCodigo And Format(Rec2!TUR_HORAD, "hh:mm") > txthorad Then
'                        txthorad = Format(Rec2!TUR_HORAD, "hh:mm")
'                        Rec2.Close
'                        Exit Sub
'                    End If
'                End If
'                Rec2.MovePrevious
'            Loop
'        End If
'        Rec2.Close
'    Else
'        MsgBox "Seleccione el Doctor", vbInformation, TIT_MSGBOX
'    End If

End Sub

Private Sub cmdCancelar_Click()
    'If MsgBox("¿Seguro desea Cancelarla Consulta Medica?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    LimpiarConsulta
    BuscaCodigoProxItemData Int(Doc), cboDocCon
End Sub

Private Sub cmdCancelarImg_Click()
    LimpiarImagen
    BuscaCodigoProxItemData Int(Doc), cboDocImg
End Sub

Private Sub cmdCancelarPedido_Click()
LimpiarPedido
End Sub

Private Sub cmdEcogra_Click()
    tabhc.Tab = 1
End Sub

Private Sub cmdFiltro_Click()
    If txtBuscaCliente.Text = "" Then
        MsgBox "Debe seleccionar un Paciente", vbInformation, TIT_MSGBOX
        grdConsultas.Rows = 1
        txtBuscaCliente.SetFocus
    Else
        CargarConsultasAnteriores
    End If
End Sub

Private Sub cmdFiltroPedidos_Click()
If txtBuscaCliente.Text = "" Then
        MsgBox "Debe seleccionar un Paciente", vbInformation, TIT_MSGBOX
        grdConsultas.Rows = 1
        txtBuscaCliente.SetFocus
    Else
        CargarPedidosAnteriores
    End If
End Sub

Private Sub cmdGineco_Click()
    tabhc.Tab = 3
End Sub

Private Sub cmdImprimirEco_Click()
    If txtNroImg.Text <> "" Then
        Rep.WindowState = crptMaximized
        Rep.WindowBorderStyle = crptNoBorder
        Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
        
        Rep.SelectionFormula = ""
        Rep.Formulas(0) = ""
                
        Rep.SelectionFormula = " {IMAGEN.IMG_CODIGO}= " & XN(txtNroImg.Text)
    
        Rep.WindowTitle = "Protocolos"
        Rep.ReportFileName = DirReport & "rptImagen.rpt"
        Rep.Action = 1
    '    lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        Rep.SelectionFormula = ""
    End If
End Sub

Private Sub cmdLabora_Click()
    tabhc.Tab = 2
End Sub

Private Sub CmdNuevo_Click()
    LimpiarConsulta
    grdConsultas.Rows = 1
    txtBuscaCliente.Text = ""
    txtBuscaCliente_LostFocus
    tabhc.Tab = 0
End Sub

Private Sub cmdPedidos_Click()
    tabhc.Tab = 4
End Sub

Private Sub cmdproximo_Click()

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command7_Click()

End Sub

Private Sub cmdSiguiente_Click()
'llamada aplicacion que de pacientes
    Dim actual As Integer '1 paciente actuial 0 no es paciente actual
    Dim horaactual As String '1 paciente actuial 0 no es paciente actual
    If cboDocCon.ListIndex <> -1 Then
        actual = 0
        'actualizo en BD que el paciente actual asistio
        'Actualizo la Base de Datos
        If MsgBox("¿Confirma la asistencia del paciente " & txtBuscarCliDescri & " ?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            sql = "UPDATE TURNOS SET "
            sql = sql & " TUR_ASISTIO = 1"
            sql = sql & " WHERE "
            sql = sql & " TUR_FECHA = " & XDQ(Date)
            sql = sql & " AND TUR_HORAD = #" & txthorad & "#"
            sql = sql & " AND VEN_CODIGO = " & XN(cboDocCon.ItemData(cboDocCon.ListIndex))
            sql = sql & " AND CLI_CODIGO = " & XN(txtCodigo)
            DBConn.Execute sql
        End If
        
        'Buscar en turno el proximo paciente
        sql = "SELECT T.*, C.CLI_NROAFIL FROM CLIENTE C,TURNOS T "
        sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO AND "
        sql = sql & " T.VEN_CODIGO = " & cboDocCon.ItemData(cboDocCon.ListIndex)
        sql = sql & " AND T.TUR_FECHA = " & XDQ(Fecha.Value)
        sql = sql & " ORDER BY T.TUR_HORAD ASC "
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            Do While Rec2.EOF = False
                If Rec2!CLI_CODIGO = txtCodigo Then
                    'PACIENTE ACTUAL
                    actual = 1
                End If
                If actual = 1 And Rec2!CLI_CODIGO <> txtCodigo Then
                    txtCodigo = Rec2!CLI_CODIGO
                    'txtMotivo.Text = Rec2!TUR_MOTIVO
                    txthorad = Format(Rec2!TUR_HORAD, "hh:mm")
                    TxtCodigo_LostFocus
                    If Rec2!TUR_OSOCIAL = "PARTICULAR" Or IsNull(Rec2!TUR_OSOCIAL) Then 'para turnos cargados antes
                        txtOSocial.Text = "PARTICULAR"
                    Else
                        txtOSocial.Text = Rec2!TUR_OSOCIAL & " - " & ChkNull(Rec2!CLI_NROAFIL)
        End If
                    Rec2.Close
                    Exit Sub
                Else
                    If actual = 1 And Rec2!CLI_CODIGO = txtCodigo And Format(Rec2!TUR_HORAD, "hh:mm") > txthorad Then
                        txthorad = Format(Rec2!TUR_HORAD, "hh:mm")
                        Rec2.Close
                        Exit Sub
                    End If
                End If
                Rec2.MoveNext
            Loop
        End If
        Rec2.Close
    Else
        MsgBox "Seleccione el Doctor", vbInformation, TIT_MSGBOX
    End If

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command5_Click()
    
End Sub

Private Sub Command6_Click()
     If txtBuscaCliente.Text = "" Then
        MsgBox "Debe seleccionar un Paciente", vbInformation, TIT_MSGBOX
        grdConsultas.Rows = 1
        txtBuscaCliente.SetFocus
    Else
        CargarImagenesAnteriores
    End If
End Sub

Private Sub Form_Activate()
    If txtCodigo <> "" Then
        TxtCodigo_LostFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If
End Sub

Private Sub Form_Load()
    Dim CodCli As Integer
    preparogrillas
    cargocombos
    Fecha.Value = Date
    FechaProx.Value = ""
    tabhc.Tab = 0
    CodCli = BuscarProxPaciente(Int(Doc), XDQ(Fecha.Value))
    If CodCli <> 0 Then
        txtBuscaCliente.Text = ChkNull(CodCli)
        txtCodigo.Text = ChkNull(CodCli)
        txtBuscaCliente_LostFocus
    End If
    'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
    If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then 'VEERR
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub
Private Function HabilitarBoton(boton As String)
'    If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then 'VEERR
'        boton.Enabled = True
'    Else
'        boton.Enabled = False
'    End If
End Function


Private Function preparogrillas()
    ' Grilla de Curso Clinico - Consulta de Historia Clinica
    grdConsultas.FormatString = "Fecha|Doctor|Motivo|Indicaciones|FechaProx|CodMedico|CCL_NUMERO|CCL_CONMUTUAL"
    grdConsultas.ColWidth(0) = 1500  'Fecha
    grdConsultas.ColWidth(1) = 2500 'Doctor
    grdConsultas.ColWidth(2) = 3500 'Motivo
    grdConsultas.ColWidth(3) = 0 'Indicaciones
    grdConsultas.ColWidth(4) = 0 'Fecha Proxima
    grdConsultas.ColWidth(5) = 0 'CodMedico
    grdConsultas.ColWidth(6) = 0 'CCL_NUMERO
    grdConsultas.ColWidth(7) = 0 'CCL_CONMUTUAL
    grdConsultas.Rows = 1
    grdConsultas.BorderStyle = flexBorderNone
    grdConsultas.row = 0
    For i = 0 To grdConsultas.Cols - 1
        grdConsultas.Col = i
        grdConsultas.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdConsultas.CellBackColor = &H808080    'GRIS OSCURO
        grdConsultas.CellFontBold = True
    Next
    grdConsultas.HighLight = flexHighlightAlways

    
    ' Grilla de pedidos
    grdPedidos.FormatString = "Fecha|Descripcion|Motivo|PED_NUMERO|Cod cliente|VEN_CODIGO|EspecialidadCOD|Especialidad"
    grdPedidos.ColWidth(0) = 1500  'Fecha
    grdPedidos.ColWidth(1) = 3500 'Descripcion
    grdPedidos.ColWidth(2) = 0 'Motivo
    grdPedidos.ColWidth(3) = 0 'num ped
    grdPedidos.ColWidth(4) = 0 'Cod cliente
    grdPedidos.ColWidth(5) = 0 'Cod vendedor
    grdPedidos.ColWidth(6) = 0 'especialidad COD
    grdPedidos.ColWidth(7) = 2000 'especialidad
    grdPedidos.Rows = 1
    grdPedidos.BorderStyle = flexBorderNone
    grdPedidos.row = 0
    For h = 0 To grdPedidos.Cols - 1
        grdPedidos.Col = h
        grdPedidos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdPedidos.CellBackColor = &H808080    'GRIS OSCURO
        grdPedidos.CellFontBold = True
    Next
        grdPedidos.HighLight = flexHighlightAlways
    
    ' Grilla de IMAGENES -
    grdImagenes.FormatString = "Fecha|Doctor|Imágen|Descripcion|TipoIMG|CodMedico|IMG_CODIGO"
    grdImagenes.ColWidth(0) = 1500  'Fecha
    grdImagenes.ColWidth(1) = 2500 'Doctor
    grdImagenes.ColWidth(2) = 3500 'Imagen
    grdImagenes.ColWidth(3) = 0 'descripcion
    grdImagenes.ColWidth(4) = 0 'tipo Img
    grdImagenes.ColWidth(5) = 0 'CodMedico
    grdImagenes.ColWidth(6) = 0 'IMG CODIGO
    grdImagenes.Rows = 1
    grdImagenes.BorderStyle = flexBorderNone
    grdImagenes.row = 0
    For i = 0 To grdImagenes.Cols - 1
        grdImagenes.Col = i
        grdImagenes.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdImagenes.CellBackColor = &H808080    'GRIS OSCURO
        grdImagenes.CellFontBold = True
    Next
    grdImagenes.HighLight = flexHighlightAlways
    
End Function
Private Function cargocombos()
    sql = "SELECT * FROM VENDEDOR"
    'cargar doctores que no sea la secretaria
    sql = sql & " WHERE PR_CODIGO > 1"
    sql = sql & " ORDER BY VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocAnt.AddItem ""
        cboDocPedidos.AddItem ""
        Do While rec.EOF = False
            cboDocCon.AddItem rec!VEN_NOMBRE
            cboDocCon.ItemData(cboDocCon.NewIndex) = rec!VEN_CODIGO
                                   
            cboDocPedido.AddItem rec!VEN_NOMBRE
            cboDocPedido.ItemData(cboDocPedido.NewIndex) = rec!VEN_CODIGO
            
            cboDocAnt.AddItem rec!VEN_NOMBRE
            cboDocAnt.ItemData(cboDocAnt.NewIndex) = rec!VEN_CODIGO
                        
            cboDocPedidos.AddItem rec!VEN_NOMBRE
            cboDocPedidos.ItemData(cboDocPedidos.NewIndex) = rec!VEN_CODIGO
            
            'IMAGENES - DOCTORES
            cboDocImg.AddItem rec!VEN_NOMBRE
            cboDocImg.ItemData(cboDocImg.NewIndex) = rec!VEN_CODIGO
            
            cboDocImgAnt.AddItem rec!VEN_NOMBRE
            cboDocImgAnt.ItemData(cboDocImgAnt.NewIndex) = rec!VEN_CODIGO
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    'BUSCO NOMBRES DE IMAGENES
    sql = "SELECT * FROM TIPO_IMAGEN"
    sql = sql & " WHERE TIP_CODIGO > 0"
    sql = sql & " ORDER BY TIP_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboImgAnt.AddItem ""
        Do While rec.EOF = False
            cboImg.AddItem rec!TIP_NOMBRE
            cboImg.ItemData(cboImg.NewIndex) = rec!TIP_CODIGO
            'imagenes anteriores
            cboImgAnt.AddItem rec!TIP_NOMBRE
            cboImgAnt.ItemData(cboImg.NewIndex) = rec!TIP_CODIGO
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    'BUSCO CODIGO DE DOCTOR POR NOMBRE DE USUARIO
    sql = "SELECT VEN_CODIGO FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO > 1 "
    If mNomUser = "A" Or mNomUser = "DIGOR" Then
        sql = sql & " AND VEN_NOMBRE LIKE '" & "SILVANA" & "%'"
    Else
        sql = sql & " AND VEN_NOMBRE LIKE '" & mNomUser & "%'"
    End If
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Doc = rec!VEN_CODIGO
    End If
    rec.Close
    
    'pone el doctor por defecto que esta guardado en el archivo de configuracion C:/Windows/DIGOR.ini
    BuscaCodigoProxItemData Int(Doc), cboDocCon
    BuscaCodigoProxItemData Int(Doc), cboDocImg

    sql = "SELECT * FROM ESPECIALIDAD"
    sql = sql & " WHERE ESP_CODIGO > 0"
    sql = sql & " ORDER BY ESP_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
       cboEspecPedido.AddItem ""
        Do While rec.EOF = False
            cboEspecPedido.AddItem rec!ESP_DESCRI
            cboEspecPedido.ItemData(cboEspecPedido.NewIndex) = rec!ESP_CODIGO
           rec.MoveNext
        Loop
    End If
    rec.Close
End Function

Private Sub grdConsultas_Click()
    If grdConsultas.Rows > 1 Then
        Fecha.Value = grdConsultas.TextMatrix(grdConsultas.RowSel, 0)
        BuscaCodigoProxItemData grdConsultas.TextMatrix(grdConsultas.RowSel, 5), cboDocCon
        'cboDocCon.ListIndex = grdConsultas.TextMatrix(grdConsultas.RowSel, 5)
        txtMotivo = grdConsultas.TextMatrix(grdConsultas.RowSel, 2)
        txtIndicaciones = grdConsultas.TextMatrix(grdConsultas.RowSel, 3)
        FechaProx.Value = grdConsultas.TextMatrix(grdConsultas.RowSel, 4)
        txtnrocon.Text = grdConsultas.TextMatrix(grdConsultas.RowSel, 6)
         'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
        If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then
            cmdAceptar.Enabled = True
        Else
            cmdAceptar.Enabled = False
        End If
    End If
    
End Sub

Private Sub grdImagenes_Click()
    If grdImagenes.Rows > 1 Then
        FechaImg.Value = grdImagenes.TextMatrix(grdImagenes.RowSel, 0)
        BuscaCodigoProxItemData grdImagenes.TextMatrix(grdImagenes.RowSel, 5), cboDocImg
        'codigo del nombre de la imagen
        BuscaCodigoProxItemData grdImagenes.TextMatrix(grdImagenes.RowSel, 4), cboImg
        'cboDocImg.ListIndex = grdImagenes.TextMatrix(grdImagenes.RowSel, 5)
        txtImgDescri = grdImagenes.TextMatrix(grdImagenes.RowSel, 3)
        txtNroImg.Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 6)
         'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
        If cboDocImg.ItemData(cboDocImg.ListIndex) = Int(Doc) Then
            cmdAceptarImg.Enabled = True
        Else
            cmdAceptarImg.Enabled = False
        End If
    End If
End Sub

Private Sub grdPedidos_Click()
    If grdPedidos.Rows > 1 Then
        FechaPed.Value = grdPedidos.TextMatrix(grdPedidos.RowSel, 0)
        BuscaCodigoProxItemData grdPedidos.TextMatrix(grdPedidos.RowSel, 5), cboDocPedido
        'cboDocCon.ListIndex = grdConsultas.TextMatrix(grdConsultas.RowSel, 5)
        txtMotivoPedido = grdPedidos.TextMatrix(grdPedidos.RowSel, 2)
        txtDescPedido = grdPedidos.TextMatrix(grdPedidos.RowSel, 1)
        txtnroPedido.Text = grdPedidos.TextMatrix(grdPedidos.RowSel, 3)
        BuscaCodigoProxItemData grdPedidos.TextMatrix(grdPedidos.RowSel, 6), cboEspecPedido
        
        'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
        If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then
            cmdAceptar.Enabled = True
        Else
            cmdAceptar.Enabled = False
        End If
    End If
End Sub

Private Sub optSI_Click()

End Sub

Private Sub optNO2_Click()
    'txtNroAfil.Text = ""
End Sub

Private Sub optSI2_Click()
    'txtNroAfil.Text = txtNAfil.Text
End Sub

Private Sub txtBuscaCliente_Change()
    grdConsultas.Rows = 1
    grdPedidos.Rows = 1
        If txtBuscaCliente.Text = "" Then
            txtBuscarCliDescri.Text = ""
            txtCodigo.Text = ""
            txtTelefono.Text = ""
            txtOSocial.Text = ""
            txthorad.Text = ""
            txtEdad = ""
        End If
        If Len(Trim(txtBuscaCliente.Text)) < 7 Then
            txtBuscaCliente.ToolTipText = "Numero de Paciente"
        Else
            txtBuscaCliente.ToolTipText = "DNI"
        End If
    LimpiarConsulta
    LimpiarPedido
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
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,CLI_NROAFIL,CLI_CUMPLE"
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
            'txtBuscaCliente.Text = rec!CLI_NRODOC
            txtBuscarCliDescri.Text = rec!CLI_RAZSOC
            txtCodigo.Text = rec!CLI_CODIGO
            txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
            'ATENCION CON O SIN OBRA SOCIAL
            If TurOSocial = "PARTICULAR" Then
                txtOSocial.Text = "PARTICULAR"
            Else
                txtOSocial.Text = TurOSocial & " - " & ChkNull(rec!CLI_NROAFIL) 'BuscarOSocial(rec!CLI_CODIGO)
            End If
            'txtNAfil.Text = ChkNull(rec!CLI_NROAFIL)
            'calculo de edad
            Calculo_Edad Chk0(rec!CLI_CUMPLE)
            
            CargarConsultasAnteriores
            CargarPedidosAnteriores
            CargarImagenesAnteriores
            'txtMotivo.SetFocus
            'ActivoGrid = 1
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
           ' txtBuscaCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtBuscarCliDescri_Change()
    If txtBuscarCliDescri.Text = "" Then
        txtBuscaCliente.Text = ""
        txtCodigo.Text = ""
        txtTelefono.Text = ""
        txtOSocial.Text = ""
        txtEdad = ""
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

Private Sub txtBuscarCliDescri_LostFocus()
    If txtBuscaCliente.Text = "" And txtBuscarCliDescri.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,CLI_CUMPLE"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            If Len(Trim(txtBuscaCliente.Text)) < 7 Then
                sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
            Else
                sql = sql & " CLI_NRODOC=" & XN(txtBuscaCliente)
            End If
            'sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtBuscarCliDescri) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtBuscaCliente", "CADENA", Trim(txtBuscarCliDescri.Text)
                If rec.State = 1 Then rec.Close
                txtBuscarCliDescri.SetFocus
            Else
                'txtBuscaCliente.Text = rec!CLI_DNI
                If Len(Trim(txtBuscaCliente.Text)) < 7 Then
                    txtBuscaCliente.Text = rec!CLI_CODIGO
                Else
                    txtBuscaCliente.Text = rec!CLI_NRODOC
                End If
                'txtBuscaCliente.Text = rec!CLI_NRODOC
                txtBuscarCliDescri.Text = rec!CLI_RAZSOC
                txtCodigo.Text = rec!CLI_CODIGO
                txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
                Calculo_Edad ChkNull(rec!CLI_CUMPLE)
            End If
            'ActivoGrid = 0
        Else
            MsgBox "No se encontro el Paciente", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
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
Private Function BuscarOSocial(CodCli As Integer) As String
    Set Rec1 = New ADODB.Recordset
    sql = "SELECT O.OS_NOMBRE FROM OBRA_SOCIAL O, CLIENTE C"
    sql = sql & " WHERE C.OS_NUMERO = O.OS_NUMERO"
    sql = sql & " AND C.CLI_CODIGO = " & CodCli
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarOSocial = Rec1!OS_NOMBRE
    Else
        BuscarOSocial = ""
    End If
    Rec1.Close
End Function
Private Function LimpiarConsulta()
    Fecha.Value = Date
    txtMotivo = ""
    txtIndicaciones = ""
    
    'cboDocCon.ListIndex = 0
    FechaProx.Value = ""
    txtnrocon = ""
End Function

Private Function CargarConsultasAnteriores()
    Dim sColor As String
    Dim USUARIO As String
    Dim Rec1 As New ADODB.Recordset
    sql = "SELECT CC.*,V.VEN_NOMBRE,C.CLI_RAZSOC"
    sql = sql & " FROM CCLINICO CC, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE CC.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND CC.VEN_CODIGO = V.VEN_CODIGO"
    If txtBuscaCliente.Text <> "" Then
        sql = sql & " AND CC.CLI_CODIGO = " & XN(txtCodigo)
    End If
    If cboDocAnt.ListIndex > 0 Then
        sql = sql & " AND CC.VEN_CODIGO = " & cboDocAnt.ItemData(cboDocAnt.ListIndex)
    End If
    If FechaDesde.Value <> "" Then sql = sql & " AND CC.CCL_FECHA>=" & XDQ(FechaDesde.Value)
    If FechaHasta.Value <> "" Then sql = sql & " AND CC.CCL_FECHA<=" & XDQ(FechaHasta.Value)
    sql = sql & " ORDER BY CCL_FECHA DESC"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdConsultas.Rows = 1
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            grdConsultas.AddItem Rec1!CCL_FECHA & Chr(9) & Rec1!VEN_NOMBRE & Chr(9) & Rec1!CCL_MOTIVO & Chr(9) & _
                                 Rec1!CCL_INDICA & Chr(9) & Rec1!CCL_FECPC & Chr(9) & Rec1!VEN_CODIGO & Chr(9) & Rec1!CCL_NUMERO & Chr(9)
                                       
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
    LimpiarConsulta
End Function
Private Function CargarImagenesAnteriores()
    Dim sColor As String
    Dim USUARIO As String
    Dim Rec1 As New ADODB.Recordset
    sql = "SELECT I.*,V.VEN_NOMBRE,C.CLI_RAZSOC, T.TIP_NOMBRE"
    sql = sql & " FROM IMAGEN I, VENDEDOR V, CLIENTE C, TIPO_IMAGEN T"
    sql = sql & " WHERE I.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND I.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND I.TIP_CODIGO = T.TIP_CODIGO"
    If txtBuscaCliente.Text <> "" Then
        sql = sql & " AND I.CLI_CODIGO = " & XN(txtCodigo)
    End If
    If cboDocImgAnt.ListIndex > -1 Then
        sql = sql & " AND I.VEN_CODIGO = " & cboDocImgAnt.ItemData(cboDocImgAnt.ListIndex)
    End If
    If cboImgAnt.ListIndex > -1 Then
        sql = sql & " AND I.TIP_CODIGO = " & cboImgAnt.ItemData(cboImgAnt.ListIndex - 1)
    End If
    If FechaDesdeImg.Value <> "" Then sql = sql & " AND I.IMG_FECHA>=" & XDQ(FechaDesdeImg.Value)
    If FechaHastaImg.Value <> "" Then sql = sql & " AND I.IMG_FECHA<=" & XDQ(FechaHastaImg.Value)
    sql = sql & " ORDER BY IMG_FECHA DESC"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdImagenes.Rows = 1
    If Rec1.EOF = False Then
        Do While Rec1.EOF = False
            grdImagenes.AddItem Rec1!IMG_FECHA & Chr(9) & Rec1!VEN_NOMBRE & Chr(9) & Rec1!TIP_NOMBRE & Chr(9) & Rec1!IMG_DESCRI & Chr(9) & _
                                    Rec1!TIP_CODIGO & Chr(9) & Rec1!VEN_CODIGO & Chr(9) & Rec1!IMG_CODIGO
                                       
            Rec1.MoveNext
        Loop
    End If
    Rec1.Close
    LimpiarImagen
End Function

Private Sub TxtCodigo_LostFocus()
    Dim edad As Integer
    Dim años As Integer
    If txtCodigo.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,CLI_NROAFIL,CLI_CUMPLE"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        sql = sql & " CLI_CODIGO=" & XN(txtCodigo)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtBuscaCliente.Text = IIf(IsNull(rec!CLI_NRODOC), rec!CLI_CODIGO, rec!CLI_NRODOC)
            txtBuscarCliDescri.Text = rec!CLI_RAZSOC
            txtCodigo.Text = rec!CLI_CODIGO
            txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
            txtNAfil.Text = ChkNull(rec!CLI_NROAFIL)
            txtOSocial.Text = BuscarOSocial(rec!CLI_CODIGO) & " - " & ChkNull(rec!CLI_NROAFIL)
            'calculo de edad
            'BuscarProxPaciente
            Calculo_Edad IIf(IsNull(rec!CLI_CUMPLE), Date, rec!CLI_CUMPLE)
            CargarConsultasAnteriores
            'txtMotivo.SetFocus
            'ActivoGrid = 1
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
        End If
        rec.Close
    End If
   txtBuscaCliente_LostFocus
End Sub

