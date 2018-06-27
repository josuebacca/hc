VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmhistoriaclinica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia Clinica"
   ClientHeight    =   10935
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   16635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   16635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAgregarPedido 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   13080
      TabIndex        =   35
      Top             =   10080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   15240
      TabIndex        =   34
      Top             =   10080
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   14160
      TabIndex        =   33
      Top             =   10080
      Width           =   1095
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
         TabIndex        =   26
         Tag             =   "Descripción"
         Top             =   360
         Width           =   555
      End
      Begin VB.TextBox txtNAfil 
         Height          =   285
         Left            =   5160
         TabIndex        =   112
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txthorad 
         Height          =   285
         Left            =   3960
         TabIndex        =   110
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   3240
         TabIndex        =   32
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
         TabIndex        =   27
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   113
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   360
         Width           =   1320
      End
   End
   Begin TabDlg.SSTab tabhc 
      Height          =   8775
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   15478
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
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "cmdEliminarEco"
      Tab(1).Control(2)=   "cmdAgregarEco"
      Tab(1).Control(3)=   "cmdVer"
      Tab(1).Control(4)=   "cmdVerEstudio"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Laboratorio"
      TabPicture(2)   =   "frmhistoriaclinica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Ginecologia"
      TabPicture(3)   =   "frmhistoriaclinica.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(1)=   "optTipoEst"
      Tab(3).Control(2)=   "optFechaGine"
      Tab(3).Control(3)=   "cboTipoEstGine"
      Tab(3).Control(4)=   "Frame3"
      Tab(3).Control(5)=   "cmdAgregarEstGine"
      Tab(3).Control(6)=   "cmdEliminarEstGine"
      Tab(3).Control(7)=   "cmdImprimirEstGine"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Pedidos"
      TabPicture(4)   =   "frmhistoriaclinica.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).Control(1)=   "Frame9"
      Tab(4).ControlCount=   2
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
         TabIndex        =   91
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
            TabIndex        =   106
            Top             =   375
            Width           =   3975
         End
         Begin VB.CommandButton cmdFiltroPedidos 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   109
            Top             =   360
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesdePedido 
            Height          =   315
            Left            =   2025
            TabIndex        =   107
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHastaPedido 
            Height          =   315
            Left            =   4575
            TabIndex        =   108
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdPedidos 
            Height          =   6375
            Left            =   120
            TabIndex        =   92
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
            TabIndex        =   95
            Top             =   435
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   2
            Left            =   3600
            TabIndex        =   94
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   840
            TabIndex        =   93
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
         TabIndex        =   82
         Top             =   480
         Width           =   8055
         Begin VB.TextBox txtnroPedido 
            Height          =   315
            Left            =   2880
            TabIndex        =   97
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
            ItemData        =   "frmhistoriaclinica.frx":008C
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":008E
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1920
            Width           =   2220
         End
         Begin VB.CommandButton cmdImprimirPedido 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   4560
            TabIndex        =   103
            Top             =   7680
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarPedido 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   104
            Top             =   7680
            Width           =   1095
         End
         Begin VB.TextBox txtMotivoPedido 
            Height          =   315
            Left            =   1305
            TabIndex        =   100
            Top             =   1440
            Width           =   6375
         End
         Begin VB.TextBox txtDescPedido 
            Height          =   5115
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   102
            Top             =   2400
            Width           =   6375
         End
         Begin VB.CommandButton cmdCancelarPedido 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   105
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
            ItemData        =   "frmhistoriaclinica.frx":0090
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":0092
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   1080
            Width           =   2220
         End
         Begin VB.TextBox txtConsultorioPedido 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   84
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtProfesionPedido 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   83
            Top             =   1080
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker FechaPed 
            Height          =   315
            Left            =   1305
            TabIndex        =   98
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   41098
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   99
            Top             =   660
            Width           =   495
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Especialidad:"
            Height          =   195
            Left            =   240
            TabIndex        =   96
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   240
            TabIndex        =   90
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   89
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   240
            TabIndex        =   88
            Top             =   1140
            Width           =   540
         End
         Begin VB.Label Label20 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3720
            TabIndex        =   87
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6240
            TabIndex        =   86
            Top             =   1110
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   120
         TabIndex        =   55
         Top             =   7500
         Width           =   8055
         Begin VB.CommandButton cmdSiguiente 
            Caption         =   "&Siguiente Paciente"
            Height          =   855
            Left            =   480
            TabIndex        =   81
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdLabora 
            Caption         =   "&Laboratorio"
            Height          =   855
            Left            =   4080
            TabIndex        =   62
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdEcogra 
            Caption         =   "&Ecografias"
            Height          =   855
            Left            =   2880
            TabIndex        =   61
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdGineco 
            Caption         =   "&Ginecologia"
            Height          =   855
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAnterior 
            Caption         =   "&Anterior Paciente"
            Height          =   855
            Left            =   1680
            TabIndex        =   59
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdPedidos 
            Caption         =   "&Pedidos"
            Height          =   855
            Left            =   6480
            TabIndex        =   58
            Top             =   120
            Width           =   1215
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
         Height          =   7095
         Left            =   120
         TabIndex        =   43
         Top             =   420
         Width           =   8055
         Begin VB.TextBox txtProfesion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   51
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtConsultorio 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   52
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtnrocon 
            Height          =   315
            Left            =   2880
            TabIndex        =   66
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComCtl2.DTPicker FechaProx 
            Height          =   315
            Left            =   2040
            TabIndex        =   57
            Top             =   6360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   43205
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   48
            Top             =   6360
            Width           =   1095
         End
         Begin VB.TextBox txtIndicaciones 
            Height          =   4035
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   2040
            Width           =   6375
         End
         Begin VB.TextBox txtMotivo 
            Height          =   315
            Left            =   1305
            TabIndex        =   54
            Top             =   1440
            Width           =   6375
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   1305
            TabIndex        =   45
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   41098
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   49
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
            ItemData        =   "frmhistoriaclinica.frx":0094
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":0096
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1080
            Width           =   2220
         End
         Begin VB.Label Label18 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6240
            TabIndex        =   80
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3720
            TabIndex        =   79
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Próxima Consulta:"
            Height          =   375
            Left            =   360
            TabIndex        =   65
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   600
            TabIndex        =   53
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Indicaciones:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   600
            TabIndex        =   46
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   600
            TabIndex        =   44
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
         Height          =   8175
         Left            =   8280
         TabIndex        =   36
         Top             =   420
         Width           =   8055
         Begin VB.CommandButton cmdFiltro 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   64
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
            TabIndex        =   37
            Top             =   375
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2025
            TabIndex        =   39
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   4575
            TabIndex        =   40
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   20971521
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdConsultas 
            Height          =   6870
            Left            =   120
            TabIndex        =   63
            Top             =   1200
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   12118
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
            TabIndex        =   42
            Top             =   795
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   0
            Left            =   3600
            TabIndex        =   41
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   1290
            TabIndex        =   38
            Top             =   435
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdImprimirEstGine 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -57480
         TabIndex        =   22
         Top             =   9060
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEstGine 
         Caption         =   "Eliminar Estudio"
         Height          =   375
         Left            =   -59400
         TabIndex        =   21
         Top             =   9060
         Width           =   1575
      End
      Begin VB.CommandButton cmdAgregarEstGine 
         Caption         =   "Agregar Estudio"
         Height          =   375
         Left            =   -61080
         TabIndex        =   20
         Top             =   9060
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estudios"
         Height          =   7095
         Left            =   -74880
         TabIndex        =   18
         Top             =   1980
         Width           =   19095
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   6615
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   18855
            _ExtentX        =   33258
            _ExtentY        =   11668
            _Version        =   393216
         End
      End
      Begin VB.ComboBox cboTipoEstGine 
         Height          =   315
         Left            =   -73320
         TabIndex        =   17
         Text            =   "Combo2"
         Top             =   1380
         Width           =   2775
      End
      Begin VB.OptionButton optFechaGine 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   -68400
         TabIndex        =   16
         Top             =   900
         Width           =   2055
      End
      Begin VB.OptionButton optTipoEst 
         Caption         =   "Tipo Estudio"
         Height          =   375
         Left            =   -73320
         TabIndex        =   15
         Top             =   900
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdVerEstudio 
         Caption         =   "Ver Estudio(VA ACA?)"
         Height          =   375
         Left            =   -60360
         TabIndex        =   13
         Top             =   8580
         Width           =   1335
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -61800
         TabIndex        =   12
         Top             =   8580
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregarEco 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -58800
         TabIndex        =   11
         Top             =   9060
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEco 
         Caption         =   "Elinimar"
         Height          =   375
         Left            =   -57000
         TabIndex        =   10
         Top             =   9060
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ecografías"
         Height          =   8055
         Left            =   -74880
         TabIndex        =   2
         Top             =   780
         Width           =   16215
         Begin VB.Frame Frame7 
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
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   8295
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   7680
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton cmdAbrir 
               Height          =   495
               Left            =   7680
               Picture         =   "frmhistoriaclinica.frx":0098
               Style           =   1  'Graphical
               TabIndex        =   78
               ToolTipText     =   "Abrir desde un archivo existente"
               Top             =   1080
               Width           =   495
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Aceptar"
               Height          =   375
               Left            =   5640
               TabIndex        =   74
               Top             =   6360
               Width           =   1095
            End
            Begin VB.TextBox txteco 
               Height          =   5115
               Left            =   1320
               MultiLine       =   -1  'True
               TabIndex        =   72
               Top             =   1080
               Width           =   6375
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   6720
               TabIndex        =   71
               Top             =   6360
               Width           =   1095
            End
            Begin VB.ComboBox Combo2 
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
               TabIndex        =   70
               Top             =   600
               Width           =   3180
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   2880
               TabIndex        =   69
               Top             =   600
               Visible         =   0   'False
               Width           =   495
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   1305
               TabIndex        =   73
               Top             =   600
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   20971521
               CurrentDate     =   41098
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Fecha:"
               Height          =   195
               Left            =   690
               TabIndex        =   77
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Indicaciones:"
               Height          =   675
               Left            =   240
               TabIndex        =   76
               Top             =   1080
               Width           =   945
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Doctor:"
               Height          =   195
               Left            =   3840
               TabIndex        =   75
               Top             =   660
               Width           =   540
            End
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   8280
            TabIndex        =   9
            Text            =   "Doctor"
            Top             =   240
            Width           =   2000
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   5040
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   240
            Width           =   2000
         End
         Begin VB.ComboBox cboEspecialidad 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Text            =   "Especialidad"
            Top             =   240
            Width           =   2000
         End
         Begin VB.OptionButton optEmpleado 
            Caption         =   "Doctor"
            Height          =   255
            Left            =   7080
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optFecha 
            Caption         =   "Fecha"
            Height          =   375
            Left            =   4080
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optEspecialidad 
            Caption         =   "Especialidad"
            Height          =   255
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   6510
            Left            =   8760
            TabIndex        =   67
            Top             =   960
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   11483
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
         Begin VB.Label Label6 
            Caption         =   "Buscar por:"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Buscar por:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   14
         Top             =   900
         Width           =   1095
      End
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Nro Carnet:"
      Height          =   195
      Left            =   5640
      TabIndex        =   111
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
    If txtIndicaciones.Text = "" Then
        MsgBox "No ha ingresado la indicación", vbCritical, TIT_MSGBOX
        txtMotivo.SetFocus
        validarcclinico = False
        Exit Function
    End If
    If txtMotivo.Text = "" Then
        MsgBox "No ha ingresado el motivo", vbCritical, TIT_MSGBOX
        txtMotivo.SetFocus
        validarcclinico = False
        Exit Function
    End If
    If cboDocCon.ListIndex = -1 Then
        MsgBox "No ha ingresado el doctor", vbCritical, TIT_MSGBOX
        cboDesde.SetFocus
        validarcclinico = False
        Exit Function
    End If
    If Fecha.Value = -1 Then
        MsgBox "No ha ingresado la fecha de la consulta", vbCritical, TIT_MSGBOX
        cbohasta.SetFocus
        validarcclinico = False
        Exit Function
    End If
    
    validarcclinico = True
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
        txtConsultorio.Text = rec!VEN_CONSULTORIO
        pro = rec!PR_CODIGO
        'defino profesion y consultorio en pedido
        txtConsultorioPedido.Text = rec!VEN_CONSULTORIO
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
    
    
    

End Sub

Private Sub cboDocPedido_Click()
sql = "SELECT PR_CODIGO, VEN_CONSULTORIO FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO = "
    sql = sql & cboDocPedido.ItemData(cboDocPedido.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtConsultorioPedido.Text = rec!VEN_CONSULTORIO
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
    If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
    
End Sub
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
            
            cboDocAnt.AddItem rec!VEN_NOMBRE
            cboDocAnt.ItemData(cboDocAnt.NewIndex) = rec!VEN_CODIGO
            
            cboDocPedido.AddItem rec!VEN_NOMBRE
            cboDocPedido.ItemData(cboDocPedido.NewIndex) = rec!VEN_CODIGO
            
            cboDocPedidos.AddItem rec!VEN_NOMBRE
            cboDocPedidos.ItemData(cboDocPedidos.NewIndex) = rec!VEN_CODIGO
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    'pone el doctor por defecto que esta guardado en el archivo de configuracion C:/Windows/DIGOR.ini
    BuscaCodigoProxItemData Int(Doc), cboDocCon
            
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
'        If grdConsultas.TextMatrix(grdConsultas.RowSel, 7) = "NO" Or IsNull(grdConsultas.TextMatrix(grdConsultas.RowSel, 7)) Then 'para turnos cargados antes
'                        optNO2.Value = True
'                        optSI2.Enabled = False
'                    Else
'                        optNO2.Enabled = False
'                        optSI2.Value = True
'        End If
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
                txtOSocial.Text = TurOSocial & " - " & NroAfil 'BuscarOSocial(rec!CLI_CODIGO)
            End If
            'txtNAfil.Text = ChkNull(rec!CLI_NROAFIL)
            'calculo de edad
            Calculo_Edad Chk0(rec!CLI_CUMPLE)
            
            CargarConsultasAnteriores
            CargarPedidosAnteriores
            'txtMotivo.SetFocus
            'ActivoGrid = 1
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
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
    Dim b As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set b = New CBusqueda
    With b
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
    
    Set b = Nothing
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
    'optSI2.Enabled = True
    'optNO2.Enabled = True
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
            txtOSocial.Text = BuscarOSocial(rec!CLI_CODIGO)
            txtNAfil.Text = ChkNull(rec!CLI_NROAFIL)
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
End Sub

