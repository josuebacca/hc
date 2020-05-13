VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmhistoriaclinica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia Clinica"
   ClientHeight    =   10500
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   16860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   16860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdzoom_out 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   16320
      TabIndex        =   123
      ToolTipText     =   "Zoom --"
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdzoom_out 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   16320
      TabIndex        =   84
      ToolTipText     =   "Zoom --"
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtindicaciones_zoom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   122
      Top             =   10320
      Visible         =   0   'False
      Width           =   16575
   End
   Begin VB.Frame fraprotocolos 
      Caption         =   "Protocolos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   8880
      TabIndex        =   115
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdAceptarP 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   4320
         TabIndex        =   119
         Top             =   7920
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalirP 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   5760
         TabIndex        =   117
         Top             =   7920
         Width           =   1455
      End
      Begin VB.TextBox txtfiltrop 
         Height          =   315
         Left            =   1800
         TabIndex        =   116
         Top             =   240
         Width           =   3855
      End
      Begin MSFlexGridLib.MSFlexGrid grdProtocolos 
         Height          =   7230
         Left            =   120
         TabIndex        =   118
         Top             =   600
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   12753
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Filtro"
         Height          =   195
         Left            =   1320
         TabIndex        =   120
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdAgregarPedido 
      Caption         =   "Agregar"
      Height          =   735
      Left            =   12960
      TabIndex        =   16
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   15120
      TabIndex        =   15
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   14040
      TabIndex        =   14
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox txtindicaciones_zoom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   83
      Top             =   9840
      Visible         =   0   'False
      Width           =   16575
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
      Width           =   16695
      Begin VB.CommandButton cmdNuevoPaciente 
         Height          =   255
         Left            =   2640
         Picture         =   "frmhistoriaclinica.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Limpia el paciente seleccionado"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdEditar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16080
         Picture         =   "frmhistoriaclinica.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Editar Paciente"
         Top             =   360
         Width           =   495
      End
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
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "Descripción"
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox txtNAfil 
         Height          =   285
         Left            =   5160
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txthorad 
         Height          =   285
         Left            =   3960
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Top             =   0
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
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   8
         Top             =   360
         Width           =   1155
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
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "Descripción"
         Top             =   360
         Width           =   2955
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
         Left            =   13800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "Descripción"
         Top             =   360
         Width           =   2235
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
         Left            =   8580
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "Descripción"
         Top             =   360
         Width           =   3555
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
         Left            =   6270
         TabIndex        =   52
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
         Left            =   2460
         TabIndex        =   12
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teléfono/Celular:"
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
         Left            =   12240
         TabIndex        =   10
         Top             =   360
         Width           =   1560
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
         Left            =   7260
         TabIndex        =   9
         Top             =   360
         Width           =   1320
      End
   End
   Begin TabDlg.SSTab tabhc 
      Height          =   8535
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   15055
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Curso Clinico"
      TabPicture(0)   =   "frmhistoriaclinica.frx":1944
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Frame4"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ecografias / Protocolos"
      TabPicture(1)   =   "frmhistoriaclinica.frx":1960
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdEliminarEco"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAgregarEco"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdzoom(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Pedidos"
      TabPicture(2)   =   "frmhistoriaclinica.frx":197C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "Frame8"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdzoom 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Zoom ++"
         Top             =   7440
         Width           =   375
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
         Height          =   7935
         Left            =   -74880
         TabIndex        =   94
         Top             =   480
         Width           =   8055
         Begin VB.TextBox txtProfesionPedido 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4560
            TabIndex        =   104
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtConsultorioPedido 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   103
            Top             =   1080
            Width           =   495
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
            ItemData        =   "frmhistoriaclinica.frx":1998
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":199A
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   1080
            Width           =   2220
         End
         Begin VB.CommandButton cmdCancelarPedido 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   101
            Top             =   7440
            Width           =   1095
         End
         Begin VB.TextBox txtDescPedido 
            Height          =   4995
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   100
            Top             =   2400
            Width           =   6375
         End
         Begin VB.TextBox txtMotivoPedido 
            Height          =   315
            Left            =   1305
            TabIndex        =   99
            Top             =   1440
            Width           =   6375
         End
         Begin VB.CommandButton cmdAceptarPedido 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   98
            Top             =   7440
            Width           =   1095
         End
         Begin VB.CommandButton cmdImprimirPedido 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   4560
            TabIndex        =   97
            Top             =   7440
            Width           =   1095
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
            ItemData        =   "frmhistoriaclinica.frx":199C
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":199E
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1920
            Width           =   2220
         End
         Begin VB.TextBox txtnroPedido 
            Height          =   315
            Left            =   2880
            TabIndex        =   95
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComCtl2.DTPicker FechaPed 
            Height          =   315
            Left            =   1305
            TabIndex        =   105
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin VB.Label Label19 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6240
            TabIndex        =   112
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3720
            TabIndex        =   111
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   240
            TabIndex        =   110
            Top             =   1140
            Width           =   540
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   109
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   240
            TabIndex        =   108
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Especialidad:"
            Height          =   195
            Left            =   240
            TabIndex        =   107
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   106
            Top             =   660
            Width           =   495
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
         Height          =   7935
         Left            =   -66600
         TabIndex        =   85
         Top             =   480
         Width           =   8055
         Begin VB.CommandButton cmdFiltroPedidos 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   87
            Top             =   360
            Width           =   855
         End
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
            TabIndex        =   86
            Top             =   375
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker FechaDesdePedido 
            Height          =   315
            Left            =   2025
            TabIndex        =   88
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHastaPedido 
            Height          =   315
            Left            =   4575
            TabIndex        =   89
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdPedidos 
            Height          =   6375
            Left            =   120
            TabIndex        =   90
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
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   840
            TabIndex        =   93
            Top             =   795
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   2
            Left            =   3600
            TabIndex        =   92
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   1290
            TabIndex        =   91
            Top             =   435
            Width           =   540
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   36
         Top             =   7260
         Width           =   8055
         Begin VB.CommandButton cmdSiguiente 
            Caption         =   "&Siguiente Paciente"
            Height          =   855
            Left            =   1560
            TabIndex        =   48
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdEcogra 
            Caption         =   "&Ecografias"
            Height          =   855
            Left            =   4040
            TabIndex        =   41
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAnterior 
            Caption         =   "&Anterior Paciente"
            Height          =   855
            Left            =   2800
            TabIndex        =   40
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdPedidos 
            Caption         =   "&Pedidos"
            Height          =   855
            Left            =   5280
            TabIndex        =   39
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
         Left            =   8520
         TabIndex        =   67
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
            TabIndex        =   80
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
            TabIndex        =   69
            Top             =   375
            Width           =   2295
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Filtro"
            Height          =   375
            Left            =   6960
            TabIndex        =   68
            Top             =   720
            Width           =   855
         End
         Begin MSComCtl2.DTPicker FechaDesdeImg 
            Height          =   315
            Left            =   1335
            TabIndex        =   70
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHastaImg 
            Height          =   315
            Left            =   4575
            TabIndex        =   71
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdImagenes 
            Height          =   6630
            Left            =   120
            TabIndex        =   72
            Top             =   1200
            Width           =   7500
            _ExtentX        =   13229
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
         Begin VB.CommandButton cmdQuitarProducto 
            Height          =   330
            Left            =   7650
            MaskColor       =   &H8000000F&
            Picture         =   "frmhistoriaclinica.frx":19A0
            Style           =   1  'Graphical
            TabIndex        =   113
            TabStop         =   0   'False
            ToolTipText     =   "Quitar Protocolo"
            Top             =   1920
            UseMaskColor    =   -1  'True
            Width           =   390
         End
         Begin VB.CommandButton cmdabrirdoc 
            Height          =   375
            Left            =   7650
            Picture         =   "frmhistoriaclinica.frx":2722
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Agregar Protocolo"
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label39 
            Caption         =   "Imágen:"
            Height          =   255
            Left            =   3480
            TabIndex        =   79
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   240
            TabIndex        =   75
            Top             =   435
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   74
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde:"
            Height          =   195
            Left            =   240
            TabIndex        =   73
            Top             =   795
            Width           =   990
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
         Left            =   -74880
         TabIndex        =   24
         Top             =   420
         Width           =   8055
         Begin VB.CommandButton cmdzoom 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   7625
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Zoom ++"
            Top             =   5760
            Width           =   375
         End
         Begin VB.TextBox txtProfesion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4680
            TabIndex        =   32
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtConsultorio 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            TabIndex        =   33
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtnrocon 
            Height          =   315
            Left            =   2880
            TabIndex        =   45
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSComCtl2.DTPicker FechaProx 
            Height          =   315
            Left            =   2040
            TabIndex        =   38
            Top             =   6360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   43205
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   29
            Top             =   6360
            Width           =   1095
         End
         Begin VB.TextBox txtIndicaciones 
            Height          =   4035
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   2040
            Width           =   6375
         End
         Begin VB.TextBox txtMotivo 
            Height          =   315
            Left            =   1305
            TabIndex        =   35
            Top             =   1440
            Width           =   6375
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   1305
            TabIndex        =   26
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   30
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
            ItemData        =   "frmhistoriaclinica.frx":89B8
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":89BA
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1080
            Width           =   2580
         End
         Begin VB.Label Label18 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6360
            TabIndex        =   47
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3960
            TabIndex        =   46
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Próxima Consulta:"
            Height          =   375
            Left            =   360
            TabIndex        =   44
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Indicaciones:"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   600
            TabIndex        =   27
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   600
            TabIndex        =   25
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
         Left            =   -66720
         TabIndex        =   17
         Top             =   420
         Width           =   8055
         Begin VB.CommandButton cmdFiltro 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   43
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
            TabIndex        =   18
            Top             =   375
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker FechaDesde 
            Height          =   315
            Left            =   2025
            TabIndex        =   20
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   4575
            TabIndex        =   21
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdConsultas 
            Height          =   6630
            Left            =   120
            TabIndex        =   42
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
            TabIndex        =   23
            Top             =   795
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   0
            Left            =   3600
            TabIndex        =   22
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   1290
            TabIndex        =   19
            Top             =   435
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdAgregarEco 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   16200
         TabIndex        =   3
         Top             =   8460
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEco 
         Caption         =   "Elinimar"
         Height          =   375
         Left            =   18000
         TabIndex        =   2
         Top             =   8460
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Protocolo"
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
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   8295
         Begin VB.TextBox txtImgDescri 
            Height          =   5115
            Index           =   5
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   133
            Top             =   1800
            Width           =   6495
         End
         Begin VB.TextBox txtImgDescri 
            Height          =   5115
            Index           =   4
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   132
            Top             =   1800
            Width           =   6495
         End
         Begin VB.TextBox txtImgDescri 
            Height          =   5115
            Index           =   3
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   131
            Top             =   1800
            Width           =   6495
         End
         Begin VB.TextBox txtImgDescri 
            Height          =   5115
            Index           =   2
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   130
            Top             =   1800
            Width           =   6495
         End
         Begin VB.TextBox txtImgDescri 
            Height          =   5115
            Index           =   1
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   129
            Top             =   1800
            Width           =   6495
         End
         Begin VB.CommandButton cmdpri 
            Caption         =   "<<"
            Height          =   255
            Left            =   3600
            TabIndex        =   128
            Top             =   7200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdult 
            Caption         =   ">>"
            Height          =   255
            Left            =   4680
            TabIndex        =   127
            Top             =   7200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdprev 
            Caption         =   "<"
            Height          =   255
            Left            =   3960
            TabIndex        =   126
            Top             =   7200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdsig 
            Caption         =   ">"
            Height          =   255
            Left            =   4320
            TabIndex        =   124
            Top             =   7200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImprimirEco 
            Caption         =   "Imprimir"
            Height          =   375
            Left            =   5640
            TabIndex        =   81
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
            ItemData        =   "frmhistoriaclinica.frx":89BC
            Left            =   1320
            List            =   "frmhistoriaclinica.frx":89BE
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   840
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
            ItemData        =   "frmhistoriaclinica.frx":89C0
            Left            =   1305
            List            =   "frmhistoriaclinica.frx":89C2
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   1320
            Width           =   2580
         End
         Begin VB.CommandButton cmdAceptarImg 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   60
            Top             =   7440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtImgDescri 
            Height          =   5115
            Index           =   0
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   1800
            Width           =   6495
         End
         Begin VB.CommandButton cmdCancelarImg 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   57
            Top             =   7440
            Width           =   1095
         End
         Begin VB.TextBox txtNroImg 
            Height          =   315
            Left            =   2880
            TabIndex        =   56
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtConsulImg 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7320
            TabIndex        =   55
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtProfImg 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4680
            TabIndex        =   54
            Top             =   1320
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker FechaImg 
            Height          =   315
            Left            =   1305
            TabIndex        =   59
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   58064897
            CurrentDate     =   41098
         End
         Begin VB.Label lblnroja 
            AutoSize        =   -1  'True
            Caption         =   "Label16"
            Height          =   195
            Left            =   4080
            TabIndex        =   125
            Top             =   6960
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Imágen:"
            Height          =   195
            Left            =   600
            TabIndex        =   77
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   600
            TabIndex        =   66
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   240
            TabIndex        =   65
            Top             =   1800
            Width           =   885
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   600
            TabIndex        =   64
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión:"
            Height          =   375
            Left            =   3960
            TabIndex        =   63
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Consultorio:"
            Height          =   255
            Left            =   6480
            TabIndex        =   62
            Top             =   1320
            Width           =   855
         End
      End
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   120
      Top             =   9600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Doctor:"
      Height          =   195
      Left            =   720
      TabIndex        =   76
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Nro Carnet:"
      Height          =   195
      Left            =   5640
      TabIndex        =   50
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
        'TurOSocial = ChkNull(Rec2!TUR_OSOCIAL)
        Calculo_Edad IIf(IsNull(Rec2!CLI_CUMPLE), Date, Rec2!CLI_CUMPLE)
    End If
    Rec2.Close
    BuscarProxPaciente = CodPac
End Function
Private Function Calculo_Edad(cumple As Date)
    'calculo de edad
    If cumple <> 0 Then
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
    If txtImgDescri(0).Text = "" Then
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
    Dim i As Integer
    FechaImg.Value = Date
    For i = 0 To 5
        txtImgDescri(i).Text = ""
    Next
    cboImg.ListIndex = -1
    txtNroImg = ""
    lblnroja.Caption = "Hoja 1"
    muestro_ImgDescri 1
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
'    Set word = CreateObject("word.Basic")
'    cmdAceptar.Enabled = True
'    On Error Resume Next
'    CommonDialog2.CancelError = True
'    CommonDialog2.DialogTitle = "Seleccione un nombre de archivo"
'    CommonDialog2.Filter = "Documents(*.doc;*.docx)"
'
'    CommonDialog2.ShowOpen
'    If Err.Number = 0 Then
'        'If CommonDialog1.FileName Like "*.bmp" _
'        'Or CommonDialog1.FileName Like "*.gif" _
'        'Or CommonDialog1.FileName Like "*.jpg" Then
'
'            'Image1.Picture = LoadPicture(CommonDialog1.FileName)
'            'txtimagen.Text = CommonDialog1.FileName
'            word.FileOpen (CommonDialog2.FileName)
'            word.AppShow
'            'word.filePrintDefault
'            On Error GoTo 0
'        'Else
'        '    MsgBox "El Archivo seleccionado no es válido", vbExclamation, Me.Caption
'        'End If
'
'    End If
'    'word.AppClose
    If txtCodigo.Text <> "" Then
        fraprotocolos.Visible = True
        grdProtocolos.SetFocus
        grdProtocolos.Rows = 1
        cargo_protocolos
    End If
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
        sql = sql & " CLI_CODIGO,VEN_CODIGO,CCL_MOTIVO,CCL_INDICA,CCL_FECPC,CCL_HORA)"
        sql = sql & " VALUES ("
        sql = sql & Num & ","
        sql = sql & XDQ(Fecha.Value) & ","
        sql = sql & XN(txtCodigo.Text) & ","
        sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex) & ","
        sql = sql & XS(txtMotivo.Text, True) & ","
        sql = sql & XS(txtIndicaciones.Text, True) & ","
        sql = sql & XDQ(ChkNull(FechaProx.Value)) & ","
        sql = sql & XS(Format(Time(), "hh:mm")) & ")"
        
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
        sql = sql & " ,CCL_MOTIVO=" & XS(txtMotivo.Text, True)
        sql = sql & " ,CCL_INDICA=" & XS(txtIndicaciones.Text, True)
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
        'If MsgBox("¿Desea cargar los Datos de la Imágen?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
        sql = sql & " CLI_CODIGO,VEN_CODIGO,TIP_CODIGO,IMG_DESCRI,"
        sql = sql & " IMG_DESCRI1,IMG_DESCRI2,IMG_DESCRI3,IMG_DESCRI4,IMG_DESCRI5)"
        sql = sql & " VALUES ("
        sql = sql & Num & ","
        sql = sql & XDQ(FechaImg.Value) & ","
        sql = sql & XN(txtCodigo.Text) & ","
        sql = sql & cboDocImg.ItemData(cboDocImg.ListIndex) & ","
        sql = sql & cboImg.ItemData(cboImg.ListIndex) & ","
        sql = sql & XS(txtImgDescri(0).Text, True) & ","
        sql = sql & XS(txtImgDescri(1).Text, True) & ","
        sql = sql & XS(txtImgDescri(2).Text, True) & ","
        sql = sql & XS(txtImgDescri(3).Text, True) & ","
        sql = sql & XS(txtImgDescri(4).Text, True) & ","
        sql = sql & XS(txtImgDescri(5).Text, True) & ")"
        DBConn.Execute sql
    Else
        'If MsgBox("¿Desea Modificar la Imagen?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        sql = "UPDATE IMAGEN SET "
        sql = sql & " IMG_FECHA = " & XDQ(FechaImg.Value)
        sql = sql & " ,CLI_CODIGO=" & XN(txtCodigo.Text)
        sql = sql & " ,VEN_CODIGO=" & cboDocCon.ItemData(cboDocCon.ListIndex)
        sql = sql & " ,TIP_CODIGO=" & cboImg.ItemData(cboImg.ListIndex)
        sql = sql & " ,IMG_DESCRI=" & XS(txtImgDescri(0).Text, True)
        sql = sql & " ,IMG_DESCRI1=" & XS(txtImgDescri(1).Text, True)
        sql = sql & " ,IMG_DESCRI2=" & XS(txtImgDescri(2).Text, True)
        sql = sql & " ,IMG_DESCRI3=" & XS(txtImgDescri(3).Text, True)
        sql = sql & " ,IMG_DESCRI4=" & XS(txtImgDescri(4).Text, True)
        sql = sql & " ,IMG_DESCRI5=" & XS(txtImgDescri(5).Text, True)
        sql = sql & " WHERE IMG_CODIGO = " & txtNroImg.Text
        DBConn.Execute sql
    End If
       
    'DBConn.CommitTrans
        
    'cboDesde.ListIndex = cboDesde.ListIndex + 1
    'Next
    'cboDesde.Text = sHoraDAux
    'If MsgBox("¿Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'ImprimirTurno
    
    
    
    
End Sub

Private Sub cmdAceptarP_Click()
'Guardar PROTOCOLO SELECCIONADO en tabla IMAGEN
    Dim i, cont As Integer
    Dim Num As Integer
    cont = 0
    For i = 1 To grdProtocolos.Rows - 1
        If grdProtocolos.TextMatrix(i, 8) = "SI" Then
            sql = "SELECT MAX(IMG_CODIGO) AS NUMERO FROM IMAGEN"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Num = Chk0(rec!Numero) + 1
            End If
            rec.Close
            
        
            sql = "INSERT INTO IMAGEN"
            sql = sql & " (IMG_CODIGO,IMG_FECHA,"
            sql = sql & " CLI_CODIGO,VEN_CODIGO,TIP_CODIGO,IMG_DESCRI,"
            sql = sql & " IMG_DESCRI1,IMG_DESCRI2,IMG_DESCRI3,IMG_DESCRI4,IMG_DESCRI5)"
            sql = sql & " VALUES ("
            sql = sql & Num & ","
            sql = sql & XDQ(FechaImg.Value) & ","
            sql = sql & txtCodigo.Text & ","
            sql = sql & cboDocImg.ItemData(cboDocImg.ListIndex) & "," 'SOLO SILVANA ES LA ECOGRAFA
            sql = sql & grdProtocolos.TextMatrix(i, 1) & ","
            sql = sql & XS(grdProtocolos.TextMatrix(i, 2)) & ","
            sql = sql & XS(grdProtocolos.TextMatrix(i, 3)) & ","
            sql = sql & XS(grdProtocolos.TextMatrix(i, 4)) & ","
            sql = sql & XS(grdProtocolos.TextMatrix(i, 5)) & ","
            sql = sql & XS(grdProtocolos.TextMatrix(i, 6)) & ","
            sql = sql & XS(grdProtocolos.TextMatrix(i, 7)) & ")"
            DBConn.Execute sql
            cont = cont + 1
        End If
    Next
    'seleccionar el reciente agregado
    'cargo_protocolo 1
    If cont > 0 Then
        MsgBox "Protocolo agregado a la Historia Clinica (Ecografias) del Paciente" & txtBuscarCliDescri.Text & ". ", vbInformation, TIT_MSGBOX
        CargarImagenesAnteriores
        fraprotocolos.Visible = False
        'frmhistoriaclinica.tabhc.Tab = 1
        'frmhistoriaclinica.txtCodigo = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
        'frmhistoriaclinica.Show vbModal
    End If
    cargo_protocolo 1
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
    If MsgBox("¿Seguro desea Cancelar la Consulta Medica?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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

Private Sub cmdEditar_Click()
    If txtCodigo.Text <> "" Then
        vMode = 2
        gPaciente = txtCodigo.Text
        ABMClientes.Show vbModal
    End If
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

End Sub

Private Sub cmdImprimirEco_Click()
    cmdAceptarImg_Click
    If txtNroImg.Text <> "" Then
        Rep.WindowState = crptMaximized
        Rep.WindowBorderStyle = crptNoBorder
        Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR

        Rep.SelectionFormula = ""
        Rep.Formulas(0) = ""

        Rep.SelectionFormula = " {IMAGEN.IMG_CODIGO}= " & XN(txtNroImg.Text)

        Rep.WindowTitle = "Protocolos"
        Select Case cboDocImg.ItemData(cboDocImg.ListIndex)
            Case 2 'Lelo
                Rep.ReportFileName = DirReport & "rptImagen_lelo.rpt"
   
            Case Else
                Rep.ReportFileName = DirReport & "rptImagen.rpt"
        End Select
        
        Rep.Action = 1
    '    lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        Rep.SelectionFormula = ""
    End If
'    Dim X As Printer
'    Dim mDriver As String
'    mDriver = IMPRESORA
'    For Each X In Printers
'        If X.DeviceName = mDriver Then
'            ' La define como predeterminada del sistema.
'            Set Printer = X
'            Exit For
'        End If
'    Next
''-----------------------------------
'    Set_Impresora
'    ImprimirProtocolo
    'LimpiarImagen
    CargarImagenesAnteriores
End Sub
Private Function ImprimirProtocolo()
    Dim Renglon As Double
    Dim canttxt As Integer
    Dim cantHojas As Integer
    Screen.MousePointer = vbHourglass
'    lblEstado.Caption = "Imprimiendo..."
    cantHojas = 0
    For i = 0 To 5
        If txtImgDescri(i).Text <> "" Then
            cantHojas = cantHojas + 1
        End If
    Next
    ImprimirEncabezado
    For w = 0 To cantHojas '1 'SE IMPRIME POR DUPLICADO
        If w = 1 Then
            Imprimir 10, 0, True, txtImgDescri(w).Text
        Else
            Imprimir 6, 0, True, txtImgDescri(w).Text
        End If
        Printer.EndDoc
    Next w
    Screen.MousePointer = vbNormal
End Function
Private Function ImprimirEncabezado()
    Imprimir 6, 1, True, "Paciente: " & txtBuscarCliDescri
    Imprimir 6.5, 1, True, "Fecha: " & FechaImg.Value
    Imprimir 7, 1, True, "Edad: " & txtEdad
    Imprimir 7.5, 1, True, "Medico Solicitante: "
End Function
Private Sub cmdLabora_Click()
    tabhc.Tab = 2
End Sub

Private Sub CmdNuevo_Click()
    LimpiarConsulta
    limpiarpaciente
    grdConsultas.Rows = 1
    'txtBuscaCliente.Text = ""
    'txtBuscaCliente_LostFocus
    tabhc.Tab = 0
End Sub
Private Function limpiarpaciente()
    txtBuscaCliente = ""
    txtBuscarCliDescri = ""
    txtCodigo = ""
    txthorad = ""
    txtNAfil = ""
    txtEdad = ""
    txtOSocial = ""
    txtTelefono = ""
End Function

Private Sub cmdNuevoPaciente_Click()
    limpiarpaciente
End Sub

Private Sub cmdPedidos_Click()
    tabhc.Tab = 2
End Sub

Private Sub cmdproximo_Click()

End Sub

Private Sub cmdprev_Click()
    If hojaactual > 0 Then
        lblnroja.Caption = "Hoja " & hojaactual
         muestro_ImgDescri hojaactual
    End If
End Sub

Private Sub cmdpri_Click()
    lblnroja.Caption = "Hoja 1"
    muestro_ImgDescri 1
    
End Sub

Private Sub cmdsig_Click()
 If hojaactual < 5 Then
    lblnroja.Caption = "Hoja " & hojaactual + 2
    muestro_ImgDescri hojaactual + 2
 End If
End Sub

Private Sub cmdult_Click()
    lblnroja.Caption = "Hoja 6"
    muestro_ImgDescri 6
End Sub

Private Sub cmdQuitarProducto_Click()
    If grdImagenes.Rows > 1 Then
        If MsgBox("¿Confirma la eliminacion del Protocolo " & grdImagenes.TextMatrix(grdImagenes.RowSel, 2) & " ?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
            If grdImagenes.Rows > 2 Then
                borrar_protocolo XN(txtCodigo.Text), XN(grdImagenes.TextMatrix(grdImagenes.RowSel, 6))
                grdImagenes.RemoveItem (grdImagenes.RowSel)
            Else
                borrar_protocolo XN(txtCodigo.Text), XN(grdImagenes.TextMatrix(grdImagenes.RowSel, 6))
                grdImagenes.Rows = 1
                cmdCancelar_Click
            End If
        End If
    End If
    LimpiarImagen
End Sub
Private Function borrar_protocolo(paciente As Integer, imagen As Integer)
    sql = "DELETE FROM IMAGEN"
    sql = sql & " WHERE CLI_CODIGO = " & paciente
    sql = sql & " AND IMG_CODIGO = " & imagen
    DBConn.Execute sql
End Function

Private Sub cmdSalir_Click()
    
    Unload Me
End Sub

Private Sub Command7_Click()

End Sub

Private Sub cmdSalirP_Click()
    fraprotocolos.Visible = False
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



Private Sub cmdzoom_Click(Index As Integer)
    Select Case Index
    Case 0
        txtindicaciones_zoom(0).Visible = True
        txtindicaciones_zoom(0).Top = 1080
        cmdzoom_out(0).Visible = True
        txtindicaciones_zoom(0).Text = txtIndicaciones.Text
    Case 1
        'txtindicaciones_zoom(1).Visible = True
        'txtindicaciones_zoom(1).Top = 1080
        'cmdzoom_out(1).Visible = True
        'txtindicaciones_zoom(1).Text = txtImgDescri.Text
    Case 2
    End Select
End Sub



Private Sub Command1_Click()

End Sub

Private Sub cmdzoom_out_Click(Index As Integer)
    Select Case Index
    Case 0 'cursoclinico
        txtindicaciones_zoom(0).Visible = False
        cmdzoom_out(0).Visible = False
        txtIndicaciones.Text = txtindicaciones_zoom(0).Text
    Case 1 'imagenes/protocolos
'        txtindicaciones_zoom(1).Visible = False
'        cmdzoom_out(1).Visible = False
'        txtImgDescri.Text = txtindicaciones_zoom(1).Text
    End Select
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
    
    'If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
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
    lblnroja.Caption = "Hoja 1"
    muestro_ImgDescri 1
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
    grdConsultas.FormatString = "Fecha|Doctor|Motivo|Indicaciones|FechaProx|CodMedico|CCL_NUMERO|CCL_CONMUTUAL|CCL_HORA"
    grdConsultas.ColWidth(0) = 1500  'Fecha
    grdConsultas.ColWidth(1) = 2500 'Doctor
    grdConsultas.ColWidth(2) = 3500 'Motivo
    grdConsultas.ColWidth(3) = 0 'Indicaciones
    grdConsultas.ColWidth(4) = 0 'Fecha Proxima
    grdConsultas.ColWidth(5) = 0 'CodMedico
    grdConsultas.ColWidth(6) = 0 'CCL_NUMERO
    grdConsultas.ColWidth(7) = 0 'CCL_CONMUTUAL
    grdConsultas.ColWidth(8) = 0 'CCL_HORA
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
    grdImagenes.FormatString = "Fecha|Doctor|<Imágen|Descripcion|TipoIMG|CodMedico|IMG_CODIGO|Descripcion1|Descripcion2|Descripcion3|Descripcion4|Descripcion5"
    
    grdImagenes.ColWidth(0) = 1500  'Fecha
    grdImagenes.ColWidth(1) = 2500 'Doctor
    grdImagenes.ColWidth(2) = 3200 'Imagen
    grdImagenes.ColWidth(3) = 0 'descripcion
    grdImagenes.ColWidth(4) = 0 'tipo Img
    grdImagenes.ColWidth(5) = 0 'CodMedico
    grdImagenes.ColWidth(6) = 0 'IMG CODIGO
    grdImagenes.ColWidth(7) = 0 'descripcion1
    grdImagenes.ColWidth(8) = 0 'descripcion2
    grdImagenes.ColWidth(9) = 0 'descripcion3
    grdImagenes.ColWidth(10) = 0 'descripcion4
    grdImagenes.ColWidth(11) = 0 'descripcion5
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
  
    grdProtocolos.FormatString = "<Protocolo|Codigo|Contenido|Contenido1|Contenido2|Contenido3|Contenido4|Contenido5|^Seleccionado"
    grdProtocolos.ColWidth(0) = 5300 'Protocolo
    grdProtocolos.ColWidth(1) = 0 'Codigo
    grdProtocolos.ColWidth(2) = 0 'Contenido
    grdProtocolos.ColWidth(3) = 0 'Contenido1
    grdProtocolos.ColWidth(4) = 0 'Contenido2
    grdProtocolos.ColWidth(5) = 0 'Contenido3
    grdProtocolos.ColWidth(6) = 0 'Contenido4
    grdProtocolos.ColWidth(7) = 0 'Contenido5
    grdProtocolos.ColWidth(8) = 1200 'Seleccionar
    grdProtocolos.Rows = 1
    grdProtocolos.HighLight = flexHighlightAlways
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
    Dim DIA As Integer
    Dim horas As Integer
    If grdConsultas.Rows > 1 Then
        Fecha.Value = grdConsultas.TextMatrix(grdConsultas.RowSel, 0)
        BuscaCodigoProxItemData grdConsultas.TextMatrix(grdConsultas.RowSel, 5), cboDocCon
        'cboDocCon.ListIndex = grdConsultas.TextMatrix(grdConsultas.RowSel, 5)
        txtMotivo = grdConsultas.TextMatrix(grdConsultas.RowSel, 2)
        txtIndicaciones = grdConsultas.TextMatrix(grdConsultas.RowSel, 3)
        FechaProx.Value = grdConsultas.TextMatrix(grdConsultas.RowSel, 4)
        txtnrocon.Text = grdConsultas.TextMatrix(grdConsultas.RowSel, 6)
         
        'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
        'solo se habilita si la consultas que estoy viendo fue hecha en menos de 24 hs
        dias = DateDiff("d", grdConsultas.TextMatrix(grdConsultas.RowSel, 0), Date)
        'dias = CDate(Date) - CDate(grdConsultas.TextMatrix(grdConsultas.RowSel, 0))
        'dias = dias * 24
        cmdAceptar.Enabled = False
        If cboDocCon.ItemData(cboDocCon.ListIndex) = Int(Doc) Then
            If dias = 0 Then
                cmdAceptar.Enabled = True
            Else
                If dias = 1 Then
                    If grdConsultas.TextMatrix(grdConsultas.RowSel, 8) <> "" Then
                        If CDate(Time() < CDate(grdConsultas.TextMatrix(grdConsultas.RowSel, 8))) Then
                            cmdAceptar.Enabled = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Sub
Private Function cargo_protocolo(fila As Integer)
    If grdImagenes.Rows > 1 Then
        FechaImg.Value = grdImagenes.TextMatrix(fila, 0)
        BuscaCodigoProxItemData grdImagenes.TextMatrix(fila, 5), cboDocImg
        'codigo del nombre de la imagen
        BuscaCodigoProxItemData grdImagenes.TextMatrix(fila, 4), cboImg
        'cboDocImg.ListIndex = grdImagenes.TextMatrix(fila, 5)
        
        'OJO ACA VER COMO CARGAMOS LA MATRIZ
        txtImgDescri(0).Text = grdImagenes.TextMatrix(fila, 3)
        txtImgDescri(1).Text = grdImagenes.TextMatrix(fila, 7)
        txtImgDescri(2).Text = grdImagenes.TextMatrix(fila, 8)
        txtImgDescri(3).Text = grdImagenes.TextMatrix(fila, 9)
        txtImgDescri(4).Text = grdImagenes.TextMatrix(fila, 10)
        txtImgDescri(5).Text = grdImagenes.TextMatrix(fila, 11)
        
        lblnroja.Caption = "Hoja 1"
        muestro_ImgDescri 1
                
        txtNroImg.Text = grdImagenes.TextMatrix(fila, 6)
         'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
        If cboDocImg.ItemData(cboDocImg.ListIndex) = Int(Doc) Then
            cmdAceptarImg.Enabled = True
        Else
            cmdAceptarImg.Enabled = False
        End If
    End If
End Function
Private Sub grdImagenes_Click()
    LimpiarImagen
    cargo_protocolo grdImagenes.RowSel
'    If grdImagenes.Rows > 1 Then
'        FechaImg.Value = grdImagenes.TextMatrix(grdImagenes.RowSel, 0)
'        BuscaCodigoProxItemData grdImagenes.TextMatrix(grdImagenes.RowSel, 5), cboDocImg
'        'codigo del nombre de la imagen
'        BuscaCodigoProxItemData grdImagenes.TextMatrix(grdImagenes.RowSel, 4), cboImg
'        'cboDocImg.ListIndex = grdImagenes.TextMatrix(grdImagenes.RowSel, 5)
'
'        'OJO ACA VER COMO CARGAMOS LA MATRIZ
'        txtImgDescri(0).Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 3)
'        txtImgDescri(1).Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 7)
'        txtImgDescri(2).Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 8)
'        txtImgDescri(3).Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 9)
'        txtImgDescri(4).Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 10)
'        txtImgDescri(5).Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 11)
'
'        lblnroja.Caption = "Hoja 1"
'        muestro_ImgDescri 1
'
'        txtNroImg.Text = grdImagenes.TextMatrix(grdImagenes.RowSel, 6)
'         'ESTO LO HAGO PARA HABILITAR EL ACEPTAR DE LA CONSULTA MEDICA
'        If cboDocImg.ItemData(cboDocImg.ListIndex) = Int(Doc) Then
'            cmdAceptarImg.Enabled = True
'        Else
'            cmdAceptarImg.Enabled = False
'        End If
'    End If
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

Private Sub grdProtocolos_Click()
    Dim J As Integer
    If grdProtocolos.TextMatrix(grdProtocolos.RowSel, 8) = "NO" Then
        grdProtocolos.TextMatrix(grdProtocolos.RowSel, 8) = "SI"
        'CAMBIAR COLOR
        'backColor = &HC000&
        'foreColor = &HFFFFFF
        For J = 0 To grdProtocolos.Cols - 1
            grdProtocolos.Col = J
            grdProtocolos.CellForeColor = &HFFFFFF
            grdProtocolos.CellBackColor = &HC000&
            grdProtocolos.CellFontBold = True
        Next
    Else
        grdProtocolos.TextMatrix(grdProtocolos.RowSel, 8) = "NO"
        For J = 0 To grdProtocolos.Cols - 1
            grdProtocolos.Col = J
            grdProtocolos.CellForeColor = &H80000008
            grdProtocolos.CellBackColor = &H80000005
            grdProtocolos.CellFontBold = False
        Next
    End If
End Sub

Private Sub grdProtocolos_DblClick()
    cmdAceptarP_Click
End Sub

Private Sub grdProtocolos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        grdProtocolos_DblClick
    End If
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
    If KeyCode = vbKeyReturn Then MySendKeys Chr(9)
End Sub

Private Sub txtBuscaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
End Sub
Private Sub txtBuscaCliente_LostFocus()
    If txtBuscaCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,CLI_CELULAR,CLI_NROAFIL,CLI_CUMPLE,CLI_EDAD"
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
            'If txtTelefono.Text <> "" Then
                txtTelefono.Text = txtTelefono.Text & "/" & ChkNull(rec!CLI_CELULAR)
            'Else
            '    txtTelefono.Text = ChkNull(rec!CLI_CELULAR)
            'End If
            'ATENCION CON O SIN OBRA SOCIAL
            If TurOSocial = "PARTICULAR" Then
                txtOSocial.Text = "PARTICULAR"
            Else
                txtOSocial.Text = BuscarOSocial(txtCodigo.Text) & " - " & ChkNull(rec!CLI_NROAFIL)
            End If
            'Calculo_Edad Chk0(rec!CLI_CUMPLE)
            txtEdad.Text = ChkNull(rec!CLI_EDAD)
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
If KeyCode = vbKeyReturn Then MySendKeys Chr(9)
End Sub

Private Sub txtBuscarCliDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    'If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
End Sub

Private Sub txtBuscarCliDescri_LostFocus()
    If txtBuscaCliente.Text = "" Or txtBuscarCliDescri.Text <> "" Then
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
            If rec.RecordCount > 2 Then
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
                Calculo_Edad Chk0(rec!CLI_CUMPLE)
                CargarConsultasAnteriores
                CargarPedidosAnteriores
                CargarImagenesAnteriores
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
                'txtBuscaCliente_LostFocus
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
    'limpiarpaciente
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
                                 Rec1!CCL_INDICA & Chr(9) & Rec1!CCL_FECPC & Chr(9) & Rec1!VEN_CODIGO & Chr(9) & _
                                 Rec1!CCL_NUMERO & Chr(9) & "" & Chr(9) & ChkNull(Rec1!CCL_HORA)
                                       
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
                                    Rec1!TIP_CODIGO & Chr(9) & Rec1!VEN_CODIGO & Chr(9) & Rec1!IMG_CODIGO & Chr(9) & _
                                    Rec1!IMG_DESCRI1 & Chr(9) & Rec1!IMG_DESCRI2 & Chr(9) & Rec1!IMG_DESCRI3 & Chr(9) & Rec1!IMG_DESCRI4 & Chr(9) & Rec1!IMG_DESCRI5
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
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,CLI_NROAFIL,CLI_CUMPLE,CLI_EDAD"
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
            If TurOSocial = "PARTICULAR" Then
                txtOSocial.Text = "PARTICULAR"
            Else
                txtOSocial.Text = BuscarOSocial(txtCodigo.Text) & " - " & ChkNull(rec!CLI_NROAFIL)
            End If
            'calculo de edad
            'BuscarProxPaciente
            'Calculo_Edad IIf(IsNull(rec!CLI_CUMPLE), Date, rec!CLI_CUMPLE)
            txtOSocial.Text = ChkNull(rec!CLI_EDAD)
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

'Private Sub txtImgDescri_Change()
'    txtindicaciones_zoom(1).Text = txtImgDescri.Text
'End Sub

Private Sub txtIndicaciones_Change()
    txtindicaciones_zoom(0).Text = txtIndicaciones.Text
End Sub


Private Function cargo_protocolos()
    
    sql = "SELECT * FROM TIPO_IMAGEN WHERE VEN_CODIGO=" & cboDocImg.ItemData(cboDocImg.ListIndex)
    If txtfiltrop.Text <> "" Then
        sql = sql & " AND TIP_NOMBRE LIKE '%" & txtfiltrop.Text & "%'"
    End If
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            grdProtocolos.AddItem ChkNull(rec!TIP_NOMBRE) & Chr(9) & _
                                  rec!TIP_CODIGO & Chr(9) & _
                                  rec!TIP_CONTEN & Chr(9) & _
                                  rec!TIP_CONTEN1 & Chr(9) & _
                                  rec!TIP_CONTEN2 & Chr(9) & _
                                  rec!TIP_CONTEN3 & Chr(9) & _
                                  rec!TIP_CONTEN4 & Chr(9) & _
                                  rec!TIP_CONTEN5 & Chr(9) & _
                                  "NO"
            rec.MoveNext
        Loop
    
    End If
    rec.Close
    
End Function

Private Sub txtindicaciones_zoom_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        txtIndicaciones.Text = txtindicaciones_zoom(0).Text
    Case 1
        'txtImgDescri.Text = txtindicaciones_zoom(1).Text
    End Select
End Sub
Private Function muestro_ImgDescri(hoja As Integer)
    Dim i As Integer
        For i = 0 To 5
            If (hoja - 1) = i Then
                txtImgDescri(i).Visible = True
            Else
                txtImgDescri(i).Visible = False
            End If
        Next

End Function
Private Function hojaactual() As Integer
    Dim i As Integer
    For i = 0 To 5
        If txtImgDescri(i).Visible = True Then
            hojaactual = i
        End If
    Next

End Function
