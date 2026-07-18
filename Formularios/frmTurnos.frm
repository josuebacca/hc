VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIGOR - Turnos de Pacientes"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19725
   ForeColor       =   &H00000000&
   Icon            =   "frmTurnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   19725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfiguracionHorarios 
      Caption         =   "Configurar cronograma de atención"
      Height          =   8415
      Left            =   6360
      TabIndex        =   81
      Top             =   960
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdCerrarDisponibilidad 
         Height          =   375
         Left            =   9480
         Picture         =   "frmTurnos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   240
         Width           =   495
      End
      Begin VB.Frame Frame6 
         Caption         =   "Excepción para la fecha actual"
         Height          =   3615
         Left            =   120
         TabIndex        =   93
         Top             =   4680
         Width           =   10095
         Begin VB.Frame Frame4 
            Height          =   1215
            Left            =   5400
            TabIndex        =   106
            Top             =   600
            Width           =   4575
            Begin VB.TextBox txtMotivoExcepcion 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   840
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   112
               Tag             =   "DescripciÃ³n"
               Top             =   600
               Width           =   3615
            End
            Begin VB.OptionButton optAtiende 
               Caption         =   "SI"
               Height          =   255
               Left            =   1080
               TabIndex        =   108
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optNoAtiende 
               Caption         =   "NO"
               Height          =   255
               Left            =   1680
               TabIndex        =   107
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label21 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   240
               TabIndex        =   111
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label18 
               Caption         =   "Atende:"
               Height          =   255
               Left            =   240
               TabIndex        =   109
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdEliminarExcepcion 
            Caption         =   "&Eliminar Excepción"
            Height          =   375
            Left            =   5280
            TabIndex        =   102
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CommandButton cmdGuardarExcepcion 
            Caption         =   "&Guardar Excepción"
            Height          =   375
            Left            =   3240
            TabIndex        =   101
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CommandButton cmdQuitarExc 
            Caption         =   "&Quitar"
            Height          =   345
            Left            =   4320
            TabIndex        =   98
            Top             =   840
            Width           =   945
         End
         Begin VB.CommandButton cmdAgregarExc 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   3240
            TabIndex        =   97
            Top             =   840
            Width           =   945
         End
         Begin MSFlexGridLib.MSFlexGrid grdHorariosExc 
            Height          =   1215
            Left            =   240
            TabIndex        =   94
            Top             =   1800
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   2143
            _Version        =   393216
            SelectionMode   =   1
         End
         Begin MSMask.MaskEdBox mskHoraDesdeExc 
            Height          =   315
            Left            =   960
            TabIndex        =   95
            Top             =   840
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraHastaExc 
            Height          =   315
            Left            =   2400
            TabIndex        =   96
            Top             =   840
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblExcepcion 
            Caption         =   "Excepción de atención para el día"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   360
            TabIndex        =   110
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label20 
            Caption         =   "Hasta:"
            Height          =   375
            Left            =   1800
            TabIndex        =   103
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Desde:"
            Height          =   375
            Left            =   360
            TabIndex        =   100
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblDiaActual 
            Caption         =   "Día:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   3960
            TabIndex        =   99
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Horario habitual"
         Height          =   4095
         Left            =   120
         TabIndex        =   82
         Top             =   600
         Width           =   10095
         Begin VB.CommandButton cmdGuardarDisponibilidad 
            Caption         =   "&Guardar"
            Height          =   375
            Left            =   4200
            TabIndex        =   91
            Top             =   3600
            Width           =   1335
         End
         Begin VB.ComboBox cboDiaSemana 
            Height          =   315
            ItemData        =   "frmTurnos.frx":0FD4
            Left            =   720
            List            =   "frmTurnos.frx":0FD6
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdAgregarHorario 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   6360
            TabIndex        =   88
            Top             =   360
            Width           =   945
         End
         Begin VB.CommandButton cmdQuitarHorario 
            Caption         =   "&Quitar"
            Height          =   345
            Left            =   7440
            TabIndex        =   90
            Top             =   360
            Width           =   945
         End
         Begin MSFlexGridLib.MSFlexGrid grdHorarios 
            Height          =   2655
            Left            =   120
            TabIndex        =   83
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   4683
            _Version        =   393216
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
         Begin MSMask.MaskEdBox mskHoraDesde 
            Height          =   315
            Left            =   3480
            TabIndex        =   85
            Top             =   360
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskHoraHasta 
            Height          =   315
            Left            =   5040
            TabIndex        =   86
            Top             =   360
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid grdDisponibilidad 
            Height          =   2670
            Left            =   2880
            TabIndex        =   105
            Top             =   840
            Width           =   7140
            _ExtentX        =   12594
            _ExtentY        =   4710
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   280
            BackColorSel    =   16761024
            AllowBigSelection=   -1  'True
            FocusRect       =   0
            HighLight       =   0
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
         Begin VB.Label Label17 
            Caption         =   "Día:"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "Desde:"
            Height          =   375
            Left            =   2880
            TabIndex        =   89
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Hasta:"
            Height          =   375
            Left            =   4440
            TabIndex        =   87
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.Frame frmVista 
      Caption         =   "Modo de vista"
      Height          =   735
      Left            =   10800
      TabIndex        =   78
      Top             =   50
      Width           =   3255
      Begin VB.CommandButton cmdLibres 
         Caption         =   "&Ver Libres"
         Height          =   375
         Left            =   1800
         TabIndex        =   80
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSoloTurnos 
         Caption         =   "&Solo Turnos"
         Height          =   375
         Left            =   360
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdEditar 
      Height          =   495
      Left            =   16560
      Picture         =   "frmTurnos.frx":0FD8
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Editar Turno"
      Top             =   50
      Width           =   495
   End
   Begin VB.ComboBox cboTamanioSlot 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   76
      Top             =   480
      Width           =   1980
   End
   Begin VB.Frame fraDisponibilidad 
      Caption         =   "Horarios de atención"
      Height          =   6975
      Left            =   12240
      TabIndex        =   72
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdCerrarDisp 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   3120
         TabIndex        =   74
         Top             =   6360
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdDisponibilidad_old 
         Height          =   5910
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   10425
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
   End
   Begin VB.CommandButton cmdHorariosAtencion 
      Caption         =   "&Horarios de atención"
      Height          =   495
      Left            =   14400
      TabIndex        =   71
      Top             =   50
      Width           =   1695
   End
   Begin VB.CommandButton cmdExportarTurno 
      Caption         =   "&Exportar turno"
      Height          =   735
      Left            =   6360
      Picture         =   "frmTurnos.frx":201A
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Listado de Turnos del dia por Doctor"
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Exportar turnero"
      Height          =   735
      Left            =   5040
      Picture         =   "frmTurnos.frx":2CE4
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Listado de Turnos del dia por Doctor"
      Top             =   9240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   8520
      Picture         =   "frmTurnos.frx":39AE
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9240
      Width           =   975
   End
   Begin VB.Frame fraListaEstudios 
      Caption         =   "Lista de estudios"
      Height          =   3135
      Left            =   8040
      TabIndex        =   64
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdCerrarFraListaEstudios 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   960
         TabIndex        =   66
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ListBox listEstudios 
         Height          =   1815
         Left            =   240
         TabIndex        =   65
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdDrive 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTurnos.frx":49F0
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Ir a protocolos"
      Top             =   50
      Width           =   615
   End
   Begin VB.CommandButton cmdProtocolos 
      Enabled         =   0   'False
      Height          =   495
      Left            =   18000
      Picture         =   "frmTurnos.frx":5ABA
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Protocolos"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdCopiar 
      Height          =   495
      Left            =   17040
      Picture         =   "frmTurnos.frx":77B4
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Copiar Turno"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdCortar 
      Enabled         =   0   'False
      Height          =   495
      Left            =   17520
      Picture         =   "frmTurnos.frx":7B3E
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Cortar Turnos"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdImpTurno 
      Enabled         =   0   'False
      Height          =   495
      Left            =   16080
      Picture         =   "frmTurnos.frx":7EC8
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "ImprimirTurno"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdOcultar 
      Height          =   495
      Left            =   18480
      TabIndex        =   54
      Top             =   50
      Width           =   495
   End
   Begin MSComCtl2.DTPicker fechaturno 
      Height          =   375
      Left            =   13560
      TabIndex        =   53
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   151715841
      CurrentDate     =   43340
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
      Height          =   8655
      Left            =   10200
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtfiltrop 
         Height          =   315
         Left            =   1800
         TabIndex        =   49
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdSalirP 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   5760
         TabIndex        =   48
         Top             =   8040
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdProtocolos 
         Height          =   7230
         Left            =   120
         TabIndex        =   46
         Top             =   720
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
      Begin VB.CommandButton cmdAceptarP 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   4320
         TabIndex        =   47
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Filtro"
         Height          =   195
         Left            =   1320
         TabIndex        =   50
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdatendido 
      BackColor       =   &H00008000&
      Height          =   315
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Atendido"
      Top             =   450
      Width           =   495
   End
   Begin VB.CommandButton cmdespera 
      BackColor       =   &H0000C0C0&
      Height          =   315
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "En Espera"
      Top             =   450
      Width           =   495
   End
   Begin VB.CommandButton cmdpendiente 
      BackColor       =   &H00800080&
      Height          =   315
      Left            =   5400
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Pendiente"
      Top             =   450
      Width           =   495
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   400
      Left            =   15600
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   35
      Tag             =   "DescripciÃ³n"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdInforTurno 
      Caption         =   "&Historial"
      Height          =   735
      Left            =   7560
      Picture         =   "frmTurnos.frx":DADA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Reporte"
      Height          =   735
      Left            =   4080
      Picture         =   "frmTurnos.frx":13D64
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Listado de Turnos del dia por Doctor"
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   3120
      Picture         =   "frmTurnos.frx":14A2E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar Turnos"
      Height          =   735
      Left            =   2160
      Picture         =   "frmTurnos.frx":15A70
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   3495
      Begin VB.ComboBox cboDoctor 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   200
         Width           =   2700
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   120
      TabIndex        =   18
      Top             =   3285
      Width           =   3495
      Begin VB.TextBox txtObservaciones 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   75
         MultiLine       =   -1  'True
         TabIndex        =   7
         Tag             =   "DescripciÃ³n"
         Top             =   4200
         Width           =   3270
      End
      Begin VB.TextBox txtcelular 
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
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   61
         Tag             =   "DescripciÃ³n"
         Top             =   1320
         Width           =   1755
      End
      Begin VB.ComboBox cboMotivo 
         BackColor       =   &H0080FF80&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdNuevoPaciente 
         Caption         =   "+"
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
         Left            =   3075
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Agregar nuevo Paciente"
         Top             =   250
         Width           =   255
      End
      Begin MSMask.MaskEdBox mebHoraD 
         Height          =   315
         Left            =   795
         TabIndex        =   8
         Top             =   5040
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDrSolicitante 
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
         Left            =   120
         MaxLength       =   75
         TabIndex        =   6
         Tag             =   "DescripciÃ³n"
         Top             =   3660
         Width           =   3270
      End
      Begin VB.OptionButton optNO 
         Caption         =   "NO"
         Height          =   315
         Left            =   2040
         TabIndex        =   40
         Top             =   1700
         Width           =   615
      End
      Begin VB.OptionButton optSI 
         Caption         =   "SI"
         Height          =   315
         Left            =   1440
         TabIndex        =   39
         Top             =   1700
         Width           =   615
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "DescripciÃ³n"
         Top             =   2040
         Width           =   3270
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   855
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
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         Tag             =   "DescripciÃ³n"
         Top             =   960
         Width           =   1755
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "DescripciÃ³n"
         Top             =   615
         Width           =   3195
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
         Height          =   315
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   2
         Top             =   250
         Width           =   1515
      End
      Begin VB.TextBox txtMotivo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   5
         Tag             =   "DescripciÃ³n"
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtimporte 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "DescripciÃ³n"
         Text            =   "0,00"
         Top             =   5415
         Width           =   1635
      End
      Begin MSMask.MaskEdBox mebHoraH 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   5040
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbohasta 
         BackColor       =   &H8000000E&
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4980
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.ComboBox cboDesde 
         BackColor       =   &H8000000E&
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   5280
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox txtOrden 
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
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "DescripciÃ³n"
         ToolTipText     =   "Orden del Turno"
         Top             =   1700
         Width           =   435
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
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
         TabIndex        =   67
         Top             =   3960
         Width           =   3270
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Celular:"
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
         TabIndex        =   62
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label12 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         TabIndex        =   52
         Top             =   5070
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   5070
         Width           =   615
      End
      Begin VB.Label lblimporte 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe:"
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
         TabIndex        =   36
         Top             =   5415
         Width           =   1605
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dr Solicitante"
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
         TabIndex        =   41
         Top             =   3300
         Width           =   3270
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   33
         Top             =   1700
         Width           =   1200
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horario:"
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
         TabIndex        =   23
         Top             =   4620
         Width           =   3270
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Motivo:"
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
         TabIndex        =   22
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paciente:"
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
         TabIndex        =   21
         Top             =   250
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   3495
      Begin MSComCtl2.MonthView MViewFecha 
         Height          =   2370
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   151715842
         TitleBackColor  =   4194304
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483638
         CurrentDate     =   40049
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdGrilla 
      Height          =   7845
      Left            =   3600
      TabIndex        =   14
      ToolTipText     =   "Doble Click para ver la Historia Clinica del Paciente"
      Top             =   1320
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   13838
      _Version        =   393216
      Rows            =   25
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   290
      BackColor       =   12648384
      ForeColor       =   49152
      ForeColorFixed  =   -2147483635
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      GridColor       =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
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
   Begin Crystal.CrystalReport Rep 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   735
      Left            =   1200
      Picture         =   "frmTurnos.frx":15DFA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9240
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   735
      Left            =   240
      Picture         =   "frmTurnos.frx":16E3C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label lblExcepcionMotivo 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Excepcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3720
      TabIndex        =   113
      Top             =   960
      Width           =   7530
   End
   Begin VB.Label lblTamanioSlot 
      AutoSize        =   -1  'True
      Caption         =   "Duración de turnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   75
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14100
      TabIndex        =   38
      Top             =   9480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe:"
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
      Left            =   9720
      TabIndex        =   37
      Top             =   11160
      Width           =   1245
   End
   Begin VB.Label lblAux 
      Caption         =   "Label7"
      Height          =   255
      Left            =   12120
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado del Turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   29
      Top             =   480
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "<F5 para actualizar el Turnero>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   16800
      TabIndex        =   28
      Top             =   690
      Width           =   2685
   End
   Begin VB.Label lbldiaTurno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   3840
      TabIndex        =   19
      Top             =   60
      Width           =   945
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Left            =   3720
      Top             =   60
      Width           =   7005
   End
End
Attribute VB_Name = "frmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Dim j As Integer
Dim hDesde As Integer
Dim hHasta As Integer
Dim ActivoGrid As Integer ' 1 actio 0 desactivo
Dim sAction As String
Dim dFechaCopy As String
Dim nDoctorCopy As String
Dim sNameDoctorCopy As String
Dim linkProtocolos As String
Dim studiesDict As Variant
Dim estudiosUrls As Object ' Dictionary para mapear índice -> URL
Dim dictLlaves As Object
Dim dicHorariosDoc As Object   ' Scripting.Dictionary
Dim modoActualizacionTurno As Integer
' ID_TURNO del turno seleccionado (solo relevante en modo edición)
Dim idTurnoSeleccionado As Long

' ------------------------------------------------------------------
'  1. CONSTANTES DE COLORES (agregar a nivel formulario)
' ------------------------------------------------------------------
Const COLOR_DISPONIBLE_BACK As Long = &HFFF8F0     ' celeste muy suave, casi blanco
Const COLOR_DISPONIBLE_FORE As Long = &H80000008    ' texto del sistema (negro)
Const COLOR_FUERA_BACK As Long = &HE0E0E0           ' gris claro
Const COLOR_FUERA_FORE As Long = &HA0A0A0           ' gris medio
' ------------------------------------------------------------------
'  VARIABLES DE EXCEPCIONES HORARIOS
' ------------------------------------------------------------------
Dim vExcDiaId As Long       ' ID de la excepción actual, 0 si no hay
Dim modoVista As Integer


' ------------------------------------------------------------------
'  2. FUNCIONES AUXILIARES
' ------------------------------------------------------------------

' Convierte "08:30" ? 510 (minutos desde medianoche)
Private Function HoraAMinutos(ByVal Hora As String) As Integer
    HoraAMinutos = CInt(Left(Hora, 2)) * 60 + CInt(Right(Hora, 2))
End Function

' Convierte 510 ? "08:30"
Private Function MinutosAHora(ByVal minutos As Integer) As String
    MinutosAHora = Format(minutos \ 60, "00") & ":" & Format(minutos Mod 60, "00")
End Function

' Verifica si una hora cae dentro de algún bloque de atención
' bloquesStr tiene formato "08:00-12:00;14:00-18:00"
Private Function EsHorarioAtencion(ByVal horaSlot As String, ByVal bloquesStr As String) As Boolean
    Dim bloques() As String
    Dim rango() As String
    Dim B As Integer
    
    EsHorarioAtencion = False
    bloques = Split(bloquesStr, ";")
    
    For B = 0 To UBound(bloques)
        rango = Split(bloques(B), "-")
        ' El slot está en atención si su inicio cae dentro del bloque
        If horaSlot >= rango(0) And horaSlot < rango(1) Then
            EsHorarioAtencion = True
            Exit Function
        End If
    Next B
End Function
' ------------------------------------------------------------------
'  3. CONFIGURACIÓN DEL COMBO DE TAMAÑO DE SLOT
'     Llamar desde Form_Load
' ------------------------------------------------------------------
Private Sub ConfigurarCboTamanioSlot()
    cboTamanioSlot.Clear
        cboTamanioSlot.AddItem "5 min"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 5
        cboTamanioSlot.AddItem "10 min"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 10
    cboTamanioSlot.AddItem "15 min"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 15
    cboTamanioSlot.AddItem "20 min"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 20
    cboTamanioSlot.AddItem "30 min"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 30
    cboTamanioSlot.AddItem "45 min"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 45
    cboTamanioSlot.AddItem "1 hora"
    cboTamanioSlot.ItemData(cboTamanioSlot.NewIndex) = 60
    cboTamanioSlot.ListIndex = 2   ' default: 30 min
End Sub

' Al cambiar el tamaño de slot, refrescar la grilla.
' Reemplazar con las variables reales de fecha y doctor actual.
Private Sub cboTamanioSlot_Click()
If cboDoctor.text <> "" Then
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
End If

End Sub


' ------------------------------------------------------------------
'  4. ARMAR FILA VACÍA (21 columnas)
' ------------------------------------------------------------------
Private Function ArmarFilaSlotVacia(ByVal horaDesde As String, _
                                     ByVal horaHasta As String, _
                                     ByVal texto As String) As String
    Dim Fila As String
    Fila = horaDesde & " a " & horaHasta _
         & Chr(9) & texto _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & "" _
         & Chr(9) & ""
    ArmarFilaSlotVacia = Fila
End Function


Private Sub cboDesde_LostFocus()
    If cboDesde.ListIndex < cboDesde.ListCount - 1 Then
        cbohasta.ListIndex = cboDesde.ListIndex + 1
    Else
        cbohasta.ListIndex = cboDesde.ListIndex
    End If
End Sub

Private Sub cboDoctor_Click()
    LimpiarComboMotivo
    'Buscar link drive
    BuscarLinkDrive
    LimpiarConfiguracionHorarios
    If cboDoctor.ListIndex <> -1 Then
        sql = "SELECT M.MOT_DESCRI"
            sql = sql & " FROM  MOTIVO_VENDEDOR MV,VENDEDOR V,MOTIVO M "
            sql = sql & " WHERE V.VEN_NOMBRE = " & XS(cboDoctor.text)
            sql = sql & " AND V.VEN_CODIGO = MV.VEN_CODIGO"
            sql = sql & " AND MV.MOT_CODIGO = M.MOT_CODIGO"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        Do While Rec1.EOF = False
                cboMotivo.AddItem Rec1!MOT_DESCRI
                Rec1.MoveNext
        Loop
        Rec1.Close
    End If
    LimpiarGrilla
    CargarHorarios cboDoctor.ItemData(cboDoctor.ListIndex)
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex), modoVista

End Sub
Private Sub LimpiarComboMotivo()
    cboMotivo.Clear
End Sub
Private Sub LimpiarComboDoctor()
    cboDoctor.Clear
End Sub
Private Sub BuscarLinkDrive()
    sql = "SELECT VEN_CODIGO, VEN_LINKPROT FROM VENDEDOR"
    sql = sql & " WHERE VEN_CODIGO= " & cboDoctor.ItemData(cboDoctor.ListIndex)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

    If rec.EOF = False Then
       linkProtocolos = ChkNull(rec!VEN_LINKPROT)
    End If
    rec.Close
End Sub

Private Sub cboDoctor_Change()
    LimpiarGrilla
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)

End Sub

Private Sub cbohasta_LostFocus()
    If cboDesde.ListIndex = -1 Then
        If cbohasta.ListIndex > 0 Then
            cboDesde.ListIndex = cbohasta.ListIndex - 1
        Else
            cboDesde.ListIndex = cbohasta.ListIndex
        End If
    End If
End Sub
Private Function ValidarHorarioTurno() As Boolean
If mebHoraH.text <= mebHoraD.text Then
    MsgBox "La hora HASTA debe ser mayor que la hora DESDE", vbCritical, TIT_MSGBOX
Else

    If ValidarRangoTurno = False Then
        MsgBox "El horario ingresado para el turno no esta disponible, por favor ingrese otro.", vbCritical, TIT_MSGBOX
        ValidarHorarioTurno = False
    Else
        ValidarHorarioTurno = True
    End If
End If
End Function
Private Function ValidarRangoTurno() As Boolean
Dim i As Integer
Dim turdesde As Date
Dim turhasta As Date
Dim hasta As Date
Dim desde As Date
hasta = mebHoraH.text
desde = cboDesde.text
If grdGrilla.rows < 2 Then
    ValidarRangoTurno = True
Else
   For i = 1 To grdGrilla.rows - 1
   turdesde = Format(Left(grdGrilla.TextMatrix(i, 0), 5), "hh:mm")
   turhasta = Format(Right(grdGrilla.TextMatrix(i, 0), 5), "hh:mm")
   'si la hora hasta es menor o igual a la desde, lo agrego
   If hasta <= turdesde Then
        ValidarRangoTurno = True
        Exit For
    End If
    'comparo si esta en un rango ya ocupado
   If (desde > turdesde And desde < turhasta Or hasta > turdesde And hasta <= turhasta) Or (desde < turdesde And hasta >= turhasta) Then
        ValidarRangoTurno = False
        Exit For
    Else
    'se puede cargar
    ValidarRangoTurno = True
   End If
   Next
End If
End Function
Private Function ValidarTurno() As Boolean
    If MViewFecha.Value < Date Then
'        MsgBox "No puede agregar un turno para ese dia", vbCritical, TIT_MSGBOX
'        MViewFecha.SetFocus
'        ValidarTurno = False
'        Exit Function
    End If
    If txtBuscaCliente.text = "" Then
        MsgBox "No ha ingresado el paciente", vbCritical, TIT_MSGBOX
        txtBuscaCliente.SetFocus
        ValidarTurno = False
        Exit Function
    End If
'    If txtMotivo.Text = "" Then
'        MsgBox "No ha ingresado el Motivo del Turno", vbCritical, TIT_MSGBOX
'        txtMotivo.SetFocus
'        ValidarTurno = False
'        Exit Function
'    End If
    If mebHoraD.text = "" Then
        MsgBox "No ha ingresado la hora de comienzo del Turno", vbCritical, TIT_MSGBOX
        mebHoraD.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If mebHoraH.text = "" Then
        MsgBox "No ha ingresado la hora de finalización del Turno", vbCritical, TIT_MSGBOX
        mebHoraH.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If mebHoraD.text >= mebHoraH.text Then
        MsgBox "La hora HASTA debe ser mayor a la hora DESDE", vbCritical, TIT_MSGBOX
        mebHoraD.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    
    ValidarTurno = True
End Function
Private Function actualizo_turno_impreso()
    sql = " UPDATE TURNOS SET TUR_IMPRESO = 1 "
    sql = sql & " WHERE "
    sql = sql & " TUR_FECHA = " & XDQ(fechaturno.Value)
    sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & mebHoraD.text & "'"
    sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
    DBConn.Execute sql
    
End Function

Private Function ImprimirTurno()
    Dim sHoraD As Date
    Dim mNombreImpresora As String
    Dim strAutoSaveDirectory As String
    Dim strAutoSaveFileName As String
    Dim Cliente As String
    Dim Fecha As String
    Dim objWSH As Object
    
    
    Cliente = Replace(txtBuscarCliDescri.text, " ", "_")
    Fecha = Replace(MViewFecha.Value, "/", "")
    
    sHoraD = mebHoraD.text
    sHoraD = Mid(mebHoraD, 1, 1)
    
    If sHoraD = "0" Then
        sHoraD = Mid(mebHoraD.text, 2, 4)
    Else
        sHoraD = Mid(mebHoraD.text, 1, 5)
    End If
    
    mNombreImpresora = Printer.DeviceName
    Rep.Destination = 1
    Call EstableceDefaultPrinter("PDFCreator")

    
    strAutoSaveDirectory = DirReport & "\Turnos\"
    strAutoSaveFileName = "TURNO_" & Cliente & "_" & Fecha & ".pdf"
'
    If Dir(strAutoSaveDirectory & strAutoSaveFileName) <> "" Then Kill (strAutoSaveDirectory & strAutoSaveFileName)
'
    Set objWSH = CreateObject("WScript.Shell")
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\UseAutoSave", 1, "REG_SZ"
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\UseAutoSaveDirectory", 1, "REG_SZ"
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutoSaveDirectory", strAutoSaveDirectory, "REG_SZ"
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutoSaveFileName", strAutoSaveFileName, "REG_SZ"
'    'Rep.Action = 1
    
    
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""

    Rep.SelectionFormula = " {TURNOS.TUR_FECHA}= DATE (" & Mid(MViewFecha.Value, 7, 4) & "," & Mid(MViewFecha.Value, 4, 2) & "," & Mid(MViewFecha.Value, 1, 2) & ")"
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.VEN_CODIGO}= " & cboDoctor.ItemData(cboDoctor.ListIndex)
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.CLI_CODIGO}= " & XN(txtCodigo.text)
    'Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.TUR_DESDE}= TIME (" & Mid(mebHoraD.Text, 1, 2) & "," & Mid(mebHoraD.Text, 4, 2) & ",00)"  '& grdGrilla.RowSel

    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    'Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.Connect = "driver=SQL Server;server=" & SERVIDOR & ";DATABASE=" & BASEDATO & ";User Id=" & USERID & ";Password=" & PASSWORD
    Rep.WindowTitle = "Impresion del Turno"
    Select Case cboDoctor.ItemData(cboDoctor.ListIndex)
        Case 1 ' Silvana
            Rep.ReportFileName = DirReport & "rptTurno_silvana.rpt"
        Case 2 'Lelo
            Rep.ReportFileName = DirReport & "rptTurno_Lelo.rpt"
        Case 16 ' carla gobbi
            Rep.ReportFileName = DirReport & "rptTurno_gobbi.rpt"
        Case 36 ' malena baudo
            Rep.ReportFileName = DirReport & "rptTurno_baudo.rpt"
        Case 21 ' Lorena Sugar
            Rep.ReportFileName = DirReport & "rptTurno_lorenasugar.rpt"
        Case 33 ' Melina Orlietti
            Rep.ReportFileName = DirReport & "rptTurno_melinaorlietti.rpt"
        Case 17 ' Nadia Perassi
            Rep.ReportFileName = DirReport & "rptTurno_Nadia_Perassi.rpt"
        Case 28 ' Jose Boldu
            Rep.ReportFileName = DirReport & "rptTurno_Jose_Boldu.rpt"
        Case 40 ' Flavia Nicolodi
            Rep.ReportFileName = DirReport & "rptTurno_Flavia_Nocolodi.rpt"
        Case 31 ' Rebeca Bonetto
            Rep.ReportFileName = DirReport & "rptTurno_Rebeca_Bonetto.rpt"
        Case 8 ' Carlos Oviedo
            Rep.ReportFileName = DirReport & "rptTurno_Carlos_Oviedo.rpt"
        Case 7 ' Carlos Audisio
            Rep.ReportFileName = DirReport & "rptTurno_Carlos_Audisio.rpt"
        Case 15 ' Robles
            Rep.ReportFileName = DirReport & "rptTurno_Robles.rpt"
        Case Else
            Rep.ReportFileName = DirReport & "rptTurno.rpt"
    End Select
    
    Rep.Action = 1
    
    actualizo_turno_impreso
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)

    
End Function

Private Sub ExportReportToPDF(ReportObject As CRAXDRT.Report, ByVal FileName As String, ByVal ReportTitle As String)
    
    Dim objExportOptions As CRAXDRT.ExportOptions
 
    ReportObject.ReportTitle = ReportTitle
    
    With ReportObject
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = True
    End With
    
    Set objExportOptions = ReportObject.ExportOptions
    
    With objExportOptions
        .DestinationType = crEDTDiskFile
        .DiskFileName = FileName
        .FormatType = crEFTPortableDocFormat
        .PDFExportAllPages = True
    End With
 
    ReportObject.Export False
 
End Sub
 


Private Sub cboMotivo_Click()
    txtMotivo.text = cboMotivo.text
    
End Sub


Private Sub cmdAceptarP_Click()
    'Guardar PROTOCOLO SELECCIONADO en tabla IMAGEN
    Dim i, cont As Integer
    Dim Num As Integer
    cont = 0
    For i = 1 To grdProtocolos.rows - 1
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
            sql = sql & XDQ(MViewFecha.Value) & ","
            sql = sql & grdGrilla.TextMatrix(grdGrilla.RowSel, 9) & ","
            sql = sql & 1 & "," 'SOLO SILVANA ES LA ECOGRAFA
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
    If cont > 0 Then
        MsgBox "Protocolo agregado a la Historia Clinica (Ecografias) del Paciente" & grdGrilla.TextMatrix(grdGrilla.RowSel, 1) & ". ", vbInformation, TIT_MSGBOX
        frmhistoriaclinica.tabhc.Tab = 1
        frmhistoriaclinica.txtCodigo = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
        frmhistoriaclinica.Show vbModal
    End If
End Sub
' ============================================================
' VARIABLE PÚBLICA DEL FORMULARIO (declarar en la sección General)
' Public modoActualizacionTurno As Integer  ' 0 = Creación, 1 = Edición
' ============================================================

Private Sub cmdAgregar_Click()
    Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    Dim años As Integer, edad As Integer
    Dim Fecha As Date
    Dim usuarioCodigoActual As Long
    Dim clave As String
    Dim Frm As New frmIngresarClave
    
    Dim idTurnoHistorico As Long
    Dim accionCodigo As Integer
    Dim bEsInsert As Boolean
    Dim RecAux As ADODB.Recordset
    
    Dim dImporte As Double
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    
    Dim sEstado As String
    
    Dim venCodigo As Integer
    venCodigo = cboDoctor.ItemData(cboDoctor.ListIndex)
    Dim puedeAtender As Boolean
    
    Dim asistio As String
    
    ' CLI_CODIGO del turno seleccionado en la grilla (para validación de solapamiento)
    Dim cliCodigoGrilla As Long
    cliCodigoGrilla = 0
    
    If modoActualizacionTurno = 1 Then
        ' Modo edición: obtener ID_TURNO y CLI_CODIGO de la grilla
        'If grdGrilla.RowSel > 0 Then
        '    If Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 20)) <> "" Then
        '        idTurnoSeleccionado = CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 20))
        '    End If
        '    If Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 9)) <> "" Then
        '        cliCodigoGrilla = CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 9))
        '    End If
        'End If
        
        If idTurnoSeleccionado = 0 Then
            MsgBox "No se pudo obtener el turno seleccionado para editar.", vbExclamation, TIT_MSGBOX
            colocarModoCreacionTurno
            Exit Sub
        End If
    End If
    
    ' Validar disponibilidad del doctor
    puedeAtender = DoctorPuedeAtender(MViewFecha.Value, venCodigo, mebHoraD, mebHoraH)
    
    If puedeAtender = False Then
        MsgBox "El doctor no está disponible en el horario ingresado. Elija otra hora desde y hasta", vbExclamation
        Exit Sub
    End If
    
    ' Validar campos requeridos
    If ValidarTurno = False Then Exit Sub
    
    ' Validar solapamiento con turnos de DISTINTO cliente
    If ValidarSolapamientoTurno(fechaturno.Value, mebHoraD.text, mebHoraH.text, _
                                 venCodigo, CLng(txtCodigo.text), idTurnoSeleccionado) = False Then
        MsgBox "Ya existe un turno de otro paciente en ese horario. Las horas se solapan.", vbCritical, TIT_MSGBOX
        Exit Sub
    End If
    
    ' Confirmar operación
    If modoActualizacionTurno = 1 Then
        If MsgBox("¿Confirma la modificación del Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    Else
        If MsgBox("¿Confirma el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    End If
    
    ' Pedir clave
    Frm.Show vbModal
    clave = Frm.ClaveIngresada
    Unload Frm
    Set Frm = Nothing
    
    If Trim(clave) = "" Then
        MsgBox "Operación cancelada. No se ingresó clave.", vbExclamation
        Exit Sub
    End If
    
    usuarioCodigoActual = ObtenerUsuarioCodigoActual(clave)
    
    If usuarioCodigoActual = 0 Then
        MsgBox "Operación cancelada. Clave inexistente.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo HayErrorTurno
    
    grdGrilla.HighLight = flexHighlightAlways
    
    i = 0
    sHoraDAux = mebHoraD.text
    
    DBConn.BeginTrans
    
    sHoraD = mebHoraD.text
    sHoraD = Mid(sHoraD, 1, 1)
    If sHoraD = "0" Then
        sHoraD = Mid(mebHoraD.text, 2, 4)
    Else
        sHoraD = Trim(mebHoraD.text)
    End If
    
    ' Determinar si tiene orden => estado en espera
    If txtOrden <> "" Then
        asistio = 2
    Else
        asistio = 0
    End If
    
    ' =============================================
    ' MODO CREACIÓN (modoActualizacionTurno = 0)
    ' =============================================
    If modoActualizacionTurno = 0 Then
        bEsInsert = True
        accionCodigo = 1
        
        sql = "INSERT INTO TURNOS"
        sql = sql & " (TUR_FECHA, TUR_HORAD, TUR_HORAH,"
        sql = sql & " VEN_CODIGO, CLI_CODIGO, TUR_MOTIVO, TUR_DRSOLICITA, TUR_OBSERV, TUR_ASISTIO, TUR_OSOCIAL, TUR_TIENEMUTUAL,"
        sql = sql & " TUR_USER,"
        sql = sql & " TUR_FECALTA, TUR_DESDE, TUR_IMPORTE, TUR_ORDEN, TUR_IMPRESO, CREADO_POR)"
        sql = sql & " VALUES ("
        sql = sql & XDQ(fechaturno.Value) & ",'"
        sql = sql & fechaturno.Value & " " & mebHoraD.text & "','"
        sql = sql & fechaturno.Value & " " & mebHoraH.text & "',"
        sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
        sql = sql & XN(txtCodigo) & ","
        sql = sql & XS(txtMotivo) & ","
        sql = sql & XS(txtDrSolicitante) & ","
        sql = sql & XS(txtObservaciones) & ","
        sql = sql & asistio & ","
        If optSI.Value = True Then
            sql = sql & XS(txtOSocial.text) & ","
        Else
            sql = sql & XS("PARTICULAR") & ","
        End If
        If txtOSocial.text <> "" Then
            sql = sql & XN("1") & ","
        Else
            sql = sql & XN("0") & ","
        End If
        sql = sql & User & ","
        sql = sql & XDQ(Date) & ","
        If i = 1 Then
            sql = sql & 1 & ","
        Else
            sql = sql & 0 & ","
        End If
        sql = sql & XN(txtimporte.text) & ","
        sql = sql & XN(txtOrden.text) & ","
        sql = sql & 0 & ","
        sql = sql & usuarioCodigoActual
        sql = sql & ")"
        
        DBConn.Execute sql
        
        ' Obtener el ID_TURNO recién generado
        sql = "SELECT ID_TURNO FROM TURNOS"
        sql = sql & " WHERE TUR_FECHA = " & XDQ(fechaturno.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & mebHoraD.text & "'"
        sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
        sql = sql & " AND CLI_CODIGO = " & XN(txtCodigo)
        sql = sql & " AND DELETED_AT IS NULL"
        Set RecAux = New ADODB.Recordset
        RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not RecAux.EOF Then
            If Not IsNull(RecAux!ID_TURNO) Then
                idTurnoHistorico = RecAux!ID_TURNO
            End If
        End If
        RecAux.Close
        Set RecAux = Nothing
    
    ' =============================================
    ' MODO EDICIÓN (modoActualizacionTurno = 1)
    ' =============================================
    Else
        bEsInsert = False
        accionCodigo = 6
        idTurnoHistorico = idTurnoSeleccionado
        
        ' Si estoy editando y tiene orden y estaba sin asistir, poner en espera
        ' Primero obtengo el asistio actual del turno
        sql = "SELECT TUR_ASISTIO FROM TURNOS WHERE ID_TURNO = " & idTurnoSeleccionado & " AND DELETED_AT IS NULL"
        Set RecAux = New ADODB.Recordset
        RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not RecAux.EOF Then
            If Not IsNull(RecAux!TUR_ASISTIO) Then
                Dim asistioActual As String
                asistioActual = RecAux!TUR_ASISTIO
            End If
        End If
        RecAux.Close
        Set RecAux = Nothing
        
        ' Si tiene orden y estaba sin asistir, poner en espera
        If txtOrden <> "" And asistioActual = "0" Then
            asistio = 2
        ElseIf txtOrden = "" Then
            asistio = 0
        Else
            asistio = asistioActual
        End If
        
        sql = "UPDATE TURNOS SET "
        sql = sql & " CLI_CODIGO = " & XN(txtCodigo.text)
        sql = sql & " ,TUR_HORAD = '" & fechaturno.Value & " " & mebHoraD.text & "'"
        sql = sql & " ,TUR_HORAH = '" & fechaturno.Value & " " & mebHoraH.text & "'"
        sql = sql & " ,TUR_MOTIVO = " & XS(txtMotivo.text)
        sql = sql & " ,TUR_DRSOLICITA = " & XS(txtDrSolicitante.text)
        'sql = sql & " ,TUR_FECALTA = " & XDQ(Date)
        If User <> 99 Then
            sql = sql & " ,TUR_USER = " & User
        End If
        sql = sql & " ,TUR_IMPORTE = " & XN(txtimporte.text)
        If optSI.Value = True Then
            sql = sql & " ,TUR_OSOCIAL = " & XS(txtOSocial.text)
        Else
            sql = sql & " ,TUR_OSOCIAL = " & XS("PARTICULAR")
        End If
        If txtOSocial.text <> "" Then
            sql = sql & " ,TUR_TIENEMUTUAL = " & XN(1)
        Else
            sql = sql & " ,TUR_TIENEMUTUAL = " & XN(0)
        End If
        sql = sql & " ,TUR_ORDEN = " & XN(txtOrden.text)
        sql = sql & " ,TUR_OBSERV = " & XS(txtObservaciones.text)
        sql = sql & " ,TUR_ASISTIO = " & asistio
        sql = sql & " ,TUR_FECHA = " & XDQ(fechaturno.Value)
        sql = sql & " ,VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
        sql = sql & " ,UPDATED_AT = GETDATE()"
        sql = sql & " ,ACTUALIZADO_POR = " & usuarioCodigoActual
        sql = sql & " WHERE ID_TURNO = " & idTurnoSeleccionado
        sql = sql & " AND DELETED_AT IS NULL"
        
        DBConn.Execute sql
    End If
    
    ' =============================================
    ' REGISTRAR EN HISTORICO_TURNO
    ' =============================================
    If idTurnoHistorico > 0 Then
        sql = "INSERT INTO HISTORICO_TURNO (ID_TURNO, ACCION_CODIGO, IMPORTE, HORA_DESDE, HORA_HASTA, ORDEN, MOTIVO, USU_NOMBRE, ESTADO_CODIGO, LLA_CODIGO)"
        sql = sql & " VALUES (" & idTurnoHistorico & ", " & accionCodigo & ", "
        sql = sql & XN(txtimporte.text) & ", "
        sql = sql & "'" & fechaturno.Value & " " & mebHoraD.text & "'" & ", "
        sql = sql & "'" & fechaturno.Value & " " & mebHoraH.text & "'" & ", "
        sql = sql & XS(txtOrden.text) & ", "
        sql = sql & XS(txtMotivo.text) & ", "
        sql = sql & XS(mNomUser) & ", "
        sql = sql & XN(asistio) & ", "
        sql = sql & usuarioCodigoActual & ")"
        DBConn.Execute sql
    End If
    
    mebHoraD.text = sHoraDAux
    
    ' =============================================
    ' CALCULAR Y ACTUALIZAR EDAD DEL PACIENTE
    ' =============================================
    Fecha = fechaturno.Value
    sql = "SELECT CLI_CUMPLE"
    sql = sql & " FROM CLIENTE"
    sql = sql & " WHERE CLI_CODIGO = " & XN(txtCodigo.text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Not (IsNull(rec!CLI_CUMPLE)) Then
        If rec.EOF = False Then
            años = Year(Date) - Year(rec!CLI_CUMPLE)
            If Month(Fecha) < Month(rec!CLI_CUMPLE) Then años = años - 1
            If Month(Now) = Month(rec!CLI_CUMPLE) And Day(Fecha) < Day(rec!CLI_CUMPLE) Then años = años - 1
            edad = años
        End If
    Else
        edad = 0
    End If
    rec.Close
    sql = "UPDATE CLIENTE SET"
    sql = sql & " CLI_EDAD=" & edad
    sql = sql & " WHERE CLI_CODIGO=" & txtCodigo.text
    DBConn.Execute sql
    
    DBConn.CommitTrans
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    
    ' =============================================
    ' VOLVER A MODO CREACIÓN
    ' =============================================
    colocarModoCreacionTurno

    If MsgBox("¿Desea imprimir el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
        LimpiarTurno
        Exit Sub
    End If
    
    ImprimirTurno
    LimpiarTurno
        
    Exit Sub
    
HayErrorTurno:
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    If Rec1.State = 1 Then Rec1.Close
    If Not RecAux Is Nothing Then
        If RecAux.State = 1 Then RecAux.Close
    End If
    DBConn.RollbackTrans
    modoActualizacionTurno = 0  ' Volver a modo creación ante error
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub


' ============================================================
' FUNCIÓN DE VALIDACIÓN DE SOLAPAMIENTO
' Valida que no exista un turno de DISTINTO CLI_CODIGO
' con horas solapadas para el mismo VEN_CODIGO y fecha.
' En modo edición, excluye el turno actual (idTurnoExcluir > 0).
' Retorna True si NO hay solapamiento (puede continuar).
' Retorna False si HAY solapamiento (debe frenar).
' ============================================================
Private Function ValidarSolapamientoTurno(ByVal sFecha As String, _
                                           ByVal sHoraD As String, _
                                           ByVal sHoraH As String, _
                                           ByVal venCodigo As Integer, _
                                           ByVal cliCodigo As Long, _
                                           ByVal idTurnoExcluir As Long) As Boolean
    Dim recVal As New ADODB.Recordset
    Dim sqlVal As String
    
    ' Armar las fechas/hora completas para comparar
    Dim sDesde As String
    Dim sHasta As String
    sDesde = sFecha & " " & sHoraD
    sHasta = sFecha & " " & sHoraH
    
    ' Buscar turnos de DISTINTO cliente cuyas horas se solapen
    ' Solapamiento: TUR_HORAD < @HoraHasta AND TUR_HORAH > @HoraDesde
    sqlVal = "SELECT ID_TURNO FROM TURNOS"
    sqlVal = sqlVal & " WHERE TUR_FECHA = " & XDQ(sFecha)
    sqlVal = sqlVal & " AND VEN_CODIGO = " & venCodigo
    sqlVal = sqlVal & " AND CLI_CODIGO <> " & cliCodigo
    sqlVal = sqlVal & " AND TUR_HORAD < '" & sHasta & "'"
    sqlVal = sqlVal & " AND TUR_HORAH > '" & sDesde & "'"
    sqlVal = sqlVal & " AND DELETED_AT IS NULL"
    
    ' En modo edición, excluir el turno que estoy editando
    If idTurnoExcluir > 0 Then
        sqlVal = sqlVal & " AND ID_TURNO <> " & idTurnoExcluir
    End If
    
    recVal.Open sqlVal, DBConn, adOpenStatic, adLockOptimistic
    
    If recVal.EOF Then
        ' No hay solapamiento
        ValidarSolapamientoTurno = True
    Else
        ' Hay solapamiento
        ValidarSolapamientoTurno = False
    End If
    
    recVal.Close
    Set recVal = Nothing
End Function

Private Sub cmdatendido_Click()
    '[NUEVO]
    Dim idTurnoHistorico As Long
    Dim RecAux As ADODB.Recordset
    Dim dImporte As Double
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    Dim sMotivo As String
    Dim sOrden As String
    Dim sUsuNombre As String
    Dim sEstado As String
    '[FIN NUEVO]
    
    If grdGrilla.RowSel <> 0 Then
        'atendido
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = 1
        cambiocolor 1
        
    
        'Actualizo la Base de Datos
        sql = "UPDATE TURNOS SET "
        sql = sql & " TUR_ASISTIO =" & grdGrilla.TextMatrix(grdGrilla.RowSel, 10)
        sql = sql & " WHERE "
        sql = sql & " TUR_FECHA = " & XDQ(fechaturno.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        DBConn.Execute sql
        
        '[NUEVO] CREO HISTORICO DE PENDIENTE
        sql = "SELECT ID_TURNO, TUR_IMPORTE, TUR_HORAD, TUR_HORAH, TUR_ORDEN, TUR_MOTIVO, TUR_ASISTIO FROM TURNOS"
        sql = sql & " WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        Set RecAux = New ADODB.Recordset
        RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not RecAux.EOF Then
            If Not IsNull(RecAux!ID_TURNO) Then
                idTurnoHistorico = RecAux!ID_TURNO
            End If
            If Not IsNull(RecAux!TUR_IMPORTE) Then dImporte = RecAux!TUR_IMPORTE
            If Not IsNull(RecAux!TUR_HORAD) Then sHoraDesde = RecAux!TUR_HORAD
            If Not IsNull(RecAux!TUR_HORAH) Then sHoraHasta = RecAux!TUR_HORAH
            If Not IsNull(RecAux!tur_orden) Then sOrden = RecAux!tur_orden
            If Not IsNull(RecAux!TUR_MOTIVO) Then sMotivo = RecAux!TUR_MOTIVO
            If Not IsNull(RecAux!TUR_ASISTIO) Then sEstado = RecAux!TUR_ASISTIO
        End If
        RecAux.Close
        Set RecAux = Nothing
        
        If idTurnoHistorico > 0 Then
            sql = "INSERT INTO HISTORICO_TURNO (ID_TURNO, ACCION_CODIGO, IMPORTE, HORA_DESDE, HORA_HASTA, ORDEN, MOTIVO, ESTADO_CODIGO, USU_NOMBRE)"
            sql = sql & " VALUES (" & idTurnoHistorico & ", " & 2 & ", " 'atendido
            sql = sql & dImporte & ", "
            sql = sql & XS(sHoraDesde) & ", "
            sql = sql & XS(sHoraHasta) & ", "
            sql = sql & XS(sOrden) & ", "
            sql = sql & XS(sMotivo) & ", "
            sql = sql & XN(sEstado) & ", "
            sql = sql & XS(mNomUser) & ") "
            DBConn.Execute sql
        End If
        '[FIN NUEVO]
    End If
End Sub

Private Sub CmdBuscar_Click()
    frmBuscarTurnos.Show vbModal
End Sub
Private Sub LimpiarTurno()
    fraprotocolos.Visible = False
    txtBuscaCliente.text = ""
    txtBuscaCliente.ToolTipText = ""
    txtCodigo.text = ""
    txtTelefono.text = ""
    txtcelular.text = ""
    txtOSocial.text = ""
    txtBuscarCliDescri.text = ""
    txtMotivo.text = ""
    txtDrSolicitante.text = ""
    txtObservaciones.text = ""
    'cboDesde.ListIndex = -1
    'cbohasta.ListIndex = -1
    mebHoraD.text = "__:__"
    mebHoraH.text = "__:__"
    txtimporte.text = "0,00"
    txtBuscaCliente.SetFocus
    cmdImpTurno.Enabled = False
    cmdCopiar.Enabled = True
    cmdCortar.Enabled = False
    cmdProtocolos.Enabled = False
    optSI.Enabled = True
    cboMotivo.ListIndex = -1
    txtOrden.text = ""
    cancelarEdicionTurno
End Sub
Private Sub cancelarEdicionTurno()
    idTurnoSeleccionado = 0
    cmdAgregar.Caption = "&" & "Agregar"
    modoActualizacionTurno = 0
End Sub

Private Sub cmdCerrarDisponibilidad_Click()
    fraConfiguracionHorarios.Visible = False
End Sub

Private Sub cmdCerrarFraListaEstudios_Click()
    fraListaEstudios.Visible = False
End Sub
Private Sub CopiarTurno()
optNO.Enabled = True
    optSI.Enabled = True
    If grdGrilla.rows > 1 Then
       If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
           txtBuscaCliente.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 11)
           'txtBuscaCliente_LostFocus
           txtCodigo.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
           txtBuscarCliDescri.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
           txtTelefono.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 3)
           txtcelular.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 4)
           txtOSocial.text = BuscarOSocial(txtCodigo.text) 'grdGrilla.TextMatrix(grdGrilla.RowSel, 5)
           
           'verifico si el paciente tiene mutual
           ' If Chk0(grdGrilla.TextMatrix(grdGrilla.RowSel, 13)) <> 1 Then 'si no tiene mutual el paciente
          If txtOSocial.text = "" Then
               optSI.Enabled = False
               optNO.Value = True
            Else
                optSI.Enabled = True
                optSI.Value = True
           End If
           
           'veriifco si es con mutual el turno
           If grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = "PARTICULAR" Then
               optNO.Value = True
               'optSI.Enabled = True
           Else
                optSI.Enabled = True
                optSI.Value = True
           End If
          
           txtMotivo.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 6)
           txtDrSolicitante.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 7)
           BuscaDescriProx Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cboDesde
           BuscaDescriProx Right(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cbohasta
           
           mebHoraD.text = Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5)
           mebHoraH.text = Right(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5)
           
           'If Chk0(grdGrilla.TextMatrix(grdGrilla.RowSel, 13)) = 0 Then
           '    optSI.Enabled = False
           'End If

           If mNomUser = "DIGOR" Or mNomUser = "SILVANA" Then
               txtimporte.text = Valido_Importe(grdGrilla.TextMatrix(grdGrilla.RowSel, 14))
           Else
               txtimporte.text = "0,00"
           End If
           txtOrden.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 15)
           
           txtObservaciones.text = grdGrilla.TextMatrix(grdGrilla.RowSel, 19)
           
           cmdImpTurno.Enabled = True
           cmdProtocolos.Enabled = True
           cmdCortar.Enabled = True
           cmdCopiar.Enabled = True
       Else
           If txtBuscaCliente.text <> "" Then
               MViewFecha.Value = Date
               txtBuscaCliente.text = ""
               txtCodigo.text = ""
               txtBuscarCliDescri.text = ""
               txtTelefono.text = ""
               txtOSocial.text = ""
               txtMotivo.text = ""
               cboDesde.ListIndex = -1
               cbohasta.ListIndex = -1
               mebHoraD.text = ""
               mebHoraH.text = ""
               
               txtimporte.text = "0,00"
           End If
       End If
    End If
End Sub

Private Sub cmdCopiar_Click()
    Dim turnovalido As Boolean
    turnovalido = validarTurnoSeleccionado
    idTurnoSeleccionado = 0
    If turnovalido = False Then
        MsgBox "Seleccione un turno cargado para Copiar", vbExclamation, "Información"
        Exit Sub
    End If
    colocarModoCreacionTurno
    CopiarTurno
End Sub

Private Sub cmdCortar_Click()
    If MsgBox("Esta a punto de Cortar los " & lbldiaTurno.Caption & " " & Chr(13) & " del Doctor: " & cboDoctor.text & _
    " Â¿Confirma Cortar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    sAction = "CORTAR"
    dFechaCopy = MViewFecha.Value
    nDoctorCopy = cboDoctor.ItemData(cboDoctor.ListIndex)
    sNameDoctorCopy = cboDoctor.text
End Sub

Private Sub cmdDrive_Click()
    If linkProtocolos <> "" Then
        Shell "cmd /c start " & linkProtocolos, vbNormalFocus
    Else
        MsgBox "No tienes configurado el link a los protocolos. Por favor contacta al administrador", vbExclamation, "Información"
    End If
End Sub

Private Sub cmdEditar_Click()
    Dim turnovalido As Boolean
    idTurnoSeleccionado = 0
    turnovalido = validarTurnoSeleccionado
    If turnovalido = False Then
        MsgBox "Seleccione un turno cargado para Editar", vbExclamation, "Información"
        Exit Sub
    End If
    idTurnoSeleccionado = CLng(grdGrilla.TextMatrix(grdGrilla.RowSel, 20))
    colocarModoEdicionTurno
    CopiarTurno
End Sub
Private Sub colocarModoCreacionTurno()
    If validarTurnoSeleccionado = False Then Exit Sub
    modoActualizacionTurno = 0 'MODO creacion
    cmdAgregar.Caption = "&" & "Agregar"
End Sub
Private Sub colocarModoEdicionTurno()
    modoActualizacionTurno = 1 'MODO EDICION
    cmdAgregar.Caption = "&" & "Editar"
End Sub
Private Function validarTurnoSeleccionado() As Boolean
Dim CodigoCli As String
CodigoCli = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
validarTurnoSeleccionado = CodigoCli <> ""
End Function

Private Sub cmdespera_Click()

    '[NUEVO]
    Dim idTurnoHistorico As Long
    Dim RecAux As ADODB.Recordset
    Dim dImporte As Double
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    Dim sMotivo As String
    Dim sOrden As String
    Dim sUsuNombre As String
    Dim sEstado As String
    
    '[FIN NUEVO]
    
    If grdGrilla.RowSel <> 0 Then
        'en espera
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = 2
        cambiocolor 2
        
        'Actualizo la Base de Datos
        sql = "UPDATE TURNOS SET "
        sql = sql & " TUR_ASISTIO =" & grdGrilla.TextMatrix(grdGrilla.RowSel, 10)
        sql = sql & " WHERE "
        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        DBConn.Execute sql
        
        
        '[NUEVO] CREO HISTORICO DE PENDIENTE
        sql = "SELECT ID_TURNO,TUR_IMPORTE, TUR_HORAD, TUR_ORDEN, TUR_MOTIVO, TUR_HORAH, TUR_ASISTIO FROM TURNOS"
        sql = sql & " WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        Set RecAux = New ADODB.Recordset
        RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not RecAux.EOF Then
            If Not IsNull(RecAux!ID_TURNO) Then
                idTurnoHistorico = RecAux!ID_TURNO
            End If
            If Not IsNull(RecAux!TUR_IMPORTE) Then dImporte = RecAux!TUR_IMPORTE
            If Not IsNull(RecAux!TUR_HORAD) Then sHoraDesde = RecAux!TUR_HORAD
            If Not IsNull(RecAux!TUR_HORAH) Then sHoraHasta = RecAux!TUR_HORAH
            If Not IsNull(RecAux!tur_orden) Then sOrden = RecAux!tur_orden
            If Not IsNull(RecAux!TUR_MOTIVO) Then sMotivo = RecAux!TUR_MOTIVO
            If Not IsNull(RecAux!TUR_ASISTIO) Then sEstado = RecAux!TUR_ASISTIO
            sUsuNombre = mNomUser
        End If
        RecAux.Close
        Set RecAux = Nothing
        
        If idTurnoHistorico > 0 Then
            sql = "INSERT INTO HISTORICO_TURNO (ID_TURNO, ACCION_CODIGO, IMPORTE, HORA_DESDE, HORA_HASTA, ORDEN, MOTIVO, ESTADO_CODIGO, USU_NOMBRE)"
            sql = sql & " VALUES (" & idTurnoHistorico & ", " & 3 & ", " 'en espera
            sql = sql & dImporte & ", "
            sql = sql & XS(sHoraDesde) & ", "
            sql = sql & XS(sHoraHasta) & ", "
            sql = sql & XS(sOrden) & ", "
            sql = sql & XS(sMotivo) & ", "
            sql = sql & XN(sEstado) & ", "
            sql = sql & XS(sUsuNombre) & ") "
            DBConn.Execute sql
        End If
        '[FIN NUEVO]
    End If
End Sub
'FUNCIONES AUXILIARES PARA CSV RECORDATORIOS
Function LimpiarNombreArchivo(Txt As String) As String
    Txt = Replace(Txt, " ", "_")
    Txt = Replace(Txt, "/", "")
    Txt = Replace(Txt, "\", "")
    Txt = Replace(Txt, ":", "")
    LimpiarNombreArchivo = Txt
End Function
Function LimpiarCSV(Txt As String) As String
    Txt = Replace(Txt, ",", " ")
    LimpiarCSV = Txt
End Function

Private Sub cmdExcel_Click()
Dim rs As New ADODB.Recordset
Dim archivo As Integer
Dim ruta As String
Dim linea As String

Dim medicoNombre As String
Dim medicoCodigo As Long
Dim fechaSeleccionada As Date

' Datos desde UI (adaptar a tus controles)
medicoCodigo = cboDoctor.ItemData(cboDoctor.ListIndex)
medicoNombre = cboDoctor.text
fechaSeleccionada = MViewFecha.Value

'fechaturno = "'" & Format(MViewFecha.Value, "dd/mm/yyyy") & "'"
    
ruta = TURNOS_EXPORTADOS_DIR

sql = "SELECT C.CLI_TELEFONO, C.CLI_CELULAR, C.CLI_RAZSOC, T.TUR_FECHA, T.TUR_HORAD, V.VEN_NOMBRE " & _
      "FROM TURNOS T " & _
      "INNER JOIN CLIENTE C ON T.CLI_CODIGO = C.CLI_CODIGO " & _
      "INNER JOIN VENDEDOR V ON T.VEN_CODIGO = V.VEN_CODIGO " & _
      "WHERE T.DELETED_AT IS NULL " & _
      "AND T.VEN_CODIGO = " & medicoCodigo & " " & _
      "AND T.TUR_FECHA = " & XDQ(fechaSeleccionada) & " " & _
      "ORDER BY T.TUR_HORAD"

rs.Open sql, DBConn

If rs.EOF Then
    MsgBox "No hay turnos para ese médico en esa fecha.", vbInformation
    Exit Sub
End If

' Crear archivo
archivo = FreeFile

Open ruta & "Turnos_" & _
     LimpiarNombreArchivo(medicoNombre) & "_" & _
     Format(fechaSeleccionada, "dd-mm-yyyy") & ".csv" _
     For Output As #archivo

' Header
Print #archivo, "Nombre,Telefono,Fecha,Hora,Medico,Enviado"

Do While Not rs.EOF

    linea = LimpiarCSV(rs!CLI_RAZSOC) & "," & _
            obtenerCelular(ChkNull(rs!CLI_CELULAR)) & "," & _
            Format(rs!TUR_FECHA, "YYYY-MM-DD") & "," & _
            Format(rs!TUR_HORAD, "hh:nn") & "," & _
            LimpiarCSV(rs!VEN_NOMBRE) & "," & _
            "NO"
            
    Print #archivo, linea

    rs.MoveNext

Loop

Close #archivo
rs.Close

MsgBox "Archivo generado correctamente.", vbInformation
End Sub

Private Sub cmdExportarTurno_Click()
    Dim archivo As Integer
    Dim ruta As String
    Dim linea As String
    Dim medicoNombre As String
    Dim fechaSeleccionada As Date
    Dim filaSeleccionada As Long
    Dim nombreCliente As String
    Dim razSoc As String
    Dim telefonoCliente As String
    Dim fechaturno As String
    Dim horaTurno As String
    Dim nombreMedico As String
    
    ' Validar que hay una fila seleccionada en la grilla
    If grdGrilla.row < grdGrilla.FixedRows Then
        MsgBox "Debe seleccionar un turno de la grilla.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    ' Validar que la grilla tiene datos
    If grdGrilla.rows <= grdGrilla.FixedRows Then
        MsgBox "No hay turnos disponibles en la grilla.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    ' Obtener fila seleccionada
    filaSeleccionada = grdGrilla.row
    
    ' Extraer datos de la grilla (ajustar índices de columna según tu grilla)
    ' Ejemplo asumiendo columnas: 0=Nombre, 1=Teléfono, 2=Fecha, 3=Hora, 4=Médico
     horaTurno = Left(grdGrilla.TextMatrix(filaSeleccionada, 0), 5)      ' HORAS
    nombreCliente = grdGrilla.TextMatrix(filaSeleccionada, 1)    ' PACIENTE
    razSoc = grdGrilla.TextMatrix(filaSeleccionada, 1)    ' PACIENTE
    telefonoCliente = grdGrilla.TextMatrix(filaSeleccionada, 3)  ' CELULAR/TELEFONO
    
    ' Datos desde UI
    nombreMedico = cboDoctor.text
    fechaSeleccionada = MViewFecha.Value
    
    ' Crear ruta y archivo
    ruta = TURNOS_EXPORTADOS_DIR
    archivo = FreeFile
    
    
    Open ruta & "Turno_" & _
         LimpiarNombreArchivo(nombreCliente) & "_" & _
         Format(fechaSeleccionada, "dd-mm-yyyy") & ".csv" _
         For Output As #archivo
    
    ' Header
    Print #archivo, "Nombre,Telefono,Fecha,Hora,Medico,Enviado"
    
    ' Escribir solo la fila seleccionada
    linea = LimpiarCSV(razSoc) & "," & _
            obtenerCelular(telefonoCliente) & "," & _
            Format(fechaSeleccionada, "YYYY-MM-DD") & "," & _
            Format(horaTurno, "hh:nn") & "," & _
            LimpiarCSV(nombreMedico) & "," & _
            "NO"
            
    Print #archivo, linea
    
    Close #archivo
    
    MsgBox "Archivo generado correctamente para el turno seleccionado.", vbInformation, "Éxito"
End Sub
Private Sub cmdImpTurno_Click()
    If txtBuscaCliente.text <> "" Then
        ImprimirTurno
    Else
        MsgBox "Seleccione un turno a imprimir", vbInformation, TIT_MSGBOX
    End If
End Sub

Private Sub cmdInforTurno_Click()

    Dim Frm As New frmAuditoriaTurno

    Dim idTurno As Double
    Dim doctor As String
    Dim Fecha As String
    Dim paciente As String
    
    If grdGrilla.row < grdGrilla.FixedRows Then
        MsgBox "Debe seleccionar un turno de la grilla.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 20) = "" Then
        MsgBox "El turno no tiene historial registrado.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    ' 1. Obtener datos de la grilla
    idTurno = grdGrilla.TextMatrix(grdGrilla.RowSel, 20)
    doctor = cboDoctor.text
    Fecha = MViewFecha.Value
    paciente = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)

    ' 4. Mostrar modal
    Call Frm.CargarDatos(idTurno, paciente, Fecha, doctor)

    Frm.Show vbModal

End Sub

Private Sub cmdLibres_Click()
modoVista = 2
BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex), modoVista
End Sub

Private Sub CmdNuevo_Click()
    LimpiarTurno
    MViewFecha.Value = Date
    fechaturno.Value = Date
    'If User <> 99 Then
    '    Call BuscaCodigoProxItemData(XN(User), cboDoctor)
    'Else
    '    cboDoctor.ListIndex = 0
    'End If
End Sub

'Private Sub cmdProtocolos_Click()
'    Dim DIA As Integer
'    Dim sDiaTurno As String
'    DIA = Weekday(dFechaCopy, vbMonday)
'    sDiaTurno = "Turnos del dia " & WeekdayName(DIA, False) & " " & Day(dFechaCopy) & " de " & MonthName(Month(dFechaCopy), False) & " de " & Year(dFechaCopy)
'
'    If sAction = "CORTAR" Then
'        For i = 1 To grdGrilla.Rows - 1
'            If grdGrilla.TextMatrix(i, 1) <> "" Then
'                Exit For
'            End If
'        Next
'        If i < grdGrilla.Rows - 1 Then
'            If MsgBox("Hay Turnos previamente cargados en este dia que se eliminaran si realiza esta acciÃ³n." & Chr(13) & _
'            " Â¿Confirma eliminar estos Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'            sql = "DELETE FROM TURNOS WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
'            sql = sql & " AND VEN_CODIGO =" & cboDoctor.ItemData(cboDoctor.ListIndex)
'            DBConn.Execute sql
'            LimpiarGrilla
'        End If
'
'         If MsgBox("Esta a punto de Pegar los " & sDiaTurno & " " & Chr(13) & "previamente cortados del Doctor: " & sNameDoctorCopy & _
'        " " & Chr(13) & "Â¿Confirma Pegar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'        sql = "UPDATE TURNOS SET"
'        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
'        sql = sql & ", VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
'        sql = sql & " WHERE TUR_FECHA = " & XDQ(dFechaCopy)
'        sql = sql & " AND VEN_CODIGO = " & XN(nDoctorCopy)
'        DBConn.Execute sql
'
'    Else
'
'        If sAction = "COPIAR" Then
'            For i = 1 To grdGrilla.Rows - 1
'                If grdGrilla.TextMatrix(i, 1) <> "" Then
'                    Exit For
'                End If
'            Next
'            If i < grdGrilla.Rows - 1 Then
'                If MsgBox("Hay Turnos previamente cargados en este dia que se eliminaran si realiza esta acciÃ³n." & Chr(13) & _
'                " Â¿Confirma eliminar estos Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'                sql = "DELETE FROM TURNOS WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
'                sql = sql & " AND VEN_CODIGO =" & cboDoctor.ItemData(cboDoctor.ListIndex)
'                DBConn.Execute sql
'                LimpiarGrilla
'            End If
'
'
'
'             If MsgBox("Esta a punto de Pegar los " & sDiaTurno & " " & Chr(13) & "previamente copiados del Doctor: " & sNameDoctorCopy & _
'            " " & Chr(13) & "Â¿Confirma Pegar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'            sql = "SELECT * FROM TURNOS WHERE TUR_FECHA = " & XDQ(dFechaCopy)
'            sql = sql & "AND VEN_CODIGO = " & XN(nDoctorCopy)
'            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
'            If rec.EOF = False Then
'                Do While rec.EOF = False
'                    sql = "INSERT INTO TURNOS"
'                    sql = sql & " (TUR_FECHA, TUR_HORAD,TUR_HORAH,"
'                    sql = sql & " VEN_CODIGO,CLI_CODIGO,"
'                    If Not IsNull(rec!TUR_MOTIVO) Then
'                        sql = sql & " TUR_MOTIVO,"
'                    End If
'                    If Not IsNull(rec!TUR_OSOCIAL) Then
'                        sql = sql & " TUR_OSOCIAL,"
'                    End If
'                    sql = sql & "TUR_ASISTIO)"
'                    sql = sql & " VALUES ("
'                    sql = sql & XDQ(MViewFecha.Value) & ",#"
'                    sql = sql & rec!TUR_HORAD & "#,#"
'                    sql = sql & rec!TUR_HORAH & "#,"
'                    sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
'                    sql = sql & XN(rec!CLI_CODIGO) & ","
'                    If Not IsNull(rec!TUR_MOTIVO) Then
'                        sql = sql & XS(rec!TUR_MOTIVO) & ","
'                    End If
'                    If Not IsNull(rec!TUR_OSOCIAL) Then
'                        sql = sql & XS(rec!TUR_OSOCIAL) & ","
'                    End If
'                    sql = sql & 0 & ")"
'
'                    DBConn.Execute sql
'
'                    rec.MoveNext
'                Loop
'            End If
'            rec.Close
'
'        End If
'    End If
'    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
'    sAction = ""
'    dFechaCopy = ""
'    nDoctorCopy = ""
'    sNameDoctorCopy = ""
'End Sub

Private Sub cmdNuevoPaciente_Click()
    'If txtCodigo.Text = "" Then
        vMode = 1
        'gPaciente = "" 'txtCodigo.Text
        vDNI = txtBuscaCliente.text
        ABMClientes.Show vbModal
        txtBuscaCliente.SetFocus
    'End If
End Sub

Private Sub cmdOcultar_Click()
    If txtTotal.Visible = True Then
        lbltotal.Visible = False
        txtTotal.Visible = False
    Else
        lbltotal.Visible = True
        txtTotal.Visible = True
    End If
End Sub

Private Sub cmdpendiente_Click()
    '[NUEVO]
    Dim idTurnoHistorico As Long
    Dim RecAux As ADODB.Recordset
    Dim dImporte As Double
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    Dim sMotivo As String
    Dim sOrden As String
    Dim sUsuNombre As String
    Dim sEstado As String

    '[FIN NUEVO]
    
    If grdGrilla.RowSel <> 0 Then
        'pendiente
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = 0
        cambiocolor 0
        
        'Actualizo la Base de Datos
        sql = "UPDATE TURNOS SET "
        sql = sql & " TUR_ASISTIO =" & grdGrilla.TextMatrix(grdGrilla.RowSel, 10)
        sql = sql & " WHERE "
        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        DBConn.Execute sql
        
        '[NUEVO] CREO HISTORICO DE PENDIENTE
        sql = "SELECT ID_TURNO, TUR_IMPORTE, TUR_HORAD, TUR_HORAH, TUR_ORDEN, TUR_MOTIVO, TUR_ASISTIO FROM TURNOS"
        sql = sql & " WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        Set RecAux = New ADODB.Recordset
        RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not RecAux.EOF Then
            If Not IsNull(RecAux!ID_TURNO) Then
                idTurnoHistorico = RecAux!ID_TURNO
            End If
            If Not IsNull(RecAux!TUR_IMPORTE) Then dImporte = RecAux!TUR_IMPORTE
            If Not IsNull(RecAux!TUR_HORAD) Then sHoraDesde = RecAux!TUR_HORAD
            If Not IsNull(RecAux!TUR_HORAH) Then sHoraHasta = RecAux!TUR_HORAH
            If Not IsNull(RecAux!tur_orden) Then sOrden = RecAux!tur_orden
            If Not IsNull(RecAux!TUR_MOTIVO) Then sMotivo = RecAux!TUR_MOTIVO
            If Not IsNull(RecAux!TUR_ASISTIO) Then sEstado = RecAux!TUR_ASISTIO
        End If
        RecAux.Close
        Set RecAux = Nothing
        
        If idTurnoHistorico > 0 Then
            sql = "INSERT INTO HISTORICO_TURNO (ID_TURNO, ACCION_CODIGO, IMPORTE, HORA_DESDE, HORA_HASTA, ORDEN, MOTIVO, ESTADO_CODIGO, USU_NOMBRE)"
            sql = sql & " VALUES (" & idTurnoHistorico & ", " & 4 & ", " 'pendiente
            sql = sql & dImporte & ", "
            sql = sql & XS(sHoraDesde) & ", "
            sql = sql & XS(sHoraHasta) & ", "
            sql = sql & XS(sOrden) & ", "
            sql = sql & XS(sMotivo) & ", "
            sql = sql & XS(sEstado) & ", "
            sql = sql & XS(mNomUser) & ") "
            DBConn.Execute sql
        End If
        '[FIN NUEVO]
        
    End If
End Sub

Private Sub cmdProtocolos_Click()
    fraprotocolos.Visible = True
    grdProtocolos.SetFocus
    grdProtocolos.rows = 1
    cargo_protocolos
End Sub
Private Function ObtenerUsuarioCodigoActual(clave As String) As Long

    Dim Item As Variant
    Dim obj As ClsLLave

    If dictLlaves Is Nothing Then
        MsgBox "Las llaves no están cargadas.", vbCritical
        Exit Function
    End If

    For Each Item In dictLlaves.Items
        Set obj = Item

        If Trim(obj.Valor) = Trim(clave) And Item.activa = "Verdadero" Then
            ObtenerUsuarioCodigoActual = obj.Codigo
            Exit Function
        End If
    Next

    ObtenerUsuarioCodigoActual = 0 ' no encontrado

End Function

Private Sub cmdQuitar_Click()
    'Controlar que se pueda eliminar el turno
    'Borrar de la Grilla
    'Borrar de la BD
    Dim clave As String
    Dim usuarioActual As Long
    Dim Frm As New frmIngresarClave
    '[NUEVO]
    Dim idTurnoHistorico As Long
    Dim RecAux As ADODB.Recordset
    Dim dImporte As Double
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    Dim sMotivo As String
    Dim sOrden As String
    Dim sUsuNombre As String
    Dim sEstado As String
    
    '[FIN NUEVO]
    
    If txtCodigo.text <> "" Then
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
            ' 1. Pedir confirmación
            If MsgBox("¿Confirma Eiminar el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
            
            ' 2. Pedir clave
            Frm.Show vbModal
            
            clave = Frm.ClaveIngresada
            
            Unload Frm
            Set Frm = Nothing
            
            If Trim(clave) = "" Then
                MsgBox "Operación cancelada. No se ingresó clave.", vbExclamation
                Exit Sub
            End If
            
            ' 3. Buscar clave en diccionario
            usuarioActual = ObtenerUsuarioCodigoActual(clave)
            
            If usuarioActual = 0 Then
                MsgBox "Operación cancelada. Clave inexistente.", vbExclamation
                Exit Sub
            End If
            
            '[NUEVO] Obtengo ID_TURNO ANTES del borrado lógico, mientras DELETED_AT aún es NULL
            sql = "SELECT ID_TURNO,TUR_IMPORTE, TUR_HORAD, TUR_HORAH, TUR_MOTIVO, TUR_ORDEN,TUR_ASISTIO FROM TURNOS"
            sql = sql & " WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
            sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
            sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
            sql = sql & " AND CLI_CODIGO = " & grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
            sql = sql & " AND DELETED_AT IS NULL"
            Set RecAux = New ADODB.Recordset
            RecAux.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Not RecAux.EOF Then
                If Not IsNull(RecAux!ID_TURNO) Then
                    idTurnoHistorico = RecAux!ID_TURNO
                End If
                If Not IsNull(RecAux!TUR_IMPORTE) Then dImporte = RecAux!TUR_IMPORTE
                If Not IsNull(RecAux!TUR_HORAD) Then sHoraDesde = RecAux!TUR_HORAD
                If Not IsNull(RecAux!TUR_HORAH) Then sHoraHasta = RecAux!TUR_HORAH
                If Not IsNull(RecAux!tur_orden) Then sOrden = RecAux!tur_orden
                If Not IsNull(RecAux!TUR_MOTIVO) Then sMotivo = RecAux!TUR_MOTIVO
                If Not IsNull(RecAux!TUR_ASISTIO) Then sEstado = RecAux!TUR_ASISTIO
            End If
            RecAux.Close
            Set RecAux = Nothing
            '[FIN NUEVO]
            
            ' 5. Soft delete
            sql = "UPDATE TURNOS SET"
            sql = sql & " DELETED_AT = GETDATE(),"
            sql = sql & " BORRADO_POR = " & usuarioActual
            sql = sql & " WHERE"
            sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
            sql = sql & " AND TUR_HORAD = '" & fechaturno.Value & " " & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "'"
            sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
            sql = sql & " AND CLI_CODIGO = " & grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
            sql = sql & " AND DELETED_AT IS NULL" 'importante para no re-borrar
            
            DBConn.Execute sql
            
            '[NUEVO] Registro en histórico luego del borrado
            If idTurnoHistorico > 0 Then
               sql = "INSERT INTO HISTORICO_TURNO (ID_TURNO, ACCION_CODIGO, IMPORTE, HORA_DESDE, HORA_HASTA, ORDEN, MOTIVO, USU_NOMBRE, ESTADO_CODIGO, LLA_CODIGO)"
                sql = sql & " VALUES (" & idTurnoHistorico & ", " & 5 & ", " 'BORRADO
                sql = sql & dImporte & ", "
                sql = sql & XS(sHoraDesde) & ", "
                sql = sql & XS(sHoraHasta) & ", "
                sql = sql & XS(sOrden) & ", "
                sql = sql & XS(sMotivo) & ", "
                sql = sql & XS(mNomUser) & ", "
                sql = sql & XN(sEstado) & ", "
                sql = sql & usuarioActual & ") "
                DBConn.Execute sql
            End If
            '[FIN NUEVO]
        
            If grdGrilla.rows = 2 Then
                grdGrilla.rows = 1
            Else
                grdGrilla.RemoveItem (grdGrilla.RowSel)
            End If
        End If
        LimpiarTurno
    Else
        MsgBox "Seleccione un turno", vbExclamation, TIT_MSGBOX
    End If
End Sub

Private Sub cmdReport_Click()
    Dim Frm As New frmReporteTurnos
    
    Frm.Inicializar MViewFecha.Value, _
                    cboDoctor.ItemData(cboDoctor.ListIndex), _
                    cboDoctor.text
    Frm.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmTurnos = Nothing
        'Set rec = Nothing
        'Set Rec1 = Nothing
        'Set Rec2 = Nothing
        Unload Me
    End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdSalirP_Click()
    fraprotocolos.Visible = False
    limpiar_protocolos
    txtfiltrop.text = ""
End Sub
Private Function limpiar_protocolos()
    Dim i, j As Integer
    For i = 1 To grdProtocolos.rows - 1
        grdProtocolos.TextMatrix(i, 3) = "NO"
        For j = 0 To grdProtocolos.Cols - 1
            grdProtocolos.row = i
            grdProtocolos.Col = j
            grdProtocolos.CellForeColor = &H80000008
            grdProtocolos.CellBackColor = &H80000005
            grdProtocolos.CellFontBold = False
        Next
    Next

End Function

Private Sub cmdSoloTurnos_Click()
modoVista = 1
BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex), modoVista
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        cmdSalir_Click
    End If
End Sub
Public Sub CargarLlavesUsuarios()

    Dim obj As ClsLLave
    Dim inactiva As String

    Set dictLlaves = CreateObject("Scripting.Dictionary")

    sql = "SELECT * FROM LLAVE_USUARIO"
    Rec1.Open sql, DBConn

    Do While Not Rec1.EOF

        Set obj = New ClsLLave

        obj.Codigo = Rec1!LLA_CODIGO
        obj.Valor = Trim(CStr(Rec1!LLA_VALOR))
        obj.USUARIO = Rec1!LLA_USUARIO
        obj.activa = Rec1!activa

        dictLlaves.Add obj.Codigo, obj

        Rec1.MoveNext
    Loop

    Rec1.Close

End Sub

Public Sub GetStudiesLoadedByDate()

    Dim request As Object
    Dim responseText As String
    Dim linkDrive As String
    Dim jsonBodyToSend As String
    Dim endpoint As String
    Dim jsonBody As String
    
    endpoint = "/api/v1/studies?date=" & Format(MViewFecha.Value, "yyyy-mm-dd")
    
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    'Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    'Set request = CreateObject("MSXML2.XMLHTTP")
    
    ' Forzar TLS 1.2
    'request.Option(6) = 268435456 ' WINHTTP_OPTION_SECURE_PROTOCOLS = TLS 1.2
    'request.Option(9) = &H80000000 ' WINHTTP_OPTION_ENABLE_TLS12

    'request.Option(9) = 2048 ' TLS 1.2
    'request.Option(6) = True ' Deshabilita la caché
    'request.Option(9) = 0    ' Modo síncrono
    'request.Option(4) = 13056 ' Ignorar errores SSL

    ' ?? Fuerza el uso de TLS 1.2
    'request.Option(2) = 268435456
    'request.SetOption 2, 268435456  ' MSXML Option(2) = SecureProtocols, 268435456 = TLS 1.2

    request.Open "GET", DIGOR_CORE_URL & endpoint, False    'populates object fields
    request.setRequestHeader "Authorization", "Bearer " & DIGOR_PUBLIC_API_KEY
    request.setRequestHeader "Content-Type", "application/json"

    request.send
    responseText = request.responseText
    
    Set request = Nothing
    
    ActualizarInfoEstudiosTurnos responseText

End Sub
Private Function parseStudiesJSON(JsonString As String) As Variant
    Dim jsonObject As Object
    Dim dataArray As Object
    Dim i As Integer
    Dim dni As String
    Dim name As String
    Dim link As String
    Dim studiesArray As Variant
    Dim fullName As String
    Dim nameParts() As String
    
    ' Parseamos el JSON
    Set jsonObject = JsonConverter.ParseJson(JsonString)
    
    ' Verificamos que el campo "studies" exista
    If Not jsonObject.Exists("studies") Then
        MsgBox "Error: No se encontró el campo 'studies' en la respuesta JSON.", vbCritical
        Exit Function
    End If
    
    ' Convertimos en colección
    Set dataArray = jsonObject("studies")
    
    ' Verificar si realmente es una colección indexada
    If Not IsArray(dataArray) And Not TypeName(dataArray) = "Collection" Then
       ' MsgBox "Error: El campo 'studies' no es una colección indexada.", vbCritical
        Exit Function
    End If

    ' Verificar que el array no esté vacío
    If dataArray.Count = 0 Then
        'MsgBox "Advertencia: No hay estudios en la respuesta JSON.", vbExclamation
        Exit Function
    End If

    ' Redimensionamos el array
    ReDim studiesArray(dataArray.Count - 1, 2)

    ' Llenamos el array
    For i = 1 To dataArray.Count ' OJO: Si la colección empieza en 1, ajustamos el índice
        If Not dataArray(i).Exists("patientDNI") Then
            MsgBox "Error: El objeto en la posición " & i & " no tiene 'patientDNI'.", vbCritical
            Exit Function
        End If
        
        dni = dataArray(i)("patientDNI")
        
        fullName = dataArray(i)("name")
        nameParts = Split(fullName, " ") ' Divide la cadena por espacios
        name = nameParts(0) ' solo el prefijo, sin la fecha
        
        link = dataArray(i)("link")
        
        studiesArray(i - 1, 0) = dni
        studiesArray(i - 1, 1) = name
        studiesArray(i - 1, 2) = link
    Next i
    
    parseStudiesJSON = studiesArray
End Function
Private Sub buildStudiesDict(studiesArray As Variant)
    ' Verificar si studiesArray está vacío
    If IsEmpty(studiesArray) Then
        Set studiesDict = CreateObject("Scripting.Dictionary")
        Exit Sub
    End If

    Dim i As Integer
    Dim dni As String
    Dim studyName As String
    Dim studyLink As String

    ' Creamos un nuevo Dictionary
    Set studiesDict = CreateObject("Scripting.Dictionary")

    ' Iteramos sobre el JSON
    For i = LBound(studiesArray) To UBound(studiesArray)
        dni = studiesArray(i, 0)
        studyName = studiesArray(i, 1)
        studyLink = studiesArray(i, 2)

        ' Si el DNI ya existe, agregamos el estudio a su colección
        If studiesDict.Exists(dni) Then
            studiesDict(dni).Add studiesDict(dni).Count, Array(studyName, studyLink)
        Else
            ' Si no existe, creamos una nueva Collection y la agregamos al diccionario
            Dim newCollection As Object
            Set newCollection = CreateObject("Scripting.Dictionary")

            newCollection.Add newCollection.Count, Array(studyName, studyLink)
            studiesDict.Add dni, newCollection
        End If
    Next i
End Sub



Private Sub ActualizarInfoEstudiosTurnos(JsonString As String)
    
    ' Simulación de la respuesta del servidor con estudios
    Dim studiesArray As Variant
    Dim jsonObject As Object
    Dim success As String
    
    'Validamos respuesta exitosa del servidor
    Set jsonObject = JsonConverter.ParseJson(JsonString)
    
    success = jsonObject("success")
    
    If success <> "Verdadero" Then
        Exit Sub
    End If
    
    studiesArray = parseStudiesJSON(JsonString)
    
    'Almacenar estudios en dicc por DNI
    buildStudiesDict studiesArray
    
    ' Ahora recorremos la grilla y marcamos los turnos que tienen estudios
    Dim j As Integer
    For j = 1 To grdGrilla.rows - 1
        Dim turnoDNI As String
        turnoDNI = grdGrilla.TextMatrix(j, 11)
        If studiesDict.Exists(turnoDNI) Then
            grdGrilla.TextMatrix(j, 18) = "Ver"
        Else
            If turnoDNI = "" Then
                grdGrilla.TextMatrix(j, 18) = "" 'ESPACIO LIBRE
            Else
                grdGrilla.TextMatrix(j, 18) = "No"
            End If
            
        End If
    Next j
End Sub

' Evento de la grilla cuando el usuario hace clic en una celda
Private Sub grdGrilla_Click()
    Dim Fila As Integer
    Dim dni As String
    Dim estudios As Variant
    Dim estudio As Variant
    Dim i As Integer

    Fila = grdGrilla.row ' Obtiene la fila seleccionada

    ' Verifica si hizo clic en la columna de Estudios
    If grdGrilla.Col = 18 Then
        If grdGrilla.text = "Ver" Then
            dni = grdGrilla.TextMatrix(Fila, 11)
            
            If studiesDict.Exists(dni) Then
                Set estudios = studiesDict(dni) ' Ahora estudios es un Dictionary
            
                ' Limpiamos la lista antes de agregar nuevos elementos
                listEstudios.Clear
                
                Set estudiosUrls = CreateObject("Scripting.Dictionary")
            
                Dim key As Variant
                For Each key In estudios.keys
                    listEstudios.AddItem estudios(key)(0) ' Nombre del estudio
                    estudiosUrls(listEstudios.NewIndex) = estudios(key)(1) ' Guardar la URL con el índice
                Next key
            
                ' Mostramos el frame con la lista de estudios
                fraListaEstudios.Visible = True
            End If

        End If
    End If
End Sub

' Evento cuando se hace clic en un estudio en la lista
Private Sub listEstudios_Click()
    irAEstudio
End Sub
Private Sub irAEstudio()
    Dim link As String
    
    If listEstudios.ListIndex <> -1 Then
        ' Obtiene el link del estudio seleccionado
        link = estudiosUrls(listEstudios.ListIndex)
        If link <> "" Then
            ' Abre la URL en el navegador
            Shell "explorer " & link, vbNormalFocus
        End If
    End If
End Sub
Private Sub bcmdCerrarFraListaEstudios_Click()
    fraListaEstudios.Visible = False
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    modoVista = 1
    
    Call Centrar_pantalla(Me)
    
    ConfigurarCboTamanioSlot
    
    CargarCboDiaSemana
    
    MViewFecha.Value = Date
    fechaturno.Value = Date
    configurodia Date
    configurogrilla
    LlenarComboDoctor
    LlenarComboHoras
    ActivoGrid = 1
    If mNomUser = "DIGOR" Or mNomUser = "SILVANA" Or mNomUser = "JAVIER" Then
        lblimporte.Visible = True
        txtimporte.Visible = True
        lbltotal.Visible = True
        txtTotal.Visible = True
        cmdOcultar.Enabled = True
        cmdInforTurno.Enabled = True
    Else
        lblimporte.Visible = False
        txtimporte.Visible = False
        lbltotal.Visible = False
        txtTotal.Visible = False
        cmdOcultar.Enabled = False
    End If
    
    If mNomUser <> "DIGOR" Then
        ocultarControlesHorarios
        DeshabilitarTurnero
        cmdReport.Enabled = False
        cmdExcel.Enabled = False
        cmdExportarTurno.Enabled = False
    End If
    
    cargo_protocolos
    
    CargarLlavesUsuarios
    
    fraListaEstudios.Visible = False
    colocarModoCreacionTurno
    
    ConfigurarGrillaHorariosExc
End Sub
Private Sub LimpiarGrilla()
    grdGrilla.rows = 1
End Sub
Private Function cargo_protocolos()
    
    sql = "SELECT * FROM TIPO_IMAGEN"
    If txtfiltrop.text <> "" Then
        sql = sql & " WHERE TIP_NOMBRE LIKE '%" & txtfiltrop.text & "%'"
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
Private Function obtenerTelefonoGrila(celular As String, telefono As String)
    Dim res As String

    res = ""
    If telefono <> "" And celular <> "" Then
        res = celular & " - " & telefono
    End If
    
    If telefono <> "" And celular = "" Then
        res = telefono
    End If
    
    If telefono = "" And celular <> "" Then
        res = celular
    End If
    obtenerTelefonoGrila = res
    
End Function
Private Function obtenerCelular(celular As String)
    Dim res As String

    res = ""
    If celular <> "" Then
        res = "+54" & celular
    End If
    obtenerCelular = res
    
End Function
Private Function obtenerTieneLinkDrive(link As String)
   Dim res As String
   res = "No"
    If link <> "" Then
        res = "Si"
    End If
    
    obtenerTieneLinkDrive = res
End Function
Public Function DoctorPuedeAtender(Fecha As Date, vencod As Integer, _
                                   horaDesde As String, horaHasta As String) As Boolean
    Dim bloques() As String
    Dim rango() As String
    Dim i As Integer
    Dim bloquesStr As String

    DoctorPuedeAtender = True

    ' Obtener bloques efectivos (excepción > habitual)
    bloquesStr = ObtenerBloquesDelDia(vencod, Fecha)

    ' Sin bloques configurados ? permitir (no restringir)
    If bloquesStr = "" Then
        ' Verificar si es porque no atiende o porque no tiene config
        If dicHorariosDoc Is Nothing Then Exit Function
        If dicHorariosDoc.Count = 0 Then Exit Function
        ' Tiene config pero no atiende este día
        DoctorPuedeAtender = False
        Exit Function
    End If

    ' Recorrer bloques: el turno debe caber dentro de al menos uno
    bloques = Split(bloquesStr, ";")
    For i = 0 To UBound(bloques)
        rango = Split(bloques(i), "-")
        If horaDesde >= rango(0) And horaHasta <= rango(1) Then
            DoctorPuedeAtender = True
            Exit Function
        End If
    Next i

    DoctorPuedeAtender = False
End Function
Private Sub ocultarControlesHorarios()
    cboTamanioSlot.Visible = False
     lblTamanioSlot.Visible = False
     cmdHorariosAtencion.Visible = False
     frmVista.Visible = False
End Sub
Private Sub DeshabilitarTurnero()
     txtBuscaCliente.Enabled = False
     txtCodigo.Enabled = False
     txtBuscarCliDescri.Enabled = False
     txtTelefono.Enabled = False
     txtOSocial.Enabled = False
     txtMotivo.Enabled = False
     cboDesde.Enabled = False
     cbohasta.Enabled = False
     mebHoraD.Enabled = False
     mebHoraH.Enabled = False
     txtimporte.Enabled = False
     txtMotivo.Enabled = False
     
     cmdAgregar.Enabled = False
     cmdQuitar.Enabled = False
     cmdOcultar.Enabled = False
     cmdNuevoPaciente.Enabled = False
     cmdEditar.Enabled = False
     cmdCopiar.Enabled = False
     cboMotivo.Enabled = False

End Sub
Private Sub HabilitarTurnero()
   ' txtBuscaCliente.text = ""
    txtBuscaCliente.Enabled = True
   ' txtCodigo.text = ""
    txtCodigo.Enabled = True
   ' txtBuscarCliDescri.text = ""
    txtBuscarCliDescri.Enabled = True
   ' txtTelefono.text = ""
    txtTelefono.Enabled = True
   ' txtOSocial.text = ""
    txtOSocial.Enabled = True
   ' txtMotivo.text = ""
    txtMotivo.Enabled = True
    'cboDesde.ListIndex = -1
    cboDesde.Enabled = True
    'cbohasta.ListIndex = -1
    cbohasta.Enabled = True
    'mebHoraD.text = "__:__"
    mebHoraD.Enabled = True
    'mebHoraH.text = "__:__"
    mebHoraH.Enabled = True
    'txtimporte.text = "0,00"
    
    txtimporte.Enabled = True
    cmdAgregar.Enabled = True
    cmdQuitar.Enabled = True
    cmdOcultar.Enabled = True
    cmdNuevoPaciente.Enabled = True
     cmdEditar.Enabled = True
     cmdCopiar.Enabled = True
     cboMotivo.Enabled = True
End Sub
' ------------------------------------------------------------------
'  5. BUSCAR TURNOS (versión con slots solo para huecos libres)
' ------------------------------------------------------------------
Private Sub BuscarTurnos(Fecha As Date, Doc As Integer, Optional Modo As Integer) 'modo 1 sin lots, 2 con slots, 0 sin seleccion
    Dim foreColor As Long
    Dim backColor As Long
    Dim total As Double
    Dim años As Integer
    Dim edad As Integer
    Dim impreso As String
    
    ' --- Variables de slot ---
    Dim tamanioSlot As Integer
    Dim minInicio As Integer
    Dim minFin As Integer
    Dim currentMin As Integer
    Dim slotEndMin As Integer
    
    ' --- Variables de disponibilidad ---
    Dim DIA As Integer
    Dim diaNombre As String
    Dim claveDia As String
    Dim bloquesStr As String
    Dim tieneBloques As Boolean
    Dim bloques() As String
    Dim rango() As String
    
    ' --- Bloques parseados a arrays de minutos ---
    Dim blkStart() As Integer
    Dim blkEnd() As Integer
    Dim blkCount As Integer
    
    ' --- Arrays para turnos en memoria ---
    Dim turHoraD() As String
    Dim turHoraH() As String
    Dim turAsistio() As Integer
    Dim turDatosStr() As String
    Dim turImporte() As Double
    Dim turCount As Integer
    
    ' --- Variables auxiliares ---
    Dim idx As Integer
    Dim B As Integer
    Dim k As Integer
    Dim turnoEncontrado As Boolean
    Dim turnoIdx As Integer
    Dim inBlock As Boolean
    Dim currentBlock As Integer
    Dim nextBlockStart As Integer
    Dim turStartMin As Integer
    Dim turEndMin As Integer
    
    Dim turMostrado() As Boolean
    Dim menorFinTurno As Integer
    
    If Doc = 0 Then Exit Sub
    
    LimpiarGrilla
    
    ' =============================================================
    '  DÍA Y VALIDACIÓN DE DISPONIBILIDAD
    ' =============================================================
    DIA = Weekday(Fecha, vbMonday)
    diaNombre = UCase(Left(WeekdayName(DIA, False), 1)) _
              & Mid(WeekdayName(DIA, False), 2)
    
    'Coloco el día actual en la configuración de horarios
    lblDiaActual = diaNombre & " " & Day(MViewFecha.Value) & " de " & MonthName(Month(MViewFecha.Value), False) & " de " & Year(MViewFecha.Value)
    
    If Not DoctorAtiendeElDia(Fecha) Then
        lbldiaTurno.Caption = "El doctor no está disponible el día " & diaNombre
        lblExcepcionMotivo.Caption = ObtenerMotivoExcepcion(Doc, Fecha)
        DeshabilitarTurnero
        Exit Sub
    Else
        lblExcepcionMotivo.Caption = ObtenerMotivoExcepcion(Doc, Fecha)
        lbldiaTurno.Caption = "Turnos del dia " & diaNombre & " " & Day(MViewFecha.Value) & " de " & MonthName(Month(MViewFecha.Value), False) & " de " & Year(MViewFecha.Value)
    End If
    
    If mNomUser = "DIGOR" Then
        HabilitarTurnero
    End If

    
    ' =============================================================
    '  TAMAÑO DE SLOT
    ' =============================================================
    If cboTamanioSlot.ListIndex = -1 Then
        tamanioSlot = 30
    Else
        tamanioSlot = cboTamanioSlot.ItemData(cboTamanioSlot.ListIndex)
    End If
    
    
    ' =============================================================
    '  OBTENER BLOQUES DE ATENCIÓN DEL DÍA
    ' =============================================================
    'claveDia = CStr(DIA)
    'tieneBloques = False
    
   ' If Not dicHorariosDoc Is Nothing Then
   '     If dicHorariosDoc.Count > 0 Then
   '         If dicHorariosDoc.Exists(claveDia) Then
   '             bloquesStr = dicHorariosDoc(claveDia)
   '             tieneBloques = True
   '         End If
   '     End If
   ' End If
    bloquesStr = ObtenerBloquesDelDia(Doc, Fecha)
    tieneBloques = (bloquesStr <> "")
    
    
    ' =============================================================
    '  CONSULTA DE TURNOS EXISTENTES
    ' =============================================================
    sql = "SELECT TOP 100 T.*,V.VEN_NOMBRE,C.CLI_RAZSOC,C.CLI_NRODOC," _
        & "C.CLI_TELEFONO,C.CLI_CELULAR,C.CLI_CUMPLE, C.CLI_LINKARCH, T.ID_TURNO" _
        & " FROM TURNOS T, VENDEDOR V, CLIENTE C" _
        & " WHERE T.CLI_CODIGO = C.CLI_CODIGO" _
        & " AND T.VEN_CODIGO = V.VEN_CODIGO" _
        & " AND T.TUR_FECHA = " & XDQ(Fecha) _
        & " AND T.VEN_CODIGO = " & Doc _
        & " AND T.DELETED_AT IS NULL" _
        & " ORDER BY T.TUR_HORAD"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    
    ' =============================================================
    '  MODO CON SLOTS (doctor tiene horarios configurados)
    ' =============================================================
    If (tieneBloques And mNomUser = "DIGOR") And Modo <> 1 Then
        
        ' ---------------------------------------------------------
        '  PASO 1: Cargar turnos existentes en arrays
        ' ---------------------------------------------------------
        turCount = 0
        
        If rec.EOF = False Then
            turCount = rec.RecordCount
            
            ReDim turHoraD(1 To turCount)
            ReDim turHoraH(1 To turCount)
            ReDim turAsistio(1 To turCount)
            ReDim turDatosStr(1 To turCount)
            ReDim turImporte(1 To turCount)
            
            idx = 1
            Do While rec.EOF = False
                turHoraD(idx) = Format(rec!TUR_HORAD, "hh:mm")
                turHoraH(idx) = Format(rec!TUR_HORAH, "hh:mm")
                turAsistio(idx) = rec!TUR_ASISTIO
                turImporte(idx) = Chk0(rec!TUR_IMPORTE)
                
                ' Calcular edad
                If Not (IsNull(rec!CLI_CUMPLE)) Then
                    años = Year(Date) - Year(rec!CLI_CUMPLE)
                    If Month(Fecha) < Month(rec!CLI_CUMPLE) Then años = años - 1
                    If Month(Now) = Month(rec!CLI_CUMPLE) And _
                       Day(Fecha) < Day(rec!CLI_CUMPLE) Then años = años - 1
                    edad = años
                Else
                    edad = 0
                End If
                
                If Chk0(rec!TUR_IMPRESO) = 1 Then
                    impreso = "SI"
                Else
                    impreso = "NO"
                End If
                
                ' Cols 1 a 20 (SIN col 0, la hora se pone al agregar)
                turDatosStr(idx) = _
                    rec!CLI_RAZSOC & Chr(9) & _
                    edad & Chr(9) & _
                    obtenerTelefonoGrila(ChkNull(rec!CLI_CELULAR), _
                                         ChkNull(rec!CLI_TELEFONO)) & Chr(9) & _
                    " " & Chr(9) & _
                    rec!TUR_OSOCIAL & Chr(9) & _
                    ChkNull(rec!TUR_MOTIVO) & Chr(9) & _
                    ChkNull(rec!TUR_DRSOLICITA) & Chr(9) & _
                    rec!VEN_CODIGO & Chr(9) & _
                    rec!CLI_CODIGO & Chr(9) & _
                    rec!TUR_ASISTIO & Chr(9) & _
                    ChkNull(rec!CLI_NRODOC) & Chr(9) & _
                    ChkNull(rec!TUR_DESDE) & Chr(9) & _
                    rec!TUR_TIENEMUTUAL & Chr(9) & _
                    Format(Chk0(rec!TUR_IMPORTE), "#,##0.00") & Chr(9) & _
                    ChkNull(rec!tur_orden) & Chr(9) & _
                    impreso & Chr(9) & _
                    obtenerTieneLinkDrive(ChkNull(rec!CLI_LINKARCH)) & Chr(9) & _
                    "" & Chr(9) & _
                    ChkNull(rec!TUR_OBSERV) & Chr(9) & _
                    ChkNull(rec!ID_TURNO)
                
                idx = idx + 1
                rec.MoveNext
            Loop
        End If
        ' Marcar todos como no mostrados
        If turCount > 0 Then
            ReDim turMostrado(1 To turCount)
            For k = 1 To turCount
                turMostrado(k) = False
            Next k
        End If
        
        
        ' ---------------------------------------------------------
        '  PASO 2: Parsear bloques a arrays de minutos
        '          y calcular rango total
        ' ---------------------------------------------------------
        bloques = Split(bloquesStr, ";")
        blkCount = UBound(bloques) + 1
        ReDim blkStart(0 To blkCount - 1)
        ReDim blkEnd(0 To blkCount - 1)
        
        minInicio = 1440
        minFin = 0
        
        For B = 0 To blkCount - 1
            rango = Split(bloques(B), "-")
            blkStart(B) = HoraAMinutos(rango(0))
            blkEnd(B) = HoraAMinutos(rango(1))
            If blkStart(B) < minInicio Then minInicio = blkStart(B)
            If blkEnd(B) > minFin Then minFin = blkEnd(B)
        Next B
        
        
        ' ---------------------------------------------------------
        '  PASO 3: Recorrer línea de tiempo y armar la grilla
        '
        '  En cada posición:
        '    ¿Hay un turno que cubre este momento?
        '       ? UNA fila con el rango real del turno, avanzar al fin
        '    ¿Estoy dentro de un bloque de atención sin turno?
        '       ? UNA fila "Disponible" del tamaño del slot
        '    ¿Estoy fuera de todo bloque?
        '       ? UNA fila "Fuera de horario" hasta el próximo bloque
        ' ---------------------------------------------------------
        ' ---------------------------------------------------------
        '  PASO 3: Recorrer línea de tiempo y armar la grilla
        ' ---------------------------------------------------------
        grdGrilla.rows = 1
        i = 1
        total = 0
        currentMin = minInicio
        
        Do While currentMin < minFin
            
            ' --- Buscar TODOS los turnos que cubren este momento ---
            menorFinTurno = minFin
            turnoEncontrado = False
            
            For k = 1 To turCount
                If Not turMostrado(k) Then
                    turStartMin = HoraAMinutos(turHoraD(k))
                    turEndMin = HoraAMinutos(turHoraH(k))
                    
                    If turStartMin <= currentMin And turEndMin > currentMin Then
                        turnoEncontrado = True
                        turMostrado(k) = True
                        
                        ' Agregar fila del turno
                        grdGrilla.AddItem turHoraD(k) & " a " _
                            & turHoraH(k) & Chr(9) & turDatosStr(k)
                        total = total + turImporte(k)
                        
                        ' Color según asistencia
                        Select Case turAsistio(k)
                            Case 0
                                backColor = &H800080: foreColor = &HFFFFFF
                            Case 1
                                backColor = &H8000&:  foreColor = &HFFFFFF
                            Case 2
                                backColor = &HC0C0&:  foreColor = &H80000008
                        End Select
                        
                        ' Colorear la fila
                        grdGrilla.row = i
                        grdGrilla.Col = 0
                        grdGrilla.CellForeColor = &HFFFFFF
                        grdGrilla.CellBackColor = &H808080
                        grdGrilla.CellFontBold = True
                        For j = 1 To grdGrilla.Cols - 1
                            grdGrilla.Col = j
                            grdGrilla.CellForeColor = foreColor
                            grdGrilla.CellBackColor = backColor
                            grdGrilla.CellFontBold = True
                        Next j
                        
                        i = i + 1
                        
                        ' Rastrear el menor fin para saber hasta dónde avanzar
                        If turEndMin < menorFinTurno Then menorFinTurno = turEndMin
                    End If
                End If
            Next k
            
            
            If turnoEncontrado Then
                ' Antes de avanzar, verificar si hay turnos no mostrados
                ' que arrancan entre currentMin y menorFinTurno
                ' (turnos solapados que empiezan después)
                Dim nextStart As Integer
                nextStart = menorFinTurno
                For k = 1 To turCount
                    If Not turMostrado(k) Then
                        turStartMin = HoraAMinutos(turHoraD(k))
                        If turStartMin > currentMin And turStartMin < nextStart Then
                            nextStart = turStartMin
                        End If
                    End If
                Next k
                currentMin = nextStart
            
            Else
                ' --- No hay turnos: ¿estoy en bloque de atención? ---
                inBlock = False
                For B = 0 To blkCount - 1
                    If currentMin >= blkStart(B) And currentMin < blkEnd(B) Then
                        inBlock = True
                        currentBlock = B
                        Exit For
                    End If
                Next B
                
                If inBlock Then
                    ' ===========================
                    '  DISPONIBLE (1 slot)
                    ' ===========================
                    slotEndMin = currentMin + tamanioSlot
                    
                    If slotEndMin > blkEnd(currentBlock) Then
                        slotEndMin = blkEnd(currentBlock)
                    End If
                    
                    ' No pisar el inicio del próximo turno no mostrado
                    For k = 1 To turCount
                        If Not turMostrado(k) Then
                            turStartMin = HoraAMinutos(turHoraD(k))
                            If turStartMin > currentMin And turStartMin < slotEndMin Then
                                slotEndMin = turStartMin
                            End If
                        End If
                    Next k
                    
                    grdGrilla.AddItem ArmarFilaSlotVacia( _
                        MinutosAHora(currentMin), _
                        MinutosAHora(slotEndMin), "Disponible")
                    
                    backColor = COLOR_DISPONIBLE_BACK
                    foreColor = COLOR_DISPONIBLE_FORE
                    
                    ' Colorear la fila
                    grdGrilla.row = i
                    grdGrilla.Col = 0
                    grdGrilla.CellForeColor = &HFFFFFF
                    grdGrilla.CellBackColor = &H808080
                    grdGrilla.CellFontBold = True
                    For j = 1 To grdGrilla.Cols - 1
                        grdGrilla.Col = j
                        grdGrilla.CellForeColor = foreColor
                        grdGrilla.CellBackColor = backColor
                        grdGrilla.CellFontBold = True
                    Next j
                    
                    i = i + 1
                    currentMin = slotEndMin
                
                Else
                    ' ===========================
                    '  FUERA DE HORARIO (1 fila)
                    ' ===========================
                    nextBlockStart = minFin
                    For B = 0 To blkCount - 1
                        If blkStart(B) > currentMin And blkStart(B) < nextBlockStart Then
                            nextBlockStart = blkStart(B)
                        End If
                    Next B
                    
                    grdGrilla.AddItem ArmarFilaSlotVacia( _
                        MinutosAHora(currentMin), _
                        MinutosAHora(nextBlockStart), "Fuera de horario")
                    
                    backColor = COLOR_FUERA_BACK
                    foreColor = COLOR_FUERA_FORE
                    
                    ' Colorear la fila
                    grdGrilla.row = i
                    grdGrilla.Col = 0
                    grdGrilla.CellForeColor = &HFFFFFF
                    grdGrilla.CellBackColor = &H808080
                    grdGrilla.CellFontBold = True
                    For j = 1 To grdGrilla.Cols - 1
                        grdGrilla.Col = j
                        grdGrilla.CellForeColor = foreColor
                        grdGrilla.CellBackColor = backColor
                        grdGrilla.CellFontBold = True
                    Next j
                    
                    i = i + 1
                    currentMin = nextBlockStart
                End If
            End If
        Loop
    
    
    ' =============================================================
    '  MODO SIN SLOTS (fallback: doctor sin horarios configurados)
    ' =============================================================
    Else
        grdGrilla.rows = 1
        If rec.EOF = False Then
            i = 1
            Do While rec.EOF = False
                Select Case rec!TUR_ASISTIO
                    Case 0
                        backColor = &H800080
                        foreColor = &HFFFFFF
                    Case 1
                        backColor = &H8000&
                        foreColor = &HFFFFFF
                    Case 2
                        backColor = &HC0C0&
                        foreColor = &H80000008
                End Select
                
                If Not (IsNull(rec!CLI_CUMPLE)) Then
                    If rec.EOF = False Then
                        años = Year(Date) - Year(rec!CLI_CUMPLE)
                        If Month(Fecha) < Month(rec!CLI_CUMPLE) Then años = años - 1
                        If Month(Now) = Month(rec!CLI_CUMPLE) And _
                           Day(Fecha) < Day(rec!CLI_CUMPLE) Then años = años - 1
                        edad = años
                    End If
                Else
                    edad = 0
                End If
                
                If Chk0(rec!TUR_IMPRESO) = 1 Then
                    impreso = "SI"
                Else
                    impreso = "NO"
                End If
                
                grdGrilla.AddItem Format(rec!TUR_HORAD, "hh:mm") & " a " & _
                    Format(rec!TUR_HORAH, "hh:mm") & Chr(9) & _
                    rec!CLI_RAZSOC & Chr(9) & edad & Chr(9) & _
                    obtenerTelefonoGrila(ChkNull(rec!CLI_CELULAR), _
                        ChkNull(rec!CLI_TELEFONO)) & Chr(9) & _
                    " " & Chr(9) & rec!TUR_OSOCIAL & Chr(9) & _
                    ChkNull(rec!TUR_MOTIVO) & Chr(9) & _
                    ChkNull(rec!TUR_DRSOLICITA) & Chr(9) & _
                    rec!VEN_CODIGO & Chr(9) & rec!CLI_CODIGO & Chr(9) & _
                    rec!TUR_ASISTIO & Chr(9) & ChkNull(rec!CLI_NRODOC) & Chr(9) & _
                    ChkNull(rec!TUR_DESDE) & Chr(9) & rec!TUR_TIENEMUTUAL & Chr(9) & _
                    Format(Chk0(rec!TUR_IMPORTE), "#,##0.00") & Chr(9) & _
                    ChkNull(rec!tur_orden) & Chr(9) & impreso & Chr(9) & _
                    obtenerTieneLinkDrive(ChkNull(rec!CLI_LINKARCH)) & Chr(9) & _
                    "" & Chr(9) & ChkNull(rec!TUR_OBSERV) & Chr(9) & _
                    ChkNull(rec!ID_TURNO)
                
                total = total + Chk0(rec!TUR_IMPORTE)
                
                grdGrilla.Col = 0
                grdGrilla.row = i
                grdGrilla.CellForeColor = &HFFFFFF
                grdGrilla.CellBackColor = &H808080
                grdGrilla.CellFontBold = True
                
                grdGrilla.row = i
                For j = 1 To grdGrilla.Cols - 1
                    grdGrilla.Col = j
                    grdGrilla.CellForeColor = foreColor
                    grdGrilla.CellBackColor = backColor
                    grdGrilla.CellFontBold = True
                Next j
                
                i = i + 1
                rec.MoveNext
            Loop
        End If
    End If
    
    
    ' =============================================================
    '  CIERRE COMÚN
    ' =============================================================
    txtTotal.text = total
    txtTotal.text = Valido_Importe(txtTotal.text)
    
    rec.Close
    grdGrilla.Col = 10
    If grdGrilla.row > 1 Then
        grdGrilla.row = 1
    End If
    GetStudiesLoadedByDate
    
End Sub
Private Function cambiocolor(asistio As Integer)
    Dim foreColor As String
    Dim backColor As String
    
    Select Case asistio
    Case 0
        backColor = &H800080
        foreColor = &HFFFFFF
    Case 1
        backColor = &H8000&
        foreColor = &HFFFFFF
    Case 2
        backColor = &HC0C0&
        foreColor = &H80000008
    End Select
    
    grdGrilla.row = grdGrilla.RowSel
    For j = 1 To grdGrilla.Cols - 1
        grdGrilla.Col = j
        grdGrilla.CellForeColor = foreColor       'FUENTE COLOR NEGRO
        grdGrilla.CellBackColor = backColor      'ROSA
        grdGrilla.CellFontBold = True
    Next
    
End Function


Private Sub LlenarComboDoctor()
    LimpiarComboDoctor
     'BUSCO CODIGO DE DOCTOR POR NOMBRE DE USUARIO
    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO > 1 and VEN_ESTADO = 'N' "
    If mNomUser <> "A" And mNomUser <> "DIGOR" Then
        sql = sql & " AND VEN_NOMBRE LIKE '%" & mNomUser & "%'"
    End If
    sql = sql & " order by VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        'cboFactura1.AddItem "(Todas)"
        cboDoctor.AddItem ""
        Do While rec.EOF = False
            cboDoctor.AddItem rec!VEN_NOMBRE
            cboDoctor.ItemData(cboDoctor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        'coloco el close aca xq lo va a usar el metodo de buscarturnos que se aciva al hacer: cboDoctor.ListIndex = 0
        rec.Close
        cboDoctor.ListIndex = 0
    End If
    'rec.Close
End Sub
Private Sub LlenarComboHoras()
    Dim cItems As Integer
    Dim cont As Integer
    Dim minutos As Integer
    Dim z As Integer
    rec.Open "SELECT HS_DESDE,HS_HASTA FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        hDesde = Hour(rec!HS_DESDE)
        hHasta = Hour(rec!HS_HASTA)
    End If
    rec.Close
    cItems = (hHasta - hDesde) * 12 + 1
    i = 0
    
    cont = 1
    j = hDesde
    Do While cont < cItems
        minutos = 0
        For z = 0 To 11
            If cont < cItems Then
                If (minutos + 5) > 60 Then
                    'cboDesde.AddItem Format(J, "00") & ":" & Format(minutos, "00") & " a " & Format(J + 1, "00") & ":" & Format(0, "00")
                    Exit For
                Else
                    cboDesde.AddItem Format(j, "00") & ":" & Format(minutos, "00")
                    cboDesde.ItemData(cboDesde.NewIndex) = cont
                    cbohasta.AddItem Format(j, "00") & ":" & Format(minutos, "00")
                    cbohasta.ItemData(cbohasta.NewIndex) = cont
                End If
            End If
            cont = cont + 1
            minutos = minutos + 5
        Next
        j = j + 1
    Loop
    cbohasta.AddItem Format(hHasta, "00") & ":" & Format(0, "00")

    cboDesde.ListIndex = -1
    cbohasta.ListIndex = -1
    
End Sub
Private Function configurogrilla()
    Dim z As Integer
    Dim minutos As Integer
    Dim minutos_sig As Integer
    Dim cont As Integer
    grdGrilla.FormatString = "^Horas|<Paciente|<Edad|<Telefono|<Celular|<Obra Social|<Motivo|Dr Solicitante|>Doctor|>Cod Pac|>Asistio|DNI|TUR_DESDE|TieneMutual|Importe|Orden|Impreso|Drive|Estudios|Observaciones|idTurno"
    grdGrilla.ColWidth(0) = 1200 'HORAS
    grdGrilla.ColWidth(1) = 2300 'PACIENTE
    grdGrilla.ColWidth(2) = 500 'EDAD
    grdGrilla.ColWidth(3) = 1700 'CELULAR/TELEFONO
    grdGrilla.ColWidth(4) = 0 'CELULAR
    grdGrilla.ColWidth(5) = 1700 'O SOCIAL
    grdGrilla.ColWidth(6) = 2000 'MOTIVO
    grdGrilla.ColWidth(7) = 1500 'Dr Solicitante
    grdGrilla.ColWidth(8) = 0 'DOCTORcre
    grdGrilla.ColWidth(9) = 0 'Codigo Paciente
    grdGrilla.ColWidth(10) = 0 'Asistio
    grdGrilla.ColWidth(11) = 0 'DNI
    grdGrilla.ColWidth(12) = 0 'TUR_DESDE
    grdGrilla.ColWidth(13) = 0 'TUR_TIENEMUTUAL
    'If User = 1 Then 'ESTA CONFIGURACION LA TOMA DEL INI
    If mNomUser = "DIGOR" Or mNomUser = "SILVANA" Then 'ESTA CONFIGURACION LA TOMA DEL USUARIO LOGUEADO
        grdGrilla.ColWidth(14) = 1200 'Importe
        grdGrilla.ColWidth(15) = 600 'ORDEN
        grdGrilla.ColWidth(16) = 0 'IMPRESO
    Else
        'oculto la columna de importe para los doctores
        grdGrilla.ColWidth(14) = 0 'Importe
        grdGrilla.ColWidth(15) = 500 'ORDEN
        grdGrilla.ColWidth(16) = 0 'IMPRESO
    End If
    grdGrilla.ColWidth(17) = 550 'TIENE LINK DRIVE
    grdGrilla.ColWidth(18) = 800 'TIENE ESTUDIOS CARGADOS
    grdGrilla.ColWidth(19) = 2050 'OBSERVACIONES
    
    'Campos auditoria
    grdGrilla.ColWidth(20) = 0 'ID TURNO
    
    grdGrilla.Cols = 21
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To grdGrilla.Cols - 1
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    ' Busco los horarios en parametros
    rec.Open "SELECT HS_DESDE,HS_HASTA FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        hDesde = Hour(rec!HS_DESDE)
        hHasta = Hour(rec!HS_HASTA)
    End If
    rec.Close
    grdGrilla.rows = (hHasta - hDesde) * 12 + 1
    
    For i = 1 To grdGrilla.rows - 1
        grdGrilla.Col = 0
        grdGrilla.row = i
        'grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        'grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR NEGRO
        grdGrilla.CellBackColor = &H800080    'VIOLETA OSCURO
        grdGrilla.CellFontBold = True
        
    Next
    
    grdGrilla.rows = 1
    
'    J = hDesde
'    cont = 1
'    Do While cont < grdGrilla.Rows
'        minutos = 0
'        For z = 0 To 11
'            If cont < grdGrilla.Rows Then
'                If (minutos + 5) = 60 Then
'                    grdGrilla.TextMatrix(cont, 0) = Format(J, "00") & ":" & Format(minutos, "00") & " a " & Format(J + 1, "00") & ":" & Format(0, "00")
'                Else
'                    grdGrilla.TextMatrix(cont, 0) = Format(J, "00") & ":" & Format(minutos, "00") & " a " & Format(J, "00") & ":" & Format(minutos + 5, "00")
'                End If
'            End If
'            cont = cont + 1
'            minutos = minutos + 5
'        Next
'        J = J + 1
'    Loop

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
    grdProtocolos.rows = 1
    grdProtocolos.HighLight = flexHighlightAlways
    
End Function

Private Sub GRDGrilla_DblClick()
    'Sino ho hay turno, no rediriko a historia clinia
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 9) = "" Then
        Exit Sub
    End If
     'BUSCO CODIGO DE DOCTOR POR NOMBRE DE USUARIO logeado
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
    
    If Doc = cboDoctor.ItemData(cboDoctor.ListIndex) Or mNomUser = "DIGOR" Then
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = "PARTICULAR" Then
            TurOSocial = "PARTICULAR"
        Else
            TurOSocial = ""
        End If
        frmhistoriaclinica.txtCodigo = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
        frmhistoriaclinica.txthorad = fechaturno & " " & Left(grdGrilla.TextMatrix(grdGrilla.RowSel, 0), 5)
        frmhistoriaclinica.txtMotivo = grdGrilla.TextMatrix(grdGrilla.RowSel, 6)
        frmhistoriaclinica.txtDoctorSolicitante = grdGrilla.TextMatrix(grdGrilla.RowSel, 7)
        frmhistoriaclinica.Show vbModal
    End If
    

End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdQuitar_Click
    End If

End Sub

Private Sub grdProtocolos_DblClick()
    Dim j As Integer
    If grdProtocolos.TextMatrix(grdProtocolos.RowSel, 8) = "NO" Then
        grdProtocolos.TextMatrix(grdProtocolos.RowSel, 8) = "SI"
        'CAMBIAR COLOR
        'backColor = &HC000&
        'foreColor = &HFFFFFF
        For j = 0 To grdProtocolos.Cols - 1
            grdProtocolos.Col = j
            grdProtocolos.CellForeColor = &HFFFFFF
            grdProtocolos.CellBackColor = &H8000&
            grdProtocolos.CellFontBold = True
        Next
    Else
        grdProtocolos.TextMatrix(grdProtocolos.RowSel, 8) = "NO"
        For j = 0 To grdProtocolos.Cols - 1
            grdProtocolos.Col = j
            grdProtocolos.CellForeColor = &H80000008
            grdProtocolos.CellBackColor = &H80000005
            grdProtocolos.CellFontBold = False
        Next
    End If
End Sub

Private Sub grdProtocolos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        grdProtocolos_DblClick
    End If
End Sub

Private Sub listEstudios_DblClick()
    irAEstudio
End Sub

Private Sub mebHoraD_LostFocus()
    If Mid(mebHoraD.text, 1, 2) = "__" And Mid(mebHoraD.text, 4, 2) <> "__" Then
        mebHoraD.text = "00:" & Mid(mebHoraD.text, 4, 2)
    End If
    If Mid(mebHoraD.text, 4, 2) = "__" And Mid(mebHoraD.text, 1, 2) <> "__" Then
        mebHoraD.text = Mid(mebHoraD.text, 1, 2) & ":00"
    End If
End Sub

Private Sub mebHoraH_LostFocus()
    If Mid(mebHoraH.text, 1, 2) = "__" And Mid(mebHoraH.text, 4, 2) <> "__" Then
        mebHoraH.text = "00:" & Mid(mebHoraH.text, 4, 2)
    End If
    If Mid(mebHoraH.text, 4, 2) = "__" And Mid(mebHoraH.text, 1, 2) <> "__" Then
        mebHoraH.text = Mid(mebHoraH.text, 1, 2) & ":00"
    End If
End Sub

Private Sub MViewFecha_DateClick(ByVal DateClicked As Date)
    configurodia MViewFecha.Value
    fechaturno.Value = MViewFecha.Value
    LimpiarGrilla
    CargarExcepcionDia MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex), modoVista

End Sub
Private Sub configurodia(Fecha As Date)
    Dim DIA As Integer
    DIA = Weekday(Fecha, vbMonday)
    lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(DIA, False) & " " & Day(Fecha) & " de " & MonthName(Month(Fecha), False) & " de " & Year(Fecha)
End Sub
Private Function BuscarOSocial(CodCli As Long) As String
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

Private Sub Option1_Click()
    txtOSocial.Enabled = True
End Sub

Private Sub Option2_Click()
    txtOSocial.Enabled = False
End Sub

Private Sub optNO_Click()
    txtOSocial.Enabled = False
End Sub

Private Sub optSI_Click()
    txtOSocial.Enabled = True
End Sub

Private Sub txtBuscaCliente_Change()
    If txtBuscaCliente.text = "" Then
        txtBuscarCliDescri.text = ""
        txtCodigo.text = ""
        txtTelefono.text = ""
        txtOSocial.text = ""
        LimpiarTurno
    End If
    If Len(Trim(txtBuscaCliente.text)) < 7 Then
        txtBuscaCliente.ToolTipText = "Numero de Paciente"
    Else
        txtBuscaCliente.ToolTipText = "DNI"
    End If
End Sub

Private Sub txtBuscaCliente_GotFocus()
    SelecTexto txtBuscaCliente
End Sub

Private Sub txtBuscaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
        ActivoGrid = 0
    End If
End Sub

Private Sub txtBuscaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtBuscaCliente_LostFocus()
    If txtBuscaCliente.text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,OS_NUMERO,CLI_CELULAR"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.text <> "" Then
            If Len(Trim(txtBuscaCliente.text)) < 7 Then
                sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
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
            txtBuscarCliDescri.text = rec!CLI_RAZSOC
            txtCodigo.text = rec!CLI_CODIGO
            txtTelefono.text = ChkNull(rec!CLI_TELEFONO)
            txtcelular.text = ChkNull(rec!CLI_CELULAR)
            txtOSocial.text = BuscarOSocial(rec!CLI_CODIGO)
            If IsNull(rec!OS_NUMERO) Then
                optSI.Enabled = False
                optNO.Value = True
            Else
                optSI.Value = True
            End If
            'txtMotivo.SetFocus
            ActivoGrid = 1
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
            'txtBuscaCliente.SetFocus
            cmdNuevo.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtBuscarCliDescri_Change()
    If txtBuscarCliDescri.text = "" Then
        txtBuscaCliente.text = ""
        txtCodigo.text = ""
        txtTelefono.text = ""
        txtcelular.text = ""
        txtOSocial.text = ""
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
    If txtBuscaCliente.text = "" And txtBuscarCliDescri.text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO, CLI_CELULAR"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.text <> "" Then
            If Len(Trim(txtBuscaCliente.text)) < 7 Then
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
                BuscarClientes "txtBuscaCliente", "CADENA", Trim(txtBuscarCliDescri.text)
                If rec.State = 1 Then rec.Close
                txtBuscarCliDescri.SetFocus
            Else
                'txtBuscaCliente.Text = rec!CLI_DNI
                If Len(Trim(txtBuscaCliente.text)) < 7 Then
                    txtBuscaCliente.text = rec!CLI_CODIGO
                Else
                    txtBuscaCliente.text = rec!CLI_NRODOC
                End If
                'txtBuscaCliente.Text = rec!CLI_NRODOC
                txtBuscarCliDescri.text = rec!CLI_RAZSOC
                txtCodigo.text = rec!CLI_CODIGO
                txtTelefono.text = ChkNull(rec!CLI_TELEFONO)
                txtcelular.text = ChkNull(rec!CLI_CELULAR)
            End If
            ActivoGrid = 0
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
        
        hSQL = "Nombre, CÃ³digo, DNI"
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
                'txtcodCli.Text = .ResultFields(2)
                'txtCodCli_LostFocus
            Else
                If .ResultFields(3) = "" Then
                    txtBuscaCliente.text = .ResultFields(2)
                    txtCodigo.text = .ResultFields(3)
                Else
                    txtBuscaCliente.text = .ResultFields(3)
                    txtCodigo.text = .ResultFields(3)
                End If
                txtBuscaCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub


Private Sub txtDrSolicitante_GotFocus()
    seltxt
End Sub
Private Sub txtObservaciones_GotFocus()
    seltxt
End Sub

Private Sub txtfiltrop_GotFocus()
    seltxt
End Sub

Private Sub txtfiltrop_LostFocus()
    grdProtocolos.rows = 1
    cargo_protocolos
End Sub

Private Sub txtimporte_GotFocus()
    seltxt
End Sub

Private Sub txtimporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtimporte, KeyAscii)
End Sub

Private Sub txtimporte_LostFocus()
    If txtimporte.text <> "" Then
        txtimporte.text = Valido_Importe(txtimporte)
    End If
End Sub

Private Sub txtMotivo_GotFocus()
    SelecTexto txtMotivo
End Sub


Private Sub txtOrden_GotFocus()
    seltxt
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtOrden, KeyAscii)
End Sub
' ------------------------------------------------------------------
'  MostrarDisponibilidad: lee de grdHorarios (no de dicHorariosDoc)
'  Así refleja cambios sin guardar
' ------------------------------------------------------------------
Private Sub MostrarDisponibilidad()
    Dim dicTemp As Object
    Set dicTemp = CreateObject("Scripting.Dictionary")
    
    Dim r As Integer
    Dim clave As String
    Dim bloque As String
    
    ' Armar diccionario temporal desde la grilla de edición
    For r = 2 To grdHorarios.rows - 1
        If grdHorarios.TextMatrix(r, 0) <> "" Then
            clave = grdHorarios.TextMatrix(r, 0)
            bloque = grdHorarios.TextMatrix(r, 2) & "-" & grdHorarios.TextMatrix(r, 3)
            
            If dicTemp.Exists(clave) Then
                dicTemp(clave) = dicTemp(clave) & ";" & bloque
            Else
                dicTemp.Add clave, bloque
            End If
        End If
    Next r
    
    ' Configurar grilla de resumen
    With grdDisponibilidad
        .rows = 1
        .Cols = 2
        .FormatString = "Día|Horarios"
        .ColWidth(0) = 1500
        .ColWidth(1) = 5400
        .BorderStyle = flexBorderNone
        .SelectionMode = flexSelectionByRow
        .row = 0
        Dim c As Integer
        For c = 0 To 1
            .Col = c
            .CellForeColor = &HFFFFFF
            .CellBackColor = &H808080
            .CellFontBold = True
        Next c
    End With
    
    ' Llenar una fila por día
    Dim i As Integer
    Dim horarios As String
    Dim bloques() As String
    Dim rango() As String
    Dim j As Integer
    
    For i = 1 To 7
        clave = CStr(i)
        
        If dicTemp.Exists(clave) Then
            bloques = Split(dicTemp(clave), ";")
            horarios = ""
            For j = 0 To UBound(bloques)
                rango = Split(bloques(j), "-")
                If horarios <> "" Then horarios = horarios & "  /  "
                horarios = horarios & rango(0) & " a " & rango(1)
            Next j
        Else
            horarios = "No atiende"
        End If
        
        grdDisponibilidad.AddItem NombreDiaSemana(i) & Chr(9) & horarios
        
        grdDisponibilidad.row = i
        grdDisponibilidad.Col = 0
        grdDisponibilidad.CellFontBold = True
        
        grdDisponibilidad.Col = 1
        If dicTemp.Exists(clave) Then
            grdDisponibilidad.CellBackColor = &HC0FFC0
            grdDisponibilidad.CellForeColor = &H80000008
        Else
            grdDisponibilidad.CellBackColor = &HC0C0C0
            grdDisponibilidad.CellForeColor = &H808080
        End If
    Next i
    
    Set dicTemp = Nothing
End Sub

' ---------------------------------------------------------------
'  Arma la grilla con la disponibilidad de los 7 días
'  usando el diccionario ya cargado (dicHorariosDoc)
' ---------------------------------------------------------------
Private Sub MostrarDisponibilidad_old()
    Dim i As Integer
    Dim j As Integer
    Dim clave As String
    Dim horarios As String
    Dim bloques() As String
    Dim rango() As String

    With grdDisponibilidad
        .rows = 1
        .Cols = 2
        .FormatString = "Día|Horarios"
        .ColWidth(0) = 1500
        .ColWidth(1) = 5400
        .BorderStyle = flexBorderNone
        .SelectionMode = flexSelectionByRow

        ' Formato encabezado
        .row = 0
        Dim c As Integer
        For c = 0 To 1
            .Col = c
            .CellForeColor = &HFFFFFF
            .CellBackColor = &H808080
            .CellFontBold = True
        Next c
    End With

    ' Una fila por día de la semana
    For i = 1 To 7
        clave = CStr(i)

        If dicHorariosDoc Is Nothing Then
            horarios = "Sin datos"
        ElseIf dicHorariosDoc.Count = 0 Then
            horarios = "Sin datos"
        ElseIf dicHorariosDoc.Exists(clave) Then
            ' Formatear bloques: "08:00-12:00;14:00-18:00" ? "08:00 a 12:00 / 14:00 a 18:00"
            bloques = Split(dicHorariosDoc(clave), ";")
            horarios = ""
            For j = 0 To UBound(bloques)
                rango = Split(bloques(j), "-")
                If horarios <> "" Then horarios = horarios & "  /  "
                horarios = horarios & rango(0) & " a " & rango(1)
            Next j
        Else
            horarios = "No atiende"
        End If

        grdDisponibilidad.AddItem NombreDiaSemana(i) & Chr(9) & horarios

        ' Colorear según si atiende o no
        grdDisponibilidad.row = i
        grdDisponibilidad.Col = 0
        grdDisponibilidad.CellFontBold = True

        grdDisponibilidad.Col = 1
        If Not dicHorariosDoc Is Nothing Then
            If dicHorariosDoc.Exists(clave) Then
                grdDisponibilidad.CellBackColor = &HC0FFC0    ' verde claro
                grdDisponibilidad.CellForeColor = &H80000008  ' negro
            Else
                grdDisponibilidad.CellBackColor = &HC0C0C0    ' gris
                grdDisponibilidad.CellForeColor = &H808080    ' gris oscuro
            End If
        End If
    Next i

    fraConfiguracionHorarios.Visible = True
    fraConfiguracionHorarios.ZOrder 0   ' traer al frente
End Sub
' ---------------------------------------------------------------
'  Auxiliar: número de día ? nombre
' ---------------------------------------------------------------
Private Function NombreDiaSemana(numDia As Integer) As String
    Select Case numDia
        Case 1: NombreDiaSemana = "Lunes"
        Case 2: NombreDiaSemana = "Martes"
        Case 3: NombreDiaSemana = "Miércoles"
        Case 4: NombreDiaSemana = "Jueves"
        Case 5: NombreDiaSemana = "Viernes"
        Case 6: NombreDiaSemana = "Sábado"
        Case 7: NombreDiaSemana = "Domingo"
    End Select
End Function
' ---------------------------------------------------------------
'  Botón "Horarios de atención"
' ---------------------------------------------------------------
Private Sub cmdHorariosAtencion_Click()
   If fraConfiguracionHorarios.Visible Then
        fraConfiguracionHorarios.Visible = False
    Else
        MostrarDisponibilidad
        CargarExcepcionDia MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
        fraConfiguracionHorarios.Visible = True
        fraConfiguracionHorarios.ZOrder 0
    End If
End Sub
Private Sub LimpiarConfiguracionHorarios()
    fraConfiguracionHorarios.Visible = False
End Sub


' ---------------------------------------------------------------
'  Botón "Cerrar" dentro del frame
' ---------------------------------------------------------------
Private Sub cmdCerrarDisp_Click()
    fraDisponibilidad.Visible = False
End Sub
' --------------------------------------------------------------------------
'  Configuración de la grilla de horarios (llamar desde Form_Load)
' --------------------------------------------------------------------------
Private Sub ConfigurarGrillaHorarios()
    With grdHorarios
        .FormatString = "Cod|Día|Desde|Hasta"
        .Cols = 4
        .rows = 1
        .ColWidth(0) = 0        ' código día (oculto)
        .ColWidth(1) = 900     ' nombre del día
        .ColWidth(2) = 700     ' hora inicio
        .ColWidth(3) = 700     ' hora fin
        .BorderStyle = flexBorderNone
        .SelectionMode = flexSelectionByRow
        
        ' formato del encabezado
        .row = 0
        Dim c As Integer
        For c = 0 To 1
            .Col = c
            .CellForeColor = &HFFFFFF
            .CellBackColor = &H808080
            .CellFontBold = True
        Next c
        
        
        ' fila vacía separadora (misma lógica que grdMotivo)
        .AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
    End With
End Sub


' --------------------------------------------------------------------------
'  Cargar el combo de días de la semana (llamar desde Form_Load)
' --------------------------------------------------------------------------
Private Sub CargarCboDiaSemana()
    cboDiaSemana.Clear
    cboDiaSemana.AddItem "Lunes"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 1
    cboDiaSemana.AddItem "Martes"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 2
    cboDiaSemana.AddItem "Miércoles"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 3
    cboDiaSemana.AddItem "Jueves"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 4
    cboDiaSemana.AddItem "Viernes"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 5
    cboDiaSemana.AddItem "Sábado"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 6
    cboDiaSemana.AddItem "Domingo"
    cboDiaSemana.ItemData(cboDiaSemana.NewIndex) = 7
End Sub


' --------------------------------------------------------------------------
'  Función auxiliar: número de día -> nombre
' --------------------------------------------------------------------------
Private Function NombreDia(numDia As Integer) As String
    Select Case numDia
        Case 1: NombreDia = "Lunes"
        Case 2: NombreDia = "Martes"
        Case 3: NombreDia = "Miércoles"
        Case 4: NombreDia = "Jueves"
        Case 5: NombreDia = "Viernes"
        Case 6: NombreDia = "Sábado"
        Case 7: NombreDia = "Domingo"
        Case Else: NombreDia = ""
    End Select
End Function


' --------------------------------------------------------------------------
'  Validar formato de hora "HH:MM"
' --------------------------------------------------------------------------
Private Function ValidarFormatoHora(ByVal Hora As String) As Boolean
    Dim h As Integer
    Dim m As Integer

    ValidarFormatoHora = False

    ' Si tiene caracteres de prompt sin llenar, está incompleta
    If InStr(Hora, "_") > 0 Then Exit Function
    If Len(Hora) <> 5 Then Exit Function

    On Error GoTo SalirFalso
    h = CInt(Left(Hora, 2))
    m = CInt(Right(Hora, 2))

    If h < 0 Or h > 23 Then Exit Function
    If m < 0 Or m > 59 Then Exit Function

    ValidarFormatoHora = True
    Exit Function

SalirFalso:
    ValidarFormatoHora = False
End Function


' --------------------------------------------------------------------------
'  Cargar horarios existentes desde la BD a la grilla y  al diccionario
' --------------------------------------------------------------------------
Public Sub CargarHorarios(vencod As Integer)
    ' Resetear diccionario
    Set dicHorariosDoc = CreateObject("Scripting.Dictionary")
    
    ' Resetear grilla
    ConfigurarGrillaHorarios
    
    Dim recH As ADODB.Recordset
    Set recH = New ADODB.Recordset
    Dim sqlH As String
    Dim clave As String
    
    sqlH = "SELECT HOR_DIASEMANA, HOR_HORAINICIO, HOR_HORAFIN" _
         & " FROM HORARIO_VENDEDOR" _
         & " WHERE VEN_CODIGO = " & vencod _
         & " ORDER BY HOR_DIASEMANA, HOR_HORAINICIO"
    
    recH.Open sqlH, DBConn, adOpenStatic, adLockReadOnly
    
    Do While recH.EOF = False
        clave = CStr(recH!HOR_DIASEMANA)
        
        ' Cargar diccionario
        If dicHorariosDoc.Exists(clave) Then
            dicHorariosDoc(clave) = dicHorariosDoc(clave) & ";" _
                & Trim(recH!HOR_HORAINICIO) & "-" & Trim(recH!HOR_HORAFIN)
        Else
            dicHorariosDoc.Add clave, _
                Trim(recH!HOR_HORAINICIO) & "-" & Trim(recH!HOR_HORAFIN)
        End If
        
        ' Cargar grilla
        grdHorarios.AddItem clave & Chr(9) _
                          & NombreDiaSemana(CInt(recH!HOR_DIASEMANA)) & Chr(9) _
                          & Trim(recH!HOR_HORAINICIO) & Chr(9) _
                          & Trim(recH!HOR_HORAFIN)
        
        recH.MoveNext
    Loop
    
    recH.Close
    Set recH = Nothing
End Sub
' --------------------------------------------------------------------------
'  Borrar todos los horarios de un vendedor en la BD
' --------------------------------------------------------------------------
Private Sub BorrarHorarios(vencod As Integer)
    Dim sqlH As String
    sqlH = "DELETE FROM HORARIO_VENDEDOR WHERE VEN_CODIGO = " & vencod
    DBConn.Execute sqlH
End Sub
'  Agregar horario a la grilla + refrescar resumen
' ------------------------------------------------------------------
Private Sub cmdAgregarHorario_Click()
    Dim numDia As Integer
    Dim i As Integer
    
    If cboDiaSemana.ListIndex = -1 Then
        Beep
        MsgBox "Seleccione un día de la semana.", vbOKOnly + vbCritical, App.Title
        cboDiaSemana.SetFocus
        Exit Sub
    End If
    
    If Not ValidarFormatoHora(mskHoraDesde.text) Then
        Beep
        MsgBox "Ingrese la hora de inicio completa (HH:MM).", vbOKOnly + vbCritical, App.Title
        mskHoraDesde.SetFocus
        Exit Sub
    End If
    
    If Not ValidarFormatoHora(mskHoraHasta.text) Then
        Beep
        MsgBox "Ingrese la hora de fin completa (HH:MM).", vbOKOnly + vbCritical, App.Title
        mskHoraHasta.SetFocus
        Exit Sub
    End If
    
    If mskHoraHasta.text <= mskHoraDesde.text Then
        Beep
        MsgBox "La hora 'Hasta' debe ser mayor a la hora 'Desde'.", _
               vbOKOnly + vbCritical, App.Title
        mskHoraHasta.SetFocus
        Exit Sub
    End If
    
    numDia = cboDiaSemana.ItemData(cboDiaSemana.ListIndex)
    
    ' Verificar solapamiento
    For i = 2 To grdHorarios.rows - 1
        If grdHorarios.TextMatrix(i, 0) <> "" Then
            If CInt(grdHorarios.TextMatrix(i, 0)) = numDia Then
                If mskHoraDesde.text < grdHorarios.TextMatrix(i, 3) And _
                   mskHoraHasta.text > grdHorarios.TextMatrix(i, 2) Then
                    Beep
                    MsgBox "El horario se superpone con otro bloque" & Chr(13) _
                         & "del " & cboDiaSemana.text & " (" _
                         & grdHorarios.TextMatrix(i, 2) & " a " _
                         & grdHorarios.TextMatrix(i, 3) & ").", _
                           vbOKOnly + vbCritical, App.Title
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    grdHorarios.AddItem CStr(numDia) & Chr(9) _
                      & cboDiaSemana.text & Chr(9) _
                      & mskHoraDesde.text & Chr(9) _
                      & mskHoraHasta.text
    
    ' Limpiar controles
    cboDiaSemana.ListIndex = -1
    mskHoraDesde.Mask = ""
    mskHoraDesde.text = ""
    mskHoraDesde.Mask = "##:##"
    mskHoraHasta.Mask = ""
    mskHoraHasta.text = ""
    mskHoraHasta.Mask = "##:##"
    
    ' Refrescar resumen visual
    MostrarDisponibilidad
End Sub
' ------------------------------------------------------------------
'  Quitar horario de la grilla + refrescar resumen
' ------------------------------------------------------------------
Private Sub cmdQuitarHorario_Click()
    If grdHorarios.row < 2 Then
        Beep
        MsgBox "Seleccione un horario de la grilla para quitar.", _
               vbOKOnly + vbCritical, App.Title
        Exit Sub
    End If
    
    If grdHorarios.TextMatrix(grdHorarios.row, 0) = "" Then
        Exit Sub
    End If
    
    grdHorarios.RemoveItem grdHorarios.row
    
    ' Refrescar resumen visual
    MostrarDisponibilidad
End Sub


' ------------------------------------------------------------------
'  Guardar disponibilidad en BD
'  (borrar todo y reinsertar desde grdHorarios)
' ------------------------------------------------------------------
Private Sub cmdGuardarDisponibilidad_Click()
    Dim vencod As Integer
    vencod = cboDoctor.ItemData(cboDoctor.ListIndex)
    
    If vencod = 0 Then Exit Sub
    
    On Error GoTo ErrorGuardar
    DBConn.BeginTrans
    
    ' 1. Borrar horarios existentes
    Dim sqlH As String
    sqlH = "DELETE FROM HORARIO_VENDEDOR WHERE VEN_CODIGO = " & vencod
    DBConn.Execute sqlH
    
    ' 2. Insertar desde la grilla
    Dim i As Integer
    For i = 2 To grdHorarios.rows - 1
        If grdHorarios.TextMatrix(i, 0) <> "" Then
            sqlH = "INSERT INTO HORARIO_VENDEDOR" _
                 & " (VEN_CODIGO, HOR_DIASEMANA, HOR_HORAINICIO, HOR_HORAFIN)" _
                 & " VALUES (" _
                 & vencod & ", " _
                 & grdHorarios.TextMatrix(i, 0) & ", " _
                 & "'" & grdHorarios.TextMatrix(i, 2) & "', " _
                 & "'" & grdHorarios.TextMatrix(i, 3) & "')"
            DBConn.Execute sqlH
        End If
    Next i
    
    DBConn.CommitTrans
    
    ' 3. Refrescar el diccionario en memoria
    CargarHorarios vencod
    
    ' 4. Refrescar la grilla de turnos
    ' BuscarTurnos fechaActual, vencod, modoActual   ? reemplazar
    
    MsgBox "Disponibilidad guardada correctamente.", vbInformation, App.Title
    
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    Exit Sub
    
ErrorGuardar:
    DBConn.RollbackTrans
    MsgBox "Error al guardar: " & Err.Description, vbCritical, App.Title
End Sub


Private Sub mskHoraDesde_Change()
    cmdGuardarDisponibilidad.Enabled = True
End Sub

Private Sub mskHoraHasta_Change()
    cmdGuardarDisponibilidad.Enabled = True
End Sub

Private Sub cboDiaSemana_Click()
    cmdGuardarDisponibilidad.Enabled = True
End Sub

Private Sub mskHoraDesde_GotFocus()
    seltxt
End Sub

Private Sub mskHoraHasta_GotFocus()
    seltxt
End Sub

' ------------------------------------------------------------------
'  EXCEPCIONES DE HORARIO ATENCIÓN
' ------------------------------------------------------------------

' ------------------------------------------------------------------
'  2. CONFIGURAR GRILLA DE EXCEPCIONES (llamar en Form_Load)
' ------------------------------------------------------------------
Private Sub ConfigurarGrillaHorariosExc()
    With grdHorariosExc
        .FormatString = "Desde|Hasta"
        .Cols = 2
        .rows = 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .BorderStyle = flexBorderNone
        .SelectionMode = flexSelectionByRow
        .row = 0
        Dim c As Integer
        For c = 0 To .Cols - 1
            .Col = c
            .CellForeColor = &HFFFFFF
            .CellBackColor = &H808080
            .CellFontBold = True
        Next c
    End With
End Sub


' ------------------------------------------------------------------
'  3. CARGAR EXCEPCIÓN PARA UNA FECHA
'     Llamar al seleccionar fecha en el calendario
' ------------------------------------------------------------------
Public Sub CargarExcepcionDia(Fecha As Date, vencod As Integer)
    Dim recExc As ADODB.Recordset
    Dim sqlExc As String
    Dim Motivo As String
    
    ' Resetear grilla
    ConfigurarGrillaHorariosExc
    vExcDiaId = 0
    
    ' Actualizar label del día
    Dim diaSem As Integer
    diaSem = Weekday(Fecha, vbMonday)
    lblDiaActual.Caption = NombreDiaSemana(diaSem) & " " _
        & Day(Fecha) & " de " & MonthName(Month(Fecha), False) _
        & " de " & Year(Fecha)
    
    ' Buscar excepción en la BD
    Set recExc = New ADODB.Recordset
    sqlExc = "SELECT EXD_ID, EXD_ATIENDE, EXD_MOTIVO FROM EXCEPCION_DIA_VENDEDOR" _
           & " WHERE VEN_CODIGO = " & vencod _
           & " AND EXD_FECHA = " & XDQ(Fecha)
    recExc.Open sqlExc, DBConn, adOpenStatic, adLockReadOnly
    
    If Not recExc.EOF Then
        ' === HAY EXCEPCIÓN ===
        vExcDiaId = recExc!EXD_ID
        Motivo = ChkNull(recExc!EXD_MOTIVO)
        
        txtMotivoExcepcion.text = ChkNull(recExc!EXD_MOTIVO)
        txtMotivoExcepcion.Enabled = True
        
        If recExc!EXD_ATIENDE = "S" Then
            ' Atiende con horario especial
           optAtiende.Value = True
            recExc.Close
            
            ' Cargar bloques
            sqlExc = "SELECT EXH_HORAINICIO, EXH_HORAFIN" _
                   & " FROM EXCEPCION_HORARIO_VENDEDOR" _
                   & " WHERE EXD_ID = " & vExcDiaId _
                   & " ORDER BY EXH_HORAINICIO"
            recExc.Open sqlExc, DBConn, adOpenStatic, adLockReadOnly
            
            Do While recExc.EOF = False
                grdHorariosExc.AddItem _
                    Trim(recExc!EXH_HORAINICIO) & Chr(9) _
                  & Trim(recExc!EXH_HORAFIN)
                recExc.MoveNext
            Loop
            recExc.Close
            
            HabilitarControlesExc True
        Else
            ' No atiende (excepción de bloqueo)
            optNoAtiende.Value = True
            recExc.Close
            HabilitarControlesExc False
        End If
        
        txtMotivoExcepcion.text = Motivo
        
        ' Mostrar botón para eliminar la excepción
        cmdEliminarExcepcion.Enabled = True
    Else
        ' === NO HAY EXCEPCIÓN ? estado según horario habitual ===
        recExc.Close
        
        Dim clave As String
        clave = CStr(diaSem)
        
        If Not dicHorariosDoc Is Nothing Then
            If dicHorariosDoc.Count > 0 And dicHorariosDoc.Exists(clave) Then
                ' Día habitual: check marcado, controles deshabilitados
                optAtiende.Value = True
            Else
                ' Día no habitual: check desmarcado
                optNoAtiende.Value = True
            End If
        Else
            optNoAtiende.Value = True
        End If
        
        HabilitarControlesExc False
        cmdEliminarExcepcion.Enabled = False
        txtMotivoExcepcion.text = ""
        txtMotivoExcepcion.Enabled = False
    End If
    
    Set recExc = Nothing
End Sub


' ------------------------------------------------------------------
'  Habilitar/deshabilitar controles de hora de excepción
' ------------------------------------------------------------------
Private Sub HabilitarControlesExc(ByVal habilitar As Boolean)
    mskHoraDesdeExc.Enabled = habilitar
    mskHoraHastaExc.Enabled = habilitar
    cmdAgregarExc.Enabled = habilitar
    cmdQuitarExc.Enabled = habilitar
End Sub


' ------------------------------------------------------------------
'  4. CHECK "ATIENDE" - toggle controles
' ------------------------------------------------------------------
Private Sub optAtiende_Click()
    HabilitarControlesExc True
    txtMotivoExcepcion.Enabled = True
End Sub

Private Sub optNoAtiende_Click()
    HabilitarControlesExc False
    ConfigurarGrillaHorariosExc
    txtMotivoExcepcion.Enabled = True
End Sub

' ------------------------------------------------------------------
'  5. AGREGAR HORARIO A EXCEPCIÓN
' ------------------------------------------------------------------
Private Sub cmdAgregarExc_Click()
    If Not ValidarFormatoHora(mskHoraDesdeExc.text) Then
        Beep
        MsgBox "Ingrese la hora de inicio completa (HH:MM).", _
               vbOKOnly + vbCritical, App.Title
        mskHoraDesdeExc.SetFocus
        Exit Sub
    End If
    
    If Not ValidarFormatoHora(mskHoraHastaExc.text) Then
        Beep
        MsgBox "Ingrese la hora de fin completa (HH:MM).", _
               vbOKOnly + vbCritical, App.Title
        mskHoraHastaExc.SetFocus
        Exit Sub
    End If
    
    If mskHoraHastaExc.text <= mskHoraDesdeExc.text Then
        Beep
        MsgBox "La hora 'Hasta' debe ser mayor a la hora 'Desde'.", _
               vbOKOnly + vbCritical, App.Title
        mskHoraHastaExc.SetFocus
        Exit Sub
    End If
    
    ' Verificar solapamiento
    Dim i As Integer
    For i = 1 To grdHorariosExc.rows - 1
        If grdHorariosExc.TextMatrix(i, 0) <> "" Then
            If mskHoraDesdeExc.text < grdHorariosExc.TextMatrix(i, 1) And _
               mskHoraHastaExc.text > grdHorariosExc.TextMatrix(i, 0) Then
                Beep
                MsgBox "El horario se superpone con otro bloque (" _
                     & grdHorariosExc.TextMatrix(i, 0) & " a " _
                     & grdHorariosExc.TextMatrix(i, 1) & ").", _
                       vbOKOnly + vbCritical, App.Title
                Exit Sub
            End If
        End If
    Next i
    
    grdHorariosExc.AddItem mskHoraDesdeExc.text & Chr(9) & mskHoraHastaExc.text
    
    ' Limpiar
    mskHoraDesdeExc.Mask = ""
    mskHoraDesdeExc.text = ""
    mskHoraDesdeExc.Mask = "##:##"
    mskHoraHastaExc.Mask = ""
    mskHoraHastaExc.text = ""
    mskHoraHastaExc.Mask = "##:##"
End Sub


' ------------------------------------------------------------------
'  QUITAR HORARIO DE EXCEPCIÓN
' ------------------------------------------------------------------
Private Sub cmdQuitarExc_Click()
    If grdHorariosExc.row < 1 Then
        Beep
        MsgBox "Seleccione un horario para quitar.", _
               vbOKOnly + vbCritical, App.Title
        Exit Sub
    End If
    If grdHorariosExc.TextMatrix(grdHorariosExc.row, 0) = "" Then
        Exit Sub
    End If
    grdHorariosExc.RemoveItem grdHorariosExc.row
End Sub


' ------------------------------------------------------------------
'  6. GUARDAR EXCEPCIÓN EN BD
' ------------------------------------------------------------------
Private Sub cmdGuardarExcepcion_Click()
    Dim vencod As Integer
    Dim Fecha As Date
    vencod = cboDoctor.ItemData(cboDoctor.ListIndex)
    Fecha = MViewFecha.Value
    
    If vencod = 0 Then Exit Sub
    
    On Error GoTo ErrorGuardarExc
    DBConn.BeginTrans
    
    Dim sqlExc As String
    
    ' 1. Borrar excepción anterior si existía
    If vExcDiaId > 0 Then
        sqlExc = "DELETE FROM EXCEPCION_HORARIO_VENDEDOR WHERE EXD_ID = " & vExcDiaId
        DBConn.Execute sqlExc
        sqlExc = "DELETE FROM EXCEPCION_DIA_VENDEDOR WHERE EXD_ID = " & vExcDiaId
        DBConn.Execute sqlExc
    End If
    
    ' 2. Insertar cabecera
    Dim atiende As String
    If optAtiende.Value = True Then
        atiende = "S"
    Else
        atiende = "N"
    End If
    
    sqlExc = "INSERT INTO EXCEPCION_DIA_VENDEDOR" _
       & " (VEN_CODIGO, EXD_FECHA, EXD_ATIENDE, EXD_MOTIVO)" _
       & " VALUES (" & vencod & ", " & XDQ(Fecha) & ", '" & atiende & "', " _
       & XS(txtMotivoExcepcion.text) & ")"
    DBConn.Execute sqlExc
    
    ' 3. Obtener el ID recién insertado
    Dim recId As ADODB.Recordset
    Set recId = New ADODB.Recordset
    recId.Open "SELECT @@IDENTITY AS NuevoId", DBConn, adOpenStatic, adLockReadOnly
    vExcDiaId = recId!nuevoid
    recId.Close
    Set recId = Nothing
    
    ' 4. Insertar bloques horarios (solo si atiende)
    If atiende = "S" Then
        Dim i As Integer
        For i = 1 To grdHorariosExc.rows - 1
            If grdHorariosExc.TextMatrix(i, 0) <> "" Then
                sqlExc = "INSERT INTO EXCEPCION_HORARIO_VENDEDOR" _
                       & " (EXD_ID, EXH_HORAINICIO, EXH_HORAFIN)" _
                       & " VALUES (" & vExcDiaId & ", " _
                       & "'" & grdHorariosExc.TextMatrix(i, 0) & "', " _
                       & "'" & grdHorariosExc.TextMatrix(i, 1) & "')"
                DBConn.Execute sqlExc
            End If
        Next i
    End If
    
    DBConn.CommitTrans
    
    cmdEliminarExcepcion.Enabled = True
    
    MsgBox "Excepción guardada correctamente.", vbInformation, App.Title
    
    ' 5. Refrescar turnos
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    
    Exit Sub
    
ErrorGuardarExc:
    DBConn.RollbackTrans
    MsgBox "Error al guardar: " & Err.Description, vbCritical, App.Title
End Sub
Private Function ObtenerMotivoExcepcion(vencod As Integer, Fecha As Date) As String
    ObtenerMotivoExcepcion = ""
    
    Dim recExc As ADODB.Recordset
    Set recExc = New ADODB.Recordset
    Dim sqlExc As String
    
    sqlExc = "SELECT EXD_ATIENDE, EXD_MOTIVO FROM EXCEPCION_DIA_VENDEDOR" _
           & " WHERE VEN_CODIGO = " & vencod _
           & " AND EXD_FECHA = " & XDQ(Fecha)
    recExc.Open sqlExc, DBConn, adOpenStatic, adLockReadOnly
    
    If Not recExc.EOF Then
        If Not IsNull(recExc!EXD_MOTIVO) Then
            If Trim(recExc!EXD_MOTIVO) <> "" Then
                If recExc!EXD_ATIENDE = "S" Then
                    ObtenerMotivoExcepcion = "El doctor ATIENDE debido a: " _
                        & recExc!EXD_MOTIVO
                Else
                    ObtenerMotivoExcepcion = "El doctor NO ATIENDE debido a: " _
                        & recExc!EXD_MOTIVO
                End If
            End If
        End If
    End If
    
    recExc.Close
    Set recExc = Nothing
End Function


' ------------------------------------------------------------------
'  7. ELIMINAR EXCEPCIÓN (volver al horario habitual)
' ------------------------------------------------------------------
Private Sub cmdEliminarExcepcion_Click()
    If vExcDiaId = 0 Then Exit Sub
    
    If MsgBox("¿Eliminar la excepción y volver al horario habitual?", _
              vbYesNo + vbQuestion, App.Title) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorEliminarExc
    DBConn.BeginTrans
    
    Dim sqlExc As String
    sqlExc = "DELETE FROM EXCEPCION_HORARIO_VENDEDOR WHERE EXD_ID = " & vExcDiaId
    DBConn.Execute sqlExc
    sqlExc = "DELETE FROM EXCEPCION_DIA_VENDEDOR WHERE EXD_ID = " & vExcDiaId
    DBConn.Execute sqlExc
    
    DBConn.CommitTrans
    
    ' Recargar estado del día (vuelve al habitual)
    Dim Fecha As Date
    Fecha = MViewFecha.Value
    CargarExcepcionDia MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    
    ' Refrescar turnos
    BuscarTurnos Fecha, cboDoctor.ItemData(cboDoctor.ListIndex)
    
    Exit Sub
    
ErrorEliminarExc:
    DBConn.RollbackTrans
    MsgBox "Error al eliminar: " & Err.Description, vbCritical, App.Title
End Sub


' ------------------------------------------------------------------
'  8. OBTENER BLOQUES DEL DÍA (excepción > habitual)
'
'     Devuelve "" si no atiende
'     Devuelve "08:00-12:00;14:00-18:00" si atiende
'     Prioriza excepción sobre horario habitual
' ------------------------------------------------------------------
Public Function ObtenerBloquesDelDia(vencod As Integer, Fecha As Date) As String
    ObtenerBloquesDelDia = ""
    
    Dim recExc As ADODB.Recordset
    Dim sqlExc As String
    
    ' --- 1. Buscar excepción ---
    Set recExc = New ADODB.Recordset
    sqlExc = "SELECT EXD_ID, EXD_ATIENDE FROM EXCEPCION_DIA_VENDEDOR" _
           & " WHERE VEN_CODIGO = " & vencod _
           & " AND EXD_FECHA = " & XDQ(Fecha)
    recExc.Open sqlExc, DBConn, adOpenStatic, adLockReadOnly
    
    If Not recExc.EOF Then
        ' Hay excepción
        If recExc!EXD_ATIENDE = "N" Then
            ' No atiende
            recExc.Close
            Set recExc = Nothing
            Exit Function   ' devuelve ""
        End If
        
        ' Atiende con horario especial: cargar bloques
        Dim exdId As Long
        exdId = recExc!EXD_ID
        recExc.Close
        
        sqlExc = "SELECT EXH_HORAINICIO, EXH_HORAFIN" _
               & " FROM EXCEPCION_HORARIO_VENDEDOR" _
               & " WHERE EXD_ID = " & exdId _
               & " ORDER BY EXH_HORAINICIO"
        recExc.Open sqlExc, DBConn, adOpenStatic, adLockReadOnly
        
        Dim bloques As String
        bloques = ""
        Do While recExc.EOF = False
            If bloques <> "" Then bloques = bloques & ";"
            bloques = bloques & Trim(recExc!EXH_HORAINICIO) _
                    & "-" & Trim(recExc!EXH_HORAFIN)
            recExc.MoveNext
        Loop
        recExc.Close
        Set recExc = Nothing
        
        ObtenerBloquesDelDia = bloques
        Exit Function
    End If
    
    recExc.Close
    Set recExc = Nothing
    
    ' --- 2. Sin excepción ? horario habitual ---
    If dicHorariosDoc Is Nothing Then Exit Function
    If dicHorariosDoc.Count = 0 Then Exit Function
    
    Dim diaSemana As Integer
    diaSemana = Weekday(Fecha, vbMonday)
    Dim clave As String
    clave = CStr(diaSemana)
    
    If dicHorariosDoc.Exists(clave) Then
        ObtenerBloquesDelDia = dicHorariosDoc(clave)
    End If
End Function


' ------------------------------------------------------------------
'  9. DOCTOR ATIENDE EL DÍA (versión actualizada)
'     Reemplaza la versión anterior
' ------------------------------------------------------------------
Public Function DoctorAtiendeElDia(Fecha As Date) As Boolean
    DoctorAtiendeElDia = True   ' default: permitir
    
    ' --- 1. Excepción siempre tiene prioridad ---
    Dim recExc As ADODB.Recordset
    Set recExc = New ADODB.Recordset
    Dim sqlExc As String
    
    sqlExc = "SELECT EXD_ATIENDE FROM EXCEPCION_DIA_VENDEDOR" _
           & " WHERE VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex) _
           & " AND EXD_FECHA = " & XDQ(Fecha)
    recExc.Open sqlExc, DBConn, adOpenStatic, adLockReadOnly
    
    If Not recExc.EOF Then
        ' Hay excepción: la respuesta es directa
        DoctorAtiendeElDia = (recExc!EXD_ATIENDE = "S")
        recExc.Close
        Set recExc = Nothing
        Exit Function
    End If
    
    recExc.Close
    Set recExc = Nothing
    
    ' --- 2. Sin excepción ? horario habitual ---
    If dicHorariosDoc Is Nothing Then Exit Function
    If dicHorariosDoc.Count = 0 Then Exit Function
    
    Dim diaSemana As Integer
    diaSemana = Weekday(Fecha, vbMonday)
    DoctorAtiendeElDia = dicHorariosDoc.Exists(CStr(diaSemana))
End Function

