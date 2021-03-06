VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIGOR - Turnos de Pacientes"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18345
   ForeColor       =   &H00000000&
   Icon            =   "frmTurnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   18345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProtocolos 
      Enabled         =   0   'False
      Height          =   375
      Left            =   17400
      Picture         =   "frmTurnos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Protocolos"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdCopiar 
      Height          =   375
      Left            =   16440
      Picture         =   "frmTurnos.frx":2004
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Copiar Turnos"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdCortar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   16920
      Picture         =   "frmTurnos.frx":238E
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Cortar Turnos"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdImpTurno 
      Enabled         =   0   'False
      Height          =   375
      Left            =   15960
      Picture         =   "frmTurnos.frx":2718
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "ImprimirTurno"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdOcultar 
      Height          =   375
      Left            =   17880
      TabIndex        =   52
      Top             =   50
      Width           =   375
   End
   Begin MSComCtl2.DTPicker fechaturno 
      Height          =   375
      Left            =   13560
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58064897
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
      Left            =   10920
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtfiltrop 
         Height          =   315
         Left            =   1800
         TabIndex        =   47
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdSalirP 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   5760
         TabIndex        =   46
         Top             =   8040
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdProtocolos 
         Height          =   7230
         Left            =   120
         TabIndex        =   44
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
         TabIndex        =   45
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Filtro"
         Height          =   195
         Left            =   1320
         TabIndex        =   48
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdatendido 
      BackColor       =   &H0000C000&
      Height          =   315
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Atendido"
      Top             =   450
      Width           =   495
   End
   Begin VB.CommandButton cmdespera 
      BackColor       =   &H0000FFFF&
      Height          =   315
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "En Espera"
      Top             =   450
      Width           =   495
   End
   Begin VB.CommandButton cmdpendiente 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   5400
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   40
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
      TabIndex        =   33
      Tag             =   "Descripción"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   5040
      Picture         =   "frmTurnos.frx":832A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Reporte"
      Height          =   735
      Left            =   4080
      Picture         =   "frmTurnos.frx":936C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Listado de Turnos del dia por Doctor"
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   3120
      Picture         =   "frmTurnos.frx":A036
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar Turnos"
      Height          =   735
      Left            =   2160
      Picture         =   "frmTurnos.frx":B078
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8760
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
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   2895
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   250
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
      Height          =   5310
      Left            =   120
      TabIndex        =   16
      Top             =   3400
      Width           =   3495
      Begin VB.ComboBox cboMotivo 
         BackColor       =   &H0080FF80&
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2280
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
         TabIndex        =   57
         ToolTipText     =   "Agregar nuevo Paciente"
         Top             =   250
         Width           =   255
      End
      Begin MSMask.MaskEdBox mebHoraD 
         Height          =   315
         Left            =   675
         TabIndex        =   6
         Top             =   4560
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
         TabIndex        =   5
         Tag             =   "Descripción"
         Top             =   3660
         Width           =   3270
      End
      Begin VB.OptionButton optNO 
         Caption         =   "NO"
         Height          =   315
         Left            =   2040
         TabIndex        =   38
         Top             =   1560
         Width           =   615
      End
      Begin VB.OptionButton optSI 
         Caption         =   "SI"
         Height          =   315
         Left            =   1440
         TabIndex        =   37
         Top             =   1560
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
         TabIndex        =   29
         Tag             =   "Descripción"
         Top             =   1920
         Width           =   3270
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1920
         TabIndex        =   23
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
         TabIndex        =   28
         Tag             =   "Descripción"
         Top             =   1080
         Width           =   1875
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
         Tag             =   "Descripción"
         Top             =   735
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
         Height          =   555
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "Descripción"
         Top             =   2640
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
         TabIndex        =   10
         Tag             =   "Descripción"
         Text            =   "0,00"
         Top             =   4935
         Width           =   1635
      End
      Begin MSMask.MaskEdBox mebHoraH 
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   4560
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
         TabIndex        =   9
         Top             =   4620
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
         TabIndex        =   8
         Top             =   4920
         Visible         =   0   'False
         Width           =   1260
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
         TabIndex        =   50
         Top             =   4590
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
         TabIndex        =   49
         Top             =   4590
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
         TabIndex        =   34
         Top             =   4935
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
         TabIndex        =   39
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
         TabIndex        =   31
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tel�fono:"
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
         TabIndex        =   30
         Top             =   1080
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
         TabIndex        =   21
         Top             =   4140
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
         TabIndex        =   20
         Top             =   2280
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
         TabIndex        =   19
         Top             =   250
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   2895
      Begin MSComCtl2.MonthView MViewFecha 
         Height          =   2370
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   58064898
         CurrentDate     =   40049
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdGrilla 
      Height          =   7965
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "Doble Click para ver la Historia Clinica del Paciente"
      Top             =   765
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   14049
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
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   735
      Left            =   240
      Picture         =   "frmTurnos.frx":B402
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   735
      Left            =   1200
      Picture         =   "frmTurnos.frx":B78C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8760
      Width           =   975
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
      TabIndex        =   36
      Top             =   8880
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
      TabIndex        =   35
      Top             =   11160
      Width           =   1245
   End
   Begin VB.Label lblAux 
      Caption         =   "Label7"
      Height          =   255
      Left            =   12120
      TabIndex        =   32
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
      TabIndex        =   27
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
      Left            =   8880
      TabIndex        =   26
      Top             =   450
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
      TabIndex        =   17
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
      Width           =   12045
   End
End
Attribute VB_Name = "frmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Dim J As Integer
Dim hDesde As Integer
Dim hHasta As Integer
Dim ActivoGrid As Integer ' 1 actio 0 desactivo
Dim sAction As String
Dim dFechaCopy As String
Dim nDoctorCopy As String
Dim sNameDoctorCopy As String



Private Sub cboDesde_LostFocus()
    If cboDesde.ListIndex < cboDesde.ListCount - 1 Then
        cbohasta.ListIndex = cboDesde.ListIndex + 1
    Else
        cbohasta.ListIndex = cboDesde.ListIndex
    End If
End Sub

Private Sub cboDoctor_Click()
    LimpiarComboMotivo
    If cboDoctor.ListIndex <> -1 Then
        sql = "SELECT M.MOT_DESCRI"
            sql = sql & " FROM  MOTIVO_VENDEDOR MV,VENDEDOR V,MOTIVO M "
            sql = sql & " WHERE V.VEN_NOMBRE = " & XS(cboDoctor.Text)
            sql = sql & " AND V.VEN_CODIGO = MV.VEN_CODIGO"
            sql = sql & " AND MV.MOT_CODIGO = M.MOT_CODIGO"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
        Do While Rec1.EOF = False
                cboMotivo.AddItem Rec1!MOT_DESCRI
                Rec1.MoveNext
        Loop
        Rec1.Close
    End If
End Sub
Private Sub LimpiarComboMotivo()
    cboMotivo.Clear
End Sub

Private Sub cboDoctor_LostFocus()
    'LimpiarTurno
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
If mebHoraH.Text <= mebHoraD.Text Then
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
hasta = mebHoraH.Text
desde = cboDesde.Text
If grdGrilla.Rows < 2 Then
    ValidarRangoTurno = True
Else
   For i = 1 To grdGrilla.Rows - 1
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
    If txtBuscaCliente.Text = "" Then
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
    If mebHoraD.Text = "" Then
        MsgBox "No ha ingresado la hora de comienzo del Turno", vbCritical, TIT_MSGBOX
        mebHoraD.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If mebHoraH.Text = "" Then
        MsgBox "No ha ingresado la hora de finalizaci�n del Turno", vbCritical, TIT_MSGBOX
        mebHoraH.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If mebHoraD.Text >= mebHoraH.Text Then
        MsgBox "La hora HASTA debe ser mayor a la hora DESDE", vbCritical, TIT_MSGBOX
        mebHoraD.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    
    ValidarTurno = True
End Function
Private Function ImprimirTurno()
    Dim sHoraD As Date
    Dim mNombreImpresora As String
    Dim strAutoSaveDirectory As String
    Dim strAutoSaveFileName As String
    Dim cliente As String
    Dim Fecha As String
    Dim objWSH As Object
    
    cliente = Replace(txtBuscarCliDescri.Text, " ", "_")
    Fecha = Replace(MViewFecha.Value, "/", "")
    
    sHoraD = mebHoraD.Text
    sHoraD = Mid(mebHoraD, 1, 1)
    
    If sHoraD = "0" Then
        sHoraD = Mid(mebHoraD.Text, 2, 4)
    Else
        sHoraD = Mid(mebHoraD.Text, 1, 5)
    End If
    
    mNombreImpresora = Printer.DeviceName
    Rep.Destination = 1
    Call EstableceDefaultPrinter("PDFCreator")

    
    strAutoSaveDirectory = DirReport & "\Turnos\"
    strAutoSaveFileName = "TURNO_" & cliente & "_" & Fecha & ".pdf"
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
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.CLI_CODIGO}= " & XN(txtCodigo.Text)
    'Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.TUR_DESDE}= TIME (" & Mid(mebHoraD.Text, 1, 2) & "," & Mid(mebHoraD.Text, 4, 2) & ",00)"  '& grdGrilla.RowSel

    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    'Rep.Connect = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & SERVIDOR & ";"
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
        Case 36 ' malena baudo
            Rep.ReportFileName = DirReport & "rptTurno_baudo.rpt"
        Case 21 ' Lorena Sugar
            Rep.ReportFileName = DirReport & "rptTurno_lorenasugar.rpt"
        Case 33 ' Melina Orlietti
            Rep.ReportFileName = DirReport & "rptTurno_melinaorlietti.rpt"
        Case Else
            Rep.ReportFileName = DirReport & "rptTurno.rpt"
    End Select
    
    Rep.Action = 1

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
    txtMotivo.Text = cboMotivo.Text
    
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

Private Sub cmdAgregar_Click()
    Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    Dim a�os, edad As Integer
    Dim Fecha As Date
    'Validar los campos requeridos
    If ValidarTurno = False Then Exit Sub
    'If ValidarHorarioTurno = False Then Exit Sub
    If MsgBox("�Confirma el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'agregar teniendo en cuentas loc combos de horas
    On Error GoTo HayErrorTurno
    
    grdGrilla.HighLight = flexHighlightAlways
    
    'nFilaD = cboDesde.ListIndex
    'nFilaH = cbohasta.ListIndex
    'nFilaD = mebHoraD.Text
    'nFilaH = mebHoraH.Text
    i = 0
    
    sHoraDAux = mebHoraD.Text
    'For i = 1 To nFilaH - nFilaD
        DBConn.BeginTrans
        
        sHoraD = mebHoraD.Text
        sHoraD = Mid(sHoraD, 1, 1)
        
        If sHoraD = "0" Then
            sHoraD = Mid(mebHoraD.Text, 2, 4)
        Else
            sHoraD = Trim(mebHoraD.Text)
        End If
        
        'ACA TENGO QUE HACER UN CONTROL POR CLAVES PRIMARIAS
        sql = "SELECT * FROM TURNOS"
        sql = sql & " WHERE TUR_FECHA = " & XDQ(fechaturno.Value)
        sql = sql & " AND TUR_HORAD = #" & sHoraD & "#"
        sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
        'sql = sql & " AND CLI_CODIGO = " & XN(txtCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Not rec.EOF = False Then
            sql = "INSERT INTO TURNOS"
            sql = sql & " (TUR_FECHA, TUR_HORAD,TUR_HORAH,"
            sql = sql & " VEN_CODIGO,CLI_CODIGO,TUR_MOTIVO,TUR_DRSOLICITA,TUR_ASISTIO,TUR_OSOCIAL,TUR_TIENEMUTUAL,"
            'If User <> 99 Then
                sql = sql & " TUR_USER, "
            'End If
            sql = sql & " TUR_FECALTA, TUR_DESDE, TUR_IMPORTE)"
            sql = sql & " VALUES ("
            sql = sql & XDQ(fechaturno.Value) & ",#"
            'sql = sql & Left(Trim(grdGrilla.TextMatrix(i + nFilaD, 0)), 5) & "#,#"
            'sql = sql & Right(Trim(grdGrilla.TextMatrix(i + nFilaD, 0)), 5) & "#,"
            sql = sql & mebHoraD.Text & "#,#"
            sql = sql & mebHoraH.Text & "#,"
            sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
            sql = sql & XN(txtCodigo) & ","
            sql = sql & XS(txtMotivo) & ","
            sql = sql & XS(txtDrSolicitante) & ","
            sql = sql & 0 & ","
            'veo si es particular o con  mutual el turno
            If optSI.Value = True Then
                sql = sql & XS(txtOSocial.Text) & ","
            Else
                sql = sql & XS("PARTICULAR") & ","
            End If
            'veo si el paciente tiene o no mutuaL
            If txtOSocial.Text <> "" Then
                sql = sql & XN("1") & ","
            Else
                sql = sql & XN("0") & ","
            End If
            'If User <> 99 Then
                sql = sql & User & ","
            'End If
            sql = sql & XDQ(Date) & ","
            If i = 1 Then
                sql = sql & 1 & ","
            Else
                sql = sql & 0 & ","
            End If
            sql = sql & XN(txtimporte.Text) & ")"
            
            
        Else
            
            If MsgBox("Ya hay un turno para ese horario �Confirma la Modificaci�n del Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
                rec.Close
                Exit Sub
            End If
            ' aca hago el update
            sql = "UPDATE TURNOS SET "
            sql = sql & " CLI_CODIGO = " & XN(txtCodigo.Text) 'CAMBIAR CUANDO CARGUEMOS DNI
            sql = sql & " ,TUR_HORAD = " & "#" & mebHoraD.Text & "#"
            sql = sql & " ,TUR_HORAH = " & "#" & mebHoraH.Text & "#"
            sql = sql & " ,TUR_MOTIVO =" & XS(txtMotivo.Text)
            sql = sql & " ,TUR_DRSOLICITA =" & XS(txtDrSolicitante.Text)
            sql = sql & " ,TUR_FECALTA =" & XDQ(Date)
            If User <> 99 Then
                sql = sql & " ,TUR_USER =" & User
            End If
            sql = sql & " ,TUR_IMPORTE =" & XN(txtimporte.Text)
            'veo si es particular o con  mutual el turno
            If optSI.Value = True Then
                sql = sql & " ,TUR_OSOCIAL =" & XS(txtOSocial.Text)
            Else
                sql = sql & " ,TUR_OSOCIAL =" & XS("PARTICULAR")
            End If
            'veo si el paciente tiene o no mutuaL
            If txtOSocial.Text <> "" Then
                sql = sql & ",TUR_TIENEMUTUAL = " & XN(1)
            Else
                sql = sql & ",TUR_TIENEMUTUAL = " & XN(0)
            End If
            sql = sql & " WHERE "
            sql = sql & " TUR_FECHA = " & XDQ(fechaturno.Value)
            sql = sql & " AND TUR_HORAD = #" & mebHoraD.Text & "#"
            sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
            
        End If

        
        rec.Close
        DBConn.Execute sql
             'VER Q HAGO ACA
            'cboDesde.ListIndex = cboDesde.ListIndex + 1
        'Next
        mebHoraD.Text = sHoraDAux
        ' busco fecha nacimiento y calculo la edad
        Fecha = fechaturno.Value
        sql = "SELECT CLI_CUMPLE"
        sql = sql & " FROM  CLIENTE "
        sql = sql & " WHERE CLI_CODIGO = " & XN(txtCodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Not (IsNull(rec!CLI_CUMPLE)) Then
                If rec.EOF = False Then
                    a�os = Year(Date) - Year(rec!CLI_CUMPLE)
                    If Month(Fecha) < Month(rec!CLI_CUMPLE) Then a�os = a�os - 1 'todavia no ha llegado el mes de su cumple
                    If Month(Now) = Month(rec!CLI_CUMPLE) And Day(Fecha) < Day(rec!CLI_CUMPLE) Then a�os = a�os - 1 'es el mes pero no ha llegado el dia de su cumple
                    edad = a�os
                End If
            Else
                edad = 0
            End If
        rec.Close
        sql = "UPDATE CLIENTE SET"
        sql = sql & " CLI_EDAD=" & edad
        sql = sql & " WHERE CLI_CODIGO=" & txtCodigo.Text
        DBConn.Execute sql
        
        DBConn.CommitTrans
        BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    
        If MsgBox("�Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
'
    
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
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
    'agregar columnas en la grilla, para guardar el codigo de doctor, paciente
    
End Sub

Private Sub cmdatendido_Click()
    If grdGrilla.RowSel <> 0 Then
        'atendido
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = 1
        cambiocolor 1
        
    
        'Actualizo la Base de Datos
        sql = "UPDATE TURNOS SET "
        sql = sql & " TUR_ASISTIO =" & grdGrilla.TextMatrix(grdGrilla.RowSel, 10)
        sql = sql & " WHERE "
        sql = sql & " TUR_FECHA = " & XDQ(fechaturno.Value)
        sql = sql & " AND TUR_HORAD = #" & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        DBConn.Execute sql
    End If
End Sub

Private Sub CmdBuscar_Click()
    frmBuscarTurnos.Show vbModal
End Sub
Private Sub LimpiarTurno()
    fraprotocolos.Visible = False
    txtBuscaCliente.Text = ""
    txtBuscaCliente.ToolTipText = ""
    txtCodigo.Text = ""
    txtTelefono.Text = ""
    txtOSocial.Text = ""
    txtBuscarCliDescri.Text = ""
    txtMotivo.Text = ""
    txtDrSolicitante.Text = ""
    'cboDesde.ListIndex = -1
    'cbohasta.ListIndex = -1
    mebHoraD.Text = "__:__"
    mebHoraH.Text = "__:__"
    txtimporte.Text = "0,00"
    txtBuscaCliente.SetFocus
    cmdImpTurno.Enabled = False
    cmdCopiar.Enabled = True
    cmdCortar.Enabled = False
    cmdProtocolos.Enabled = False
    optSI.Enabled = True
    cboMotivo.ListIndex = -1
    If User = 1 Then
        cmdAgregar.Enabled = True
    Else
        cmdAgregar.Enabled = False
    End If
    
End Sub

Private Sub cmdCopiar_Click()
'    If MsgBox("Esta a punto de  Copiar los " & lbldiaTurno.Caption & " " & Chr(13) & " del Doctor: " & cboDoctor.Text & _
'    " ¿Confirma Copiar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'    sAction = "COPIAR"
'    dFechaCopy = MViewFecha.Value
'    nDoctorCopy = cboDoctor.ItemData(cboDoctor.ListIndex)
'    sNameDoctorCopy = cboDoctor.Text
    optNO.Enabled = True
    optSI.Enabled = True
    If grdGrilla.Rows > 1 Then
       If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
           txtBuscaCliente.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 11)
           'txtBuscaCliente_LostFocus
           txtCodigo.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
           txtBuscarCliDescri.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
           txtTelefono.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 3)
           txtOSocial.Text = BuscarOSocial(txtCodigo.Text) 'grdGrilla.TextMatrix(grdGrilla.RowSel, 5)
           
           'verifico si el paciente tiene mutual
           ' If Chk0(grdGrilla.TextMatrix(grdGrilla.RowSel, 13)) <> 1 Then 'si no tiene mutual el paciente
          If txtOSocial.Text = "" Then
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
          
           txtMotivo.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 6)
           txtDrSolicitante.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 7)
           BuscaDescriProx Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cboDesde
           BuscaDescriProx Right(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cbohasta
           
           mebHoraD.Text = Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5)
           mebHoraH.Text = Right(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5)
           
           'If Chk0(grdGrilla.TextMatrix(grdGrilla.RowSel, 13)) = 0 Then
           '    optSI.Enabled = False
           'End If

           If User = 1 Then
               txtimporte.Text = Valido_Importe(grdGrilla.TextMatrix(grdGrilla.RowSel, 14))
           Else
               txtimporte.Text = "0,00"
           End If
           cmdImpTurno.Enabled = True
           cmdProtocolos.Enabled = True
           cmdCortar.Enabled = True
           cmdCopiar.Enabled = True
       Else
           If txtBuscaCliente.Text <> "" Then
               MViewFecha.Value = Date
               txtBuscaCliente.Text = ""
               txtCodigo.Text = ""
               txtBuscarCliDescri.Text = ""
               txtTelefono.Text = ""
               txtOSocial.Text = ""
               txtMotivo.Text = ""
               cboDesde.ListIndex = -1
               cbohasta.ListIndex = -1
               mebHoraD.Text = ""
               mebHoraH.Text = ""
               
               txtimporte.Text = "0,00"
           End If
       End If
    End If

End Sub

Private Sub cmdCortar_Click()
    If MsgBox("Esta a punto de Cortar los " & lbldiaTurno.Caption & " " & Chr(13) & " del Doctor: " & cboDoctor.Text & _
    " ¿Confirma Cortar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    sAction = "CORTAR"
    dFechaCopy = MViewFecha.Value
    nDoctorCopy = cboDoctor.ItemData(cboDoctor.ListIndex)
    sNameDoctorCopy = cboDoctor.Text
End Sub

Private Sub cmdespera_Click()
    If grdGrilla.RowSel <> 0 Then
        'en espera
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = 2
        cambiocolor 2
        
        'Actualizo la Base de Datos
        sql = "UPDATE TURNOS SET "
        sql = sql & " TUR_ASISTIO =" & grdGrilla.TextMatrix(grdGrilla.RowSel, 10)
        sql = sql & " WHERE "
        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = #" & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        DBConn.Execute sql
    End If
End Sub

Private Sub cmdImpTurno_Click()
    If txtBuscaCliente.Text <> "" Then
        ImprimirTurno
    Else
        MsgBox "Seleccione un turno a imprimir", vbInformation, TIT_MSGBOX
    End If
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
'            If MsgBox("Hay Turnos previamente cargados en este dia que se eliminaran si realiza esta acción." & Chr(13) & _
'            " ¿Confirma eliminar estos Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
'
'            sql = "DELETE FROM TURNOS WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
'            sql = sql & " AND VEN_CODIGO =" & cboDoctor.ItemData(cboDoctor.ListIndex)
'            DBConn.Execute sql
'            LimpiarGrilla
'        End If
'
'         If MsgBox("Esta a punto de Pegar los " & sDiaTurno & " " & Chr(13) & "previamente cortados del Doctor: " & sNameDoctorCopy & _
'        " " & Chr(13) & "¿Confirma Pegar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
'                If MsgBox("Hay Turnos previamente cargados en este dia que se eliminaran si realiza esta acción." & Chr(13) & _
'                " ¿Confirma eliminar estos Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
'            " " & Chr(13) & "¿Confirma Pegar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
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
        vDNI = txtBuscaCliente.Text
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
    If grdGrilla.RowSel <> 0 Then
        'pendiente
        grdGrilla.TextMatrix(grdGrilla.RowSel, 10) = 0
        cambiocolor 0
        
        'Actualizo la Base de Datos
        sql = "UPDATE TURNOS SET "
        sql = sql & " TUR_ASISTIO =" & grdGrilla.TextMatrix(grdGrilla.RowSel, 10)
        sql = sql & " WHERE "
        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = #" & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#"
        sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 8))
        DBConn.Execute sql
    End If
End Sub

Private Sub cmdProtocolos_Click()
    fraprotocolos.Visible = True
    grdProtocolos.SetFocus
    grdProtocolos.Rows = 1
    cargo_protocolos
End Sub

Private Sub cmdQuitar_Click()
    'Controlar que se pueda eliminar el turno
    'Borrar de la Grilla
    'Borrar de la BD
    If txtCodigo.Text <> "" Then
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
            If MsgBox("�Confirma Eiminar el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
                
            sql = "DELETE FROM TURNOS WHERE"
            sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
            sql = sql & " AND TUR_HORAD = #" & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#"
            sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
            sql = sql & " AND CLI_CODIGO = " & grdGrilla.TextMatrix(grdGrilla.RowSel, 9)
            DBConn.Execute sql
            
            'ESTO LO HAGO PARA AUDITAR LO TURNOS BORRADOS
            'ver si hay algun turno borrado igual
            'sql = "SELECT * FROM DEL_TURNOS"
            'sql = sql & " WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
            'sql = sql & " AND TUR_HORAD = " & "#" & cboDesde.Text & "#"
            'sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
            'sql = sql & " AND CLI_CODIGO = " & XN(txtCodigo.Text)
            'Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
            'si no hay agrego
            'If Rec2.EOF = True Then
             '   sql = "INSERT INTO DEL_TURNOS"
              '  sql = sql & " (TUR_FECHA, TUR_HORAD,"
               ' sql = sql & " VEN_CODIGO,CLI_CODIGO,"
               ' sql = sql & " TUR_USER,TUR_FECBAJA)"
               ' sql = sql & " VALUES ("
               ' sql = sql & XDQ(MViewFecha.Value) & ",#"
               ' sql = sql & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#,"
               ' sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
               ' sql = sql & grdGrilla.TextMatrix(grdGrilla.RowSel, 8) & ","
               ' sql = sql & User & ","
                'sql = sql & XDQ(Date) & ")"
                'DBConn.Execute sql
            'Else
            'si hay no hago nada
            'End If
        
            If grdGrilla.Rows = 2 Then
                grdGrilla.Rows = 1
            Else
                grdGrilla.RemoveItem (grdGrilla.RowSel)
            End If
        
    '        'LIMPIO LA GRILLA
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 0) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = ""
    '        grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = ""
    '        LimpiarTurno
    '
    '        grdGrilla.row = grdGrilla.RowSel
    '        For J = 1 To grdGrilla.Cols - 1
    '            grdGrilla.Col = J
    '            grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR BLANCO
    '            grdGrilla.CellBackColor = &HC0FFC0       'ROSA
    '            grdGrilla.CellFontBold = True
    '        Next
        End If
    LimpiarTurno
Else
    MsgBox "Seleccione un turno", vbExclamation, TIT_MSGBOX
End If
End Sub

Private Sub cmdReport_Click()
    Dim ultimoimporte As Double
    Dim ultimoid As Integer
    'If txtCodCliente.Text = "" Or GrillaAplicar.Rows = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    'lblEstado.Caption = "Buscando Recibo..."

    sql = "DELETE FROM TMP_TURNOS"
    DBConn.Execute sql
    i = 1
    
    For i = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(i, 1) <> "" Then
            sql = "INSERT INTO TMP_TURNOS "
            sql = sql & " (TMP_ID,TMP_HORA,TMP_FECHA,TMP_DOCTOR,TMP_PACIENTE,TMP_EDAD,TMP_TELEFONO,TMP_CELULAR,TMP_OSOCIAL,TMP_MOTIVO,TMP_DRSOLICITA,TMP_IMPORTE)"
            sql = sql & " VALUES ( "
            sql = sql & i & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 0)) & ","
            sql = sql & XDQ(MViewFecha.Value) & ","
            sql = sql & XS(cboDoctor.Text) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 1)) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 2)) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 3)) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 4)) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 5)) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 6)) & ","
            sql = sql & XS(grdGrilla.TextMatrix(i, 7)) & ","
            sql = sql & XN(grdGrilla.TextMatrix(i, 14)) & ")"
            DBConn.Execute sql
        End If
    Next
    ultimoimporte = XN(grdGrilla.TextMatrix(grdGrilla.Rows - 1, 14))
    ultimoid = grdGrilla.Rows - 1
    
    'actualizo tabla para solucionar lo del ultimo registro
    sql = "UPDATE TMP_TURNOS"
    sql = sql & " SET TMP_IMPORTE=" & ultimoimporte
    sql = sql & " WHERE TMP_ID=" & ultimoid
    DBConn.Execute sql

    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR

    Rep.WindowTitle = "Listado de Turnos del dia"
    Rep.ReportFileName = DirReport & "rptTurnosDiario.rpt"
    Rep.Action = 1
'    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    Rep.SelectionFormula = ""
    
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
    txtfiltrop.Text = ""
End Sub
Private Function limpiar_protocolos()
    Dim i, J As Integer
    For i = 1 To grdProtocolos.Rows - 1
        grdProtocolos.TextMatrix(i, 3) = "NO"
        For J = 0 To grdProtocolos.Cols - 1
            grdProtocolos.row = i
            grdProtocolos.Col = J
            grdProtocolos.CellForeColor = &H80000008
            grdProtocolos.CellBackColor = &H80000005
            grdProtocolos.CellFontBold = False
        Next
    Next

End Function
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

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    
    MViewFecha.Value = Date
    fechaturno.Value = Date
    'MsgBox WeekdayName(7, False)
    'lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(Weekday(Date) - 1, False) & " " & day(Date) & " de " & MonthName(Month(Date), False) & " de " & Year(Date)
    configurodia Date
    configurogrilla
    LlenarComboDoctor
    LlenarComboHoras
    BuscarTurnos Date, cboDoctor.ItemData(cboDoctor.ListIndex)
    ActivoGrid = 1
    If User = 1 Then
        cmdAgregar.Enabled = True
        cmdAgregar.Enabled = True
        lblimporte.Visible = True
        txtimporte.Visible = True
        lbltotal.Visible = True
        txtTotal.Visible = True
    Else
        cmdAgregar.Enabled = False
        lblimporte.Visible = False
        txtimporte.Visible = False
        lbltotal.Visible = False
        txtTotal.Visible = False
    End If
    
    cargo_protocolos
End Sub
Private Sub LimpiarGrilla()
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.row = i
        For J = 1 To grdGrilla.Cols - 1
            grdGrilla.Col = J
            grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR BLANCO
            grdGrilla.CellBackColor = &HC0FFC0       'ROSA
            grdGrilla.CellFontBold = True
        Next
    Next
End Sub
Private Function cargo_protocolos()
    
    sql = "SELECT * FROM TIPO_IMAGEN"
    If txtfiltrop.Text <> "" Then
        sql = sql & " WHERE TIP_NOMBRE LIKE '%" & txtfiltrop.Text & "%'"
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
Private Sub BuscarTurnos(Fecha As Date, Doc As Integer)
    Dim foreColor As String
    Dim backColor As String
    Dim total As Double
    Dim a�os As Integer
    Dim edad As Integer
    sql = "SELECT T.*,V.VEN_NOMBRE,C.CLI_RAZSOC,C.CLI_NRODOC,C.CLI_TELEFONO,C.CLI_CELULAR,C.CLI_CUMPLE"
    sql = sql & " FROM TURNOS T, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND T.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND T.TUR_FECHA = " & XDQ(Fecha)
    sql = sql & " AND T.VEN_CODIGO = " & Doc
    sql = sql & " ORDER BY T.TUR_HORAD"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdGrilla.Rows = 1
    If rec.EOF = False Then
        i = 1
        Do While rec.EOF = False
            Select Case rec!TUR_ASISTIO
            Case 0
                backColor = &HC0C0FF
                foreColor = &H80000008
            Case 1
                backColor = &HC000&
                foreColor = &HFFFFFF
            Case 2
                backColor = &HFFFF&
                foreColor = &H80000008
            End Select
                    
            'calculo edad de paciente
            If Not (IsNull(rec!CLI_CUMPLE)) Then
                If rec.EOF = False Then
                    a�os = Year(Date) - Year(rec!CLI_CUMPLE)
                    If Month(Fecha) < Month(rec!CLI_CUMPLE) Then a�os = a�os - 1 'todavia no ha llegado el mes de su cumple
                    If Month(Now) = Month(rec!CLI_CUMPLE) And Day(Fecha) < Day(rec!CLI_CUMPLE) Then a�os = a�os - 1 'es el mes pero no ha llegado el dia de su cumple
                    edad = a�os
                End If
            Else
                edad = 0
            End If
    
            grdGrilla.AddItem Format(rec!TUR_HORAD, "hh:mm") & " a " & Format(rec!TUR_HORAH, "hh:mm") & Chr(9) & rec!CLI_RAZSOC & Chr(9) & edad & Chr(9) & ChkNull(rec!CLI_TELEFONO) & Chr(9) & ChkNull(rec!CLI_CELULAR) & Chr(9) & rec!TUR_OSOCIAL & Chr(9) & ChkNull(rec!TUR_MOTIVO) & Chr(9) & _
                                     ChkNull(rec!TUR_DRSOLICITA) & Chr(9) & rec!VEN_CODIGO & Chr(9) & rec!CLI_CODIGO & Chr(9) & rec!TUR_ASISTIO & Chr(9) & ChkNull(rec!CLI_NRODOC) & Chr(9) & ChkNull(rec!TUR_DESDE) & Chr(9) & rec!TUR_TIENEMUTUAL & Chr(9) & Format(Chk0(rec!TUR_IMPORTE), "#,##0.00")
                
            total = total + Chk0(rec!TUR_IMPORTE)
            'COLOR DE COLUMNA 1
            grdGrilla.Col = 0
            grdGrilla.row = i
            grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
            grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
            grdGrilla.CellFontBold = True
            
            
            'COLOR DE FILAS
            grdGrilla.row = i
            For J = 1 To grdGrilla.Cols - 1
                grdGrilla.Col = J
                grdGrilla.CellForeColor = foreColor       'FUENTE COLOR NEGRO
                grdGrilla.CellBackColor = backColor      'ROSA
                grdGrilla.CellFontBold = True
            Next
            
            i = i + 1
            rec.MoveNext
        Loop
    End If
    txtTotal.Text = total
    txtTotal.Text = Valido_Importe(txtTotal.Text)
    
    rec.Close
    grdGrilla.Col = 10
    If grdGrilla.row > 1 Then
        grdGrilla.row = 1
    End If
    'txtEdit.Visible = True
End Sub
Private Function cambiocolor(asistio As Integer)
    Dim foreColor As String
    Dim backColor As String
    
    Select Case asistio
    Case 0
        backColor = &HC0C0FF
        foreColor = &H80000008
    Case 1
        backColor = &HC000&
        foreColor = &HFFFFFF
    Case 2
        backColor = &HFFFF&
        foreColor = &H80000008
    End Select
    
    grdGrilla.row = grdGrilla.RowSel
    For J = 1 To grdGrilla.Cols - 1
        grdGrilla.Col = J
        grdGrilla.CellForeColor = foreColor       'FUENTE COLOR NEGRO
        grdGrilla.CellBackColor = backColor      'ROSA
        grdGrilla.CellFontBold = True
    Next
    
End Function


Private Sub LlenarComboDoctor()
     'BUSCO CODIGO DE DOCTOR POR NOMBRE DE USUARIO
    sql = "SELECT VEN_CODIGO FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO > 1 "
    If mNomUser = "A" Or mNomUser = "DIGOR" Or mNomUser = "PILI" Then
        sql = sql & " AND VEN_NOMBRE LIKE '" & "SILVANA" & "%'"
    Else
        sql = sql & " AND VEN_NOMBRE LIKE '" & mNomUser & "%'"
    End If
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        User = rec!VEN_CODIGO
    End If
    rec.Close

    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO > 1"
    sql = sql & " ORDER BY VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        'cboFactura1.AddItem "(Todas)"
        cboDoctor.AddItem ""
        Do While rec.EOF = False
            cboDoctor.AddItem rec!VEN_NOMBRE
            cboDoctor.ItemData(cboDoctor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        If mNomUser = "A" Or mNomUser = "DIGOR" Or mNomUser = "PILI" Then
            cboDoctor.ListIndex = 0
        Else
            Call BuscaCodigoProxItemData(XN(User), cboDoctor)
        End If
        
    End If
    rec.Close
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
    J = hDesde
    Do While cont < cItems
        minutos = 0
        For z = 0 To 11
            If cont < cItems Then
                If (minutos + 5) > 60 Then
                    'cboDesde.AddItem Format(J, "00") & ":" & Format(minutos, "00") & " a " & Format(J + 1, "00") & ":" & Format(0, "00")
                    Exit For
                Else
                    cboDesde.AddItem Format(J, "00") & ":" & Format(minutos, "00")
                    cboDesde.ItemData(cboDesde.NewIndex) = cont
                    cbohasta.AddItem Format(J, "00") & ":" & Format(minutos, "00")
                    cbohasta.ItemData(cbohasta.NewIndex) = cont
                End If
            End If
            cont = cont + 1
            minutos = minutos + 5
        Next
        J = J + 1
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
    grdGrilla.FormatString = "^Horas|<Paciente|<Edad|<Telefono|<Celular|<Obra Social|<Motivo|Dr Solicitante|>Doctor|>Cod Pac|>Asistio|DNI|TUR_DESDE|TieneMutual|Importe"
    grdGrilla.ColWidth(0) = 1400 'HORAS
    grdGrilla.ColWidth(1) = 2500 'PACIENTE
    grdGrilla.ColWidth(2) = 700 'EDAD
    grdGrilla.ColWidth(3) = 1300 'TELEFONO
    grdGrilla.ColWidth(4) = 1600 'CELULAR
    grdGrilla.ColWidth(5) = 1800 'O SOCIAL
    grdGrilla.ColWidth(6) = 2000 'MOTIVO
    grdGrilla.ColWidth(7) = 1500 'Dr Solicitante
    grdGrilla.ColWidth(8) = 0 'DOCTOR
    grdGrilla.ColWidth(9) = 0 'Codigo Paciente
    grdGrilla.ColWidth(10) = 0 'Asistio
    grdGrilla.ColWidth(11) = 0 'DNI
    grdGrilla.ColWidth(12) = 0 'TUR_DESDE
    grdGrilla.ColWidth(13) = 0 'TUR_TIENEMUTUAL
    'If User = 1 Then 'ESTA CONFIGURACION LA TOMA DEL INI
    If mNomUser = "DIGOR" Then 'ESTA CONFIGURACION LA TOMA DEL USUARIO LOGUEADO
        grdGrilla.ColWidth(14) = 1200 'Importe
    Else
        'oculto la columna de importe para los doctores
        grdGrilla.ColWidth(14) = 0 'Importe
    End If
    
    grdGrilla.Cols = 15
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
    grdGrilla.Rows = (hHasta - hDesde) * 12 + 1
    
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.Col = 0
        grdGrilla.row = i
        'grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        'grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR NEGRO
        grdGrilla.CellBackColor = &HC0C0FF      'ROSA
        grdGrilla.CellFontBold = True
        
    Next
    
    grdGrilla.Rows = 1
    
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
    grdProtocolos.Rows = 1
    grdProtocolos.HighLight = flexHighlightAlways
    
End Function

Private Sub GRDGrilla_DblClick()
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
        frmhistoriaclinica.Show vbModal
    End If
    

End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdQuitar_Click
    End If

End Sub

Private Sub grdProtocolos_DblClick()
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

Private Sub grdProtocolos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        grdProtocolos_DblClick
    End If
End Sub

Private Sub mebHoraD_LostFocus()
    If Mid(mebHoraD.Text, 1, 2) = "__" And Mid(mebHoraD.Text, 4, 2) <> "__" Then
        mebHoraD.Text = "00:" & Mid(mebHoraD.Text, 4, 2)
    End If
    If Mid(mebHoraD.Text, 4, 2) = "__" And Mid(mebHoraD.Text, 1, 2) <> "__" Then
        mebHoraD.Text = Mid(mebHoraD.Text, 1, 2) & ":00"
    End If
End Sub

Private Sub mebHoraH_LostFocus()
    If Mid(mebHoraH.Text, 1, 2) = "__" And Mid(mebHoraH.Text, 4, 2) <> "__" Then
        mebHoraH.Text = "00:" & Mid(mebHoraH.Text, 4, 2)
    End If
    If Mid(mebHoraH.Text, 4, 2) = "__" And Mid(mebHoraH.Text, 1, 2) <> "__" Then
        mebHoraH.Text = Mid(mebHoraH.Text, 1, 2) & ":00"
    End If
End Sub

Private Sub MViewFecha_DateClick(ByVal DateClicked As Date)
    'lbldiaTurno.Caption = "Turnos del dia " & MViewFecha.Value
    'lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(Weekday(MViewFecha.Value) - 1, False) & " " & day(MViewFecha.Value) & " de " & MonthName(Month(MViewFecha.Value), False) & " de " & Year(MViewFecha.Value)
    configurodia MViewFecha.Value
    fechaturno.Value = MViewFecha.Value
    LimpiarGrilla
    'LimpiarTurno
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
End Sub
Private Sub configurodia(Fecha As Date)
    Dim DIA As Integer
    DIA = Weekday(Fecha, vbMonday)
    lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(DIA, False) & " " & Day(Fecha) & " de " & MonthName(Month(Fecha), False) & " de " & Year(Fecha)
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
    If txtBuscaCliente.Text = "" Then
        txtBuscarCliDescri.Text = ""
        txtCodigo.Text = ""
        txtTelefono.Text = ""
        txtOSocial.Text = ""
    End If
    If Len(Trim(txtBuscaCliente.Text)) < 7 Then
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
    If txtBuscaCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO,OS_NUMERO"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            If Len(Trim(txtBuscaCliente.Text)) < 7 Then
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
            txtBuscarCliDescri.Text = rec!CLI_RAZSOC
            txtCodigo.Text = rec!CLI_CODIGO
            txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
            txtOSocial.Text = BuscarOSocial(rec!CLI_CODIGO)
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
    If txtBuscarCliDescri.Text = "" Then
        txtBuscaCliente.Text = ""
        txtCodigo.Text = ""
        txtTelefono.Text = ""
        txtOSocial.Text = ""
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
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO"
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
                'txtcodCli.Text = .ResultFields(2)
                'txtCodCli_LostFocus
            Else
                If .ResultFields(3) = "" Then
                    txtBuscaCliente.Text = .ResultFields(2)
                    txtCodigo.Text = .ResultFields(3)
                Else
                    txtBuscaCliente.Text = .ResultFields(3)
                    txtCodigo.Text = .ResultFields(3)
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

Private Sub txtfiltrop_GotFocus()
    seltxt
End Sub

Private Sub txtfiltrop_LostFocus()
    grdProtocolos.Rows = 1
    cargo_protocolos
End Sub

Private Sub txtimporte_GotFocus()
    seltxt
End Sub

Private Sub txtimporte_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtimporte, KeyAscii)
End Sub

Private Sub txtimporte_LostFocus()
    If txtimporte.Text <> "" Then
        txtimporte.Text = Valido_Importe(txtimporte)
    End If
End Sub

Private Sub txtMotivo_GotFocus()
    SelecTexto txtMotivo
End Sub


