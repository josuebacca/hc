VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmhistoriaclinica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia Clinica"
   ClientHeight    =   10905
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   16635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   16635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAgregarPedido 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   13080
      TabIndex        =   35
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   15240
      TabIndex        =   34
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminarImag 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   14160
      TabIndex        =   33
      Top             =   9960
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
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   26
         Tag             =   "Descripción"
         Top             =   360
         Width           =   4155
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
         TabIndex        =   25
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
         Left            =   9540
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "Descripción"
         Top             =   360
         Width           =   2715
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
         Left            =   3061
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
         Left            =   12483
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
         Left            =   8222
         TabIndex        =   28
         Top             =   360
         Width           =   1320
      End
   End
   Begin TabDlg.SSTab tabhc 
      Height          =   9615
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   120
         TabIndex        =   54
         Top             =   7560
         Width           =   8055
         Begin VB.CommandButton cmdLabora 
            Caption         =   "&Laboratorio"
            Height          =   855
            Left            =   4800
            TabIndex        =   59
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdEcogra 
            Caption         =   "&Ecografias"
            Height          =   855
            Left            =   3600
            TabIndex        =   58
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdGineco 
            Caption         =   "&Ginecologia"
            Height          =   855
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdImag 
            Caption         =   "&Imagenes"
            Height          =   855
            Left            =   2400
            TabIndex        =   56
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdPedidos 
            Caption         =   "&Pedidos"
            Height          =   855
            Left            =   1200
            TabIndex        =   55
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
         Top             =   480
         Width           =   8055
         Begin VB.TextBox txtCodCC 
            Height          =   375
            Left            =   5040
            TabIndex        =   63
            Top             =   120
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker FechaProx 
            Height          =   375
            Left            =   2040
            TabIndex        =   62
            Top             =   6360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   105578497
            CurrentDate     =   43205
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
            Left            =   4515
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   600
            Width           =   3180
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   5640
            TabIndex        =   51
            Top             =   6360
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6720
            TabIndex        =   50
            Top             =   6360
            Width           =   1095
         End
         Begin VB.TextBox txtIndicaciones 
            Height          =   4635
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   1560
            Width           =   6375
         End
         Begin VB.TextBox txtMotivo 
            Height          =   315
            Left            =   1320
            TabIndex        =   47
            Top             =   1080
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
            Format          =   105578497
            CurrentDate     =   41098
         End
         Begin VB.Label Label14 
            Caption         =   "Label14"
            Height          =   255
            Left            =   2760
            TabIndex        =   65
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Próxima Consulta:"
            Height          =   375
            Left            =   360
            TabIndex        =   64
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   3840
            TabIndex        =   53
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Indicaciones:"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   1560
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   690
            TabIndex        =   46
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Motivo:"
            Height          =   195
            Left            =   660
            TabIndex        =   44
            Top             =   1080
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
         Top             =   480
         Width           =   8055
         Begin VB.CommandButton cmdFiltro 
            Caption         =   "Filtro"
            Height          =   735
            Left            =   6360
            TabIndex        =   61
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
            Width           =   3850
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
            Format          =   105578497
            CurrentDate     =   41098
         End
         Begin MSComCtl2.DTPicker FechaHasta 
            Height          =   315
            Left            =   4455
            TabIndex        =   40
            Top             =   735
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   105578497
            CurrentDate     =   41098
         End
         Begin MSFlexGridLib.MSFlexGrid grdConsultas 
            Height          =   6870
            Left            =   120
            TabIndex        =   60
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
            Left            =   960
            TabIndex        =   42
            Top             =   795
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta:"
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   41
            Top             =   795
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Doctor:"
            Height          =   195
            Left            =   960
            TabIndex        =   38
            Top             =   435
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdImprimirEstGine 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -57480
         TabIndex        =   23
         Top             =   8760
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEstGine 
         Caption         =   "Eliminar Estudio"
         Height          =   375
         Left            =   -59400
         TabIndex        =   22
         Top             =   8760
         Width           =   1575
      End
      Begin VB.CommandButton cmdAgregarEstGine 
         Caption         =   "Agregar Estudio"
         Height          =   375
         Left            =   -61080
         TabIndex        =   21
         Top             =   8760
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estudios"
         Height          =   7095
         Left            =   -74880
         TabIndex        =   19
         Top             =   1680
         Width           =   19095
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   6615
            Left            =   120
            TabIndex        =   20
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
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton optFechaGine 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   -68400
         TabIndex        =   17
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optTipoEst 
         Caption         =   "Tipo Estudio"
         Height          =   375
         Left            =   -73320
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdVerEstudio 
         Caption         =   "Ver Estudio(VA ACA?)"
         Height          =   375
         Left            =   -60360
         TabIndex        =   14
         Top             =   8760
         Width           =   1335
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -61800
         TabIndex        =   13
         Top             =   8760
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregarEco 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -58800
         TabIndex        =   12
         Top             =   8760
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEco 
         Caption         =   "Elinimar"
         Height          =   375
         Left            =   -57000
         TabIndex        =   11
         Top             =   8760
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ecografías"
         Height          =   8055
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   16215
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   6255
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   11033
            _Version        =   393216
            Rows            =   4
            Cols            =   5
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   11880
            TabIndex        =   9
            Text            =   "Doctor"
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   8760
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox cboEspecialidad 
            Height          =   315
            Left            =   4800
            TabIndex        =   7
            Text            =   "Especialidad"
            Top             =   840
            Width           =   2655
         End
         Begin VB.OptionButton optEmpleado 
            Caption         =   "Doctor"
            Height          =   255
            Left            =   11880
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton optFecha 
            Caption         =   "Fecha"
            Height          =   375
            Left            =   8760
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optEspecialidad 
            Caption         =   "Especialidad"
            Height          =   255
            Left            =   4800
            TabIndex        =   4
            Top             =   360
            Width           =   2175
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
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmhistoriaclinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
        MsgBox "No ha ingresado la hora de comienzo del Turno", vbCritical, TIT_MSGBOX
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



Private Sub cmdAceptar_Click()
Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    Dim i As Integer
    Dim num As Integer

    
    'Validar los campos requeridos
    If validarcclinico = False Then Exit Sub
    If MsgBox("¿Desea cargar la consulta?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'agregar teniendo en cuentas loc combos de horas
    'On Error GoTo HayErrorTurno
    
    'grdGrilla.HighLight = flexHighlightAlways
    
    'nFilaD = cboDesde.ListIndex
    'nFilaH = cbohasta.ListIndex
    i = 0
    sql = "SELECT TOP 1 CCL_NUMERO FROM CCLINICO"
    sql = sql & " ORDER BY CCL_NUMERO DESC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    num = rec
    Label14.Caption = num
    rec.Close
    
    'sHoraDAux = cboDesde.Text
    'For i = 1 To nFilaH - nFilaD
        'DBConn.BeginTrans
        
        'sHoraD = cboDesde.Text
        'sHoraD = Mid(sHoraD, 1, 1)
        
        'If sHoraD = "0" Then
         '   sHoraD = Mid(cboDesde.Text, 2, 4)
        'Else
        '   sHoraD = Trim(cboDesde.Text)
        'End If
        
        'ACA TENGO QUE HACER UN CONTROL POR CLAVES PRIMARIAS
        'sql = "SELECT * FROM CCLINICO"
        'sql = sql & " WHERE CCL_FECHA = " & XDQ(Fecha.Value)
        'sql = sql & " AND TUR_HORAD = #" & sHoraD & "#"
        'sql = sql & " AND VEN_CODIGO = " & cboDocCon.ItemData(cboDocCon.ListIndex)
        'sql = sql & " AND CLI_CODIGO = " & XN(txtcodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        'If Not rec.EOF = False Then
            sql = "INSERT INTO CCLINICO"
            sql = sql & " (CCL_NUMERO,CCL_FECHA,"
            sql = sql & " CLI_CODIGO,VEN_CODIGO,CCL_INDICA"
        If FechaProx.Value <> False Then
            sql = sql & "," & "CCL_FECPC"
        End If
             sql = sql & ")"
        
            'If User <> 99 Then
                'sql = sql & " TUR_USER, "
            'End If
            'sql = sql & " )"
            sql = sql & " VALUES ("
            sql = sql & XN(txtCodCC.Text) & ","
            sql = sql & XDQ(Fecha.Value) & ","
            sql = sql & XN(txtBuscaCliente.Text) & ","
            sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex) & ","
            sql = sql & XS(txtIndicaciones.Text)
            If FechaProx.Value <> False Then
                sql = sql & "," & XDQ(FechaProx.Value)
            End If
            sql = sql & ")"
            'sql = sql & XDQ(Fecha.Value) & ",#"
            'sql = sql & Left(Trim(grdGrilla.TextMatrix(i + nFilaD, 0)), 5) & "#,#"
            'sql = sql & Right(Trim(grdGrilla.TextMatrix(i + nFilaD, 0)), 5) & "#,"
            'sql = sql & cboDocCon.ItemData(cboDocCon.ListIndex) & ","
            'sql = sql & XN(txtCodigo) & ","
            'sql = sql & XS(txtMotivo) & ","
            'sql = sql & 0 & ","
            'sql = sql & XS(txtOSocial.Text) & ","
            'If User <> 99 Then
                'sql = sql & User & ","
            'End If
            'sql = sql & XDQ(Fecha) & ")"
            
            'ACTUALIZO LA GRILLA
            'grdGrilla.row = nFilaD + i
            'For J = 1 To grdGrilla.Cols - 1
                'grdGrilla.Col = J
                'grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR NEGRO
                'grdGrilla.CellBackColor = &HC0C0FF          'ROSA
                'grdGrilla.CellFontBold = True
            'Next
            'grdGrilla.TextMatrix(i + nFilaD, 1) = txtBuscarCliDescri.Text
            'grdGrilla.TextMatrix(i + nFilaD, 2) = txtTelefono.Text
            'grdGrilla.TextMatrix(i + nFilaD, 3) = txtOSocial.Text
            'grdGrilla.TextMatrix(i + nFilaD, 1) = Fecha.Value
            'grdGrilla.TextMatrix(i + nFilaD, 2) = txtMotivo.Text
            'grdGrilla.TextMatrix(i + nFilaD, 3) = cboDocCon.ItemData(cboDocCon.ListIndex)
            'grdGrilla.TextMatrix(i + nFilaD, 6) = txtCodigo.Text
            'grdGrilla.TextMatrix(i + nFilaD, 7) = 0
            'grdGrilla.TextMatrix(i + nFilaD, 8) = txtBuscaCliente.Text
            
        'Else
            'MsgBox "Ya hay un turno para ese horario", vbExclamation, TIT_MSGBOX
            'cboDesde.Text = sHoraDAux
'            rec.Close
'            DBConn.Execute sql
'            DBConn.CommitTrans
'            Exit Sub
            'If MsgBox("Ya hay un turno para ese horario ¿Confirma la Modificación del Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
              '  rec.Close
               ' Exit Sub
        'End If
            ' aca hago el update
            'sql = "UPDATE CCLINICO SET "
            'sql = sql & " CCL_FECHA = " & XDQ(Fecha.Value)
            'sql = sql & " ,CLI_CODIGO =" & XN(txtBuscaCliente.Text) 'CAMBIAR CUANDO CARGUEMOS DNI
            'sql = sql & " AND VEN_CODIGO = " & cboDocCon.ItemData(cboDocCon.ListIndex)
            'sql = sql & " ,CCL_INDICA = " & XS(txtIndicaciones.Text)
            'sql = sql & " ,CCL_MOTIVO =" & XS(txtMotivo.Text)
            'sql = sql & " ,TUR_OSOCIAL =" & XS(txtOSocial.Text)
            'If User <> 99 Then
               ' sql = sql & " ,TUR_USER =" & User
            'End If
            'sql = sql & " WHERE "
            'PORQUE?
            
            'sql = sql & " AND TUR_HORAD = #" & cboDesde.Text & "#"
            
            
            'grdGrilla.TextMatrix(i + nFilaD, 1) = txtBuscarCliDescri.Text
            'grdGrilla.TextMatrix(i + nFilaD, 2) = txtTelefono.Text
            'grdGrilla.TextMatrix(i + nFilaD, 3) = txtOSocial.Text
            'grdGrilla.TextMatrix(i + nFilaD, 1) = Fecha.Value
            'grdGrilla.TextMatrix(i + nFilaD, 2) = txtMotivo.Text
            'grdGrilla.TextMatrix(i + nFilaD, 3) = cboDocCon.ItemData(cboDocCon.ListIndex)
            'grdGrilla.TextMatrix(i + nFilaD, 6) = txtCodigo.Text
            'grdGrilla.TextMatrix(i + nFilaD, 7) = 0
            'grdGrilla.TextMatrix(i + nFilaD, 8) = txtBuscaCliente.Text
            'BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
        'End If

        
        rec.Close
        DBConn.Execute sql
        'DBConn.CommitTrans
        
        'cboDesde.ListIndex = cboDesde.ListIndex + 1
    'Next
    'cboDesde.Text = sHoraDAux
    'If MsgBox("¿Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'ImprimirTurno
    
    'LimpiarTurno
            
    Exit Sub
    
'HayErrorTurno:
    'Screen.MousePointer = vbNormal
    'If rec.State = 1 Then rec.Close
    'If Rec1.State = 1 Then Rec1.Close
    'DBConn.RollbackTrans
    'MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
    'agregar columnas en la grilla, para guardar el codigo de doctor, paciente
End Sub

Private Sub cmdEcogra_Click()
    tabhc.Tab = 1
End Sub

Private Sub cmdFiltro_Click()
Dim sColor As String
    Dim USUARIO As String
    sql = "SELECT CC.*,V.VEN_NOMBRE,C.CLI_RAZSOC"
    sql = sql & " FROM CCLINICO CC, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE CC.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND CC.VEN_CODIGO = V.VEN_CODIGO"
    If txtBuscaCliente.Text <> "" Then
        sql = sql & " AND CC.CLI_CODIGO = " & XN(txtCodigo)
    End If
    If cboDocAnt.ListIndex <> -1 Then
        sql = sql & " AND CC.VEN_CODIGO = " & cboDocAnt.ItemData(cboDocAnt.ListIndex)
    End If
    If FechaDesde.Value <> "" Then sql = sql & " AND CC.CCL_FECHA>=" & XDQ(FechaDesde.Value)
    If FechaHasta.Value <> "" Then sql = sql & " AND CC.CCL_FECHA<=" & XDQ(FechaHasta.Value)
    'If cboDesde.ListIndex <> -1 Then
        'sql = sql & " AND T.TUR_HORAD = " & cboDesde.ItemData(cboDesde.ListIndex)
    'End If
    sql = sql & " ORDER BY CCL_FECHA DESC"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    grdConsultas.Rows = 1
    i = 1
    If rec.EOF = False Then
        Do While rec.EOF = False
          '  If Chk0(rec!TUR_USER) <> 0 Then
                'USUARIO = BuscarUser(ChkNull(rec!TUR_USER))
           ' End If
            grdConsultas.AddItem rec!CCL_FECHA & Chr(9) & rec!VEN_NOMBRE & Chr(9) & _
                                     Chr(9) & rec!CCL_INDICA & Chr(9)
                 
            'Select Case rec!TUR_ASISTIO
            'Case 0
            '    sColor = &H80000008
            'Case 1
             '   sColor = &HC000&
            'Case 2
             '   sColor = &HFF&
            'End Select
            grdConsultas.row = i
            For J = 0 To grdConsultas.Cols - 1
                grdConsultas.Col = J
                'grdConsultas.CellForeColor = sColor        'FUENTE COLOR negro
                grdConsultas.CellBackColor = &HC0C0FF          'ROSA
                grdConsultas.CellFontBold = True
            Next
            i = i + 1
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub cmdGineco_Click()
    tabhc.Tab = 3
End Sub

Private Sub cmdLabora_Click()
    tabhc.Tab = 2
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
    preparogrillas
    cargocombos
End Sub
Private Function preparogrillas()
    ' Grilla de Curso Clinico - Consulta de Historia Clinica
    grdConsultas.FormatString = "Fecha|Doctor|Indicaciones"
    grdConsultas.ColWidth(0) = 1500  'Fecha
    grdConsultas.ColWidth(1) = 2500 'Doctor
    grdConsultas.ColWidth(2) = 0 'Inidicaciones
    grdConsultas.Rows = 1
    grdConsultas.BorderStyle = flexBorderNone
    grdConsultas.row = 0
    For i = 0 To 2
        grdConsultas.Col = i
        grdConsultas.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdConsultas.CellBackColor = &H808080    'GRIS OSCURO
        grdConsultas.CellFontBold = True
    Next
    grdConsultas.HighLight = flexHighlightNever
    
End Function
Private Function cargocombos()
    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO =1"
    sql = sql & " ORDER BY VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboDocCon.AddItem rec!VEN_NOMBRE
            cboDocCon.ItemData(cboDocCon.NewIndex) = rec!VEN_CODIGO
            
            cboDocAnt.AddItem rec!VEN_NOMBRE
            cboDocAnt.ItemData(cboDocAnt.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        'If User <> 99 Then
        '    Call BuscaCodigoProxItemData(XN(User), cboDoctor)
        'Else
        '    cboDocCon.ListIndex = 0
        'End If
    End If
    rec.Close
End Function

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
        'ActivoGrid = 0
    End If
End Sub

Private Sub txtBuscaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtBuscaCliente_LostFocus()
    If txtBuscaCliente.Text <> "" Then
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
                'txtcodCli.Text = .ResultFields(2)
                'txtCodCli_LostFocus
            Else
                If .ResultFields(3) = "" Then
                    txtBuscaCliente.Text = .ResultFields(2)
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
