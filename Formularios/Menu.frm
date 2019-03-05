VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Menu 
   BackColor       =   &H8000000C&
   ClientHeight    =   5655
   ClientLeft      =   105
   ClientTop       =   2085
   ClientWidth     =   10755
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Picture         =   "Menu.frx":08CA
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrPrincipal 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Turnos"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Pacientes"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Tratamientos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Medicamentos"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cumpleaños"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Control"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   0
      Top             =   4920
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10755
      TabIndex        =   2
      Top             =   420
      Width           =   10755
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   3000
         ScaleHeight     =   495
         ScaleWidth      =   9375
         TabIndex        =   7
         Top             =   120
         Width           =   9375
      End
      Begin VB.CommandButton cmdCommunicator 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   12720
         Picture         =   "Menu.frx":9252
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdActLetrero 
         Height          =   495
         Left            =   2520
         Picture         =   "Menu.frx":AB70
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Actualizar Letrero"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtPaciente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Paciente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   0
         Width           =   1470
      End
   End
   Begin ComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   556
      SimpleText      =   "Listo."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   6526
            MinWidth        =   6526
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7673
            MinWidth        =   7673
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "NÚM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "MAYÚS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "19:34"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "04/03/2019"
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":B43A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":B614
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":B92E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":BC48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":BF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":C27C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":C596
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivos 
      Caption         =   "Archivos"
      Begin VB.Menu mnuconectar 
         Caption         =   "Conectar"
      End
      Begin VB.Menu mnudesconectar 
         Caption         =   "Desconectar"
      End
      Begin VB.Menu mnuRayaConectar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsuario 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuPermisos 
         Caption         =   "Permisos"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "Parametros"
      End
      Begin VB.Menu mnuRayaSalir 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArcSal 
         Caption         =   "Sali&r"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEstablecer 
      Caption         =   "Establecer"
      Begin VB.Menu mnuPacientes 
         Caption         =   "Pacientes"
      End
      Begin VB.Menu mnurayapacientes 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPersonal 
         Caption         =   "Personal"
      End
      Begin VB.Menu mnuProfesiones 
         Caption         =   "Profesiones"
      End
      Begin VB.Menu mnuRayaProf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProtocolos 
         Caption         =   "Protocolos"
      End
      Begin VB.Menu mnuTratamientos 
         Caption         =   "Tratamientos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMedicamentos 
         Caption         =   "Medicamentos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLabDentales 
         Caption         =   "Laboratorios Dentales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLabClinicos 
         Caption         =   "Laboratorios Clinicos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGrupos 
         Caption         =   "Grupos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRaya11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuABMLocalidades 
         Caption         =   "Localidades"
      End
      Begin VB.Menu mnuABMProvincias 
         Caption         =   "Provincias"
      End
      Begin VB.Menu mnuObrasSociales 
         Caption         =   "Obras Sociales"
      End
      Begin VB.Menu Motivo 
         Caption         =   "Motivos de Turnos"
      End
   End
   Begin VB.Menu mnuTurnos 
      Caption         =   "Turnos"
      Begin VB.Menu mnuAsignarTurnos 
         Caption         =   "Asignar Turnos"
      End
      Begin VB.Menu mnurayaatenciones 
         Caption         =   "-"
      End
      Begin VB.Menu mnuatenciones 
         Caption         =   "Listado de Atenciones Medicas"
      End
   End
   Begin VB.Menu mnuDatosPacientes 
      Caption         =   "&Historias Clinicas"
      Begin VB.Menu mnuArchivoActualizaciones 
         Caption         =   "Ingreso de Datos Pacientes"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnuCumpleaños 
         Caption         =   "Cumpleaños"
      End
      Begin VB.Menu mnuControlVisitas 
         Caption         =   "Control de Visitas"
      End
      Begin VB.Menu mnuRayaUtil1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNotas 
         Caption         =   "Bloc de Notas"
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "Calculadora"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuReporteTurnos 
         Caption         =   "Turnos"
      End
      Begin VB.Menu mnuRepPacientes 
         Caption         =   "Pacientes"
      End
   End
   Begin VB.Menu mnuMantenimiento 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnuBkpArchivos 
         Caption         =   "Backup de Archivos"
      End
      Begin VB.Menu mnuRestArchivos 
         Caption         =   "Restaurar Archivos"
      End
      Begin VB.Menu mnuAAcerca 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu ContextBaseABM 
      Caption         =   "ContextBaseABM"
      Visible         =   0   'False
      Begin VB.Menu mnuContextABM 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Editar"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Eliminar"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Refrescar"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Buscar"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Imprimir"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextABM 
         Caption         =   "Ver Datos"
         Index           =   9
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim TituloPrincipal As String

Private Declare Function ShellAbout Lib "shell32.dll" Alias _
"ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Dim Letrero As String
Dim i As Integer


Private Sub cmdActLetrero_Click()
    Letrero = ""
    configuroLetrero Date
End Sub

Private Sub cmdCommunicator_Click()
    frmCommunicator.Show
End Sub

Private Sub MDIForm_Load()
    'If Dir("c:\windows\cpce.ini") = "" Then
        'Menu.Picture = LoadPicture(App.Path & "\fotos\Demaría.bmp")
    'End If
    
    TituloPrincipal = TIT_MSGBOX '"Sistema de Gestión y Administración"
    Me.Caption = TituloPrincipal
    'inicio
    Me.Show
    FrmInicio.Show vbModal
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"

    sql = "SELECT RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'If rec.EOF = False Then
    '    lblpertenece.Caption = rec!RAZ_SOCIAL
    'End If
    rec.Close
    'MenuCore.Show
    'Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    'Me.Caption = TituloPrincipal '& "    V. " & App.Major & "." & App.Minor & "." & App.Revision & "          - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    'Menu.mnuConectar.Enabled = False
    'MenuCore.Show
    
    'buscarcumples Date
    buscarRecordatorio Date
    configuroLetrero Date
End Sub

Private Function inicio()
 Dim TxtUsuario As String
 Dim TxtClave As String
 Set rec = New ADODB.Recordset
    TxtUsuario = "a"
    TxtClave = "az"
    
    mNomUser = Trim(TxtUsuario)
    
    
    Conexion TxtUsuario, TxtClave
        
    
    If Not CONECCION Then
        If Err.Description <> "" Then
            MsgBox Err.Description
        End If
            
        'CUANTAS_VECES = CUANTAS_VECES + 1
        'If CUANTAS_VECES = 4 Then
        '    End
        'End If
        'txtusuario.SetFocus
        'Exit Sub
    End If


    sql = "SELECT * FROM USUARIO WHERE " & _
          "USU_NOMBRE = '" & Trim(TxtUsuario) & "' AND " & _
           "USU_CLAVE = '" & Trim(TxtClave) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

        mNomUser = Trim(TxtUsuario)
        mPassword = Trim(TxtClave)
        
        'BUSCO SUCURSAL---
        BuscoNroSucursal
        '-----------------
        'Unload Me
        'Set FrmInicio = Nothing
    'End If
End Function
Public Function Conexion(TxtUsuario As String, TxtClave As String)
Dim DSN_DEF As String
    Screen.MousePointer = vbHourglass
    CONECCION = False

    On Error GoTo ErrorIni
    LeoIni
    
    On Error GoTo ErrorTrans
    'ME CONECTO !
    Set DBConn = New ADODB.Connection
'    mNomUser = TxtUsuario.Text
'    mPassword = TxtClave.Text
'    DSN_DEF = "VAUDAGNA"
    'DBConn.ConnectionString = "ODBC;DATABASE=;UID=" & mNomUser & ";PWD=" & mPassword & ";DSN=" & DSN_DEF
    'DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CORE"
    'DBConn.ConnectionString = "driver={SQL Server}; server=DANIEL;database=VAUDAGNA"

    'DBConn.ConnectionTimeout = 30       'Default msado10.hlp => 15
    
    ' lee un path
    'DBConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & SERVIDOR & ";"
    DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    DBConn.Open
    
    
    'DBConn.CommandTimeout = 0          'Default msado10.hlp => 30
    'DBConn.Open , TxtUsuario, TxtClave
    'DBConn.Open DBConn.ConnectionString, TxtUsuario, TxtClave
    
       
    If DBConn.State = adStateOpen Then CONECCION = True
        
    PERMISOS mNomUser
    Screen.MousePointer = vbNormal
    Exit Function
    
ErrorTrans:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo establecer la conección a la Base de Datos." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
        Exit Function
ErrorIni:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo leer el archivo de configuración del sistema." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
End Function
Private Sub MDIForm_Unload(Cancel As Integer)
    Call mnuArcSal_Click
End Sub

Private Sub mnuAAcerca_Click()
    Call ShellAbout(Me.hWnd, "Sistema de Historias Clinicas", "Copyright 2009, Daniel Quiroga", Me.Icon)
End Sub



Private Sub mnuABMPais_Click()

End Sub

Private Sub mnuABMLocalidades_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMLocalidad = New CListaBaseABM
    
    With vABMLocalidad
        .Caption = "Actualización de Localidades"
        .sql = "SELECT L.LOC_DESCRI, L.LOC_CODIGO, P.PRO_DESCRI, P.PRO_CODIGO, PA.PAI_DESCRI, P.PAI_CODIGO" & _
               " FROM LOCALIDAD L, PROVINCIA P, PAIS PA" & _
               " WHERE P.PAI_CODIGO=PA.PAI_CODIGO" & _
               " AND L.PAI_CODIGO=PA.PAI_CODIGO" & _
               " AND L.PRO_CODIGO=P.PRO_CODIGO"
        .HeaderSQL = "Descripción, Código, Provincia, Código ,País, Código"
        .FieldID = "LOC_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormLocalidad
        Set .FormDatos = ABMLocalidad
    End With
    
    Set auxDllActiva = vABMLocalidad
    
    vABMLocalidad.Show
End Sub

Private Sub mnuABMProvincias_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMProvincia = New CListaBaseABM
    
    With vABMProvincia
        .Caption = "Actualización de Provincias"
        .sql = "SELECT P.PRO_DESCRI, P.PRO_CODIGO, PA.PAI_DESCRI, P.PAI_CODIGO" & _
               " FROM PROVINCIA P, PAIS PA" & _
               " WHERE P.PAI_CODIGO=PA.PAI_CODIGO"
        .HeaderSQL = "Descripción, Código, País, Código"
        .FieldID = "PRO_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormProvincia
        Set .FormDatos = ABMProvincia
    End With
    
    Set auxDllActiva = vABMProvincia
    
    vABMProvincia.Show
End Sub

Private Sub mnuArcSal_Click()
    On Error Resume Next
    'verifico si la conexión esta abierta antes de salir
    'If Me.mnuConexion.Enabled = False Then
    DBConn.CloseConnection
    Set DBConn = Nothing
    'End If
    Set Menu = Nothing
    End
End Sub

Private Sub mnuAsignarTurnos_Click()
    frmTurnos.Show
End Sub

Private Sub mnuatenciones_Click()
    frmListadoCantVendidasVendedor.Show
End Sub

Private Sub mnuBkpArchivos_Click()
    With frmRestaurarBD
        .Caption = "Backup de Archivos"
        .optCopiarA.Value = True
        .Label1 = "Guardar Backup en:"
        .Show
    End With
End Sub

Private Sub mnuconectar_Click()
    FrmInicio.Show vbModal
    'inicio
    Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    Me.mnuconectar.Enabled = False
End Sub

Public Sub mnuContextABM_Click(Index As Integer)

Dim auxListView As ListView
Dim auxModo As Integer
    
    auxModo = 0
    Select Case Index
        Case 0 'nuevo
            auxModo = 1
        Case 1 'editar
            auxModo = 2
        Case 2 'eliminar
            auxModo = 4
        Case 9 ' ver datos
            auxModo = 3
        'Case 7 ' imprimir
        '   auxModo = 7
    End Select
    
    If auxModo > 0 Then
        Set auxListView = auxDllActiva.FormBase.lstvLista
        auxDllActiva.FormDatos.SetWindow auxDllActiva.FormBase, auxDllActiva.sql, auxModo, auxListView, auxDllActiva.FieldID
        auxDllActiva.FormDatos.Show vbModal
    Else
        'si es una acción de edición de datos
        Select Case Index
            Case 4 'refresh
                Screen.MousePointer = vbHourglass
                With auxDllActiva
                    Set auxListView = .FormBase.lstvLista
                    CargarListView .FormBase, auxListView, .sql, .FieldID, .HeaderSQL, .FormBase.ImgLstLista
                    .FormBase.sBarEstado.Panels(1).Text = auxListView.ListItems.Count & " Registro(s)"
                End With
                Screen.MousePointer = vbDefault

            Case 5 'refresh
                'auxDllActiva.FormBase.txtBusqueda.Text = ""
                'auxDllActiva.FormBase.fraFiltro.Visible = True
                'auxDllActiva.FormBase.txtBusqueda.SetFocus
                With auxDllActiva
                    If .Caption = "Actualización de Productos" Then
                        frmFiltroProducto.Show
                    Else
                        frmFiltro.Show
                    End If
                        
                End With

            Case 6 'Buscar
                    auxDllActiva.Find
                
            Case 7 'imprimir
                Select Case mQuienLlamo
                    Case "ABMProducto"
                        frmImprimeProducto.Show vbModal
                    Case Else
                        On Error GoTo ErrorReport
                        auxDllActiva.FormBase.rptListado.Action = 1
                        On Error GoTo 0
                End Select
        End Select
    End If
    Exit Sub
    
ErrorReport:
    
    Beep
    MsgBox "Error " & Err.Number & Chr(13) & Err.Description, vbCritical + vbOKOnly, App.Title
    
End Sub
Private Sub mnuArchivoActualizaciones_Click()
    frmhistoriaclinica.Show
End Sub
Private Sub mnuCalculadora_Click()
    On Error Resume Next
    Shell "C:\WINDOWS\system32\calc.exe", vbNormalFocus
    'Form1.Show vbModal
End Sub
Private Sub mnuCalendario_Click()
    frmCalendario.Show
    
End Sub
Private Sub mnuWordPad_Click()
    On Error Resume Next
    Shell "C:\Archivos de programa\Windows NT\Accesorios\wordpad.exe", vbNormalFocus
    'Form1.Show vbModal
End Sub


Private Sub mnuControlVisitas_Click()
    frmControlVisitas.Show
End Sub

Private Sub mnuCumpleaños_Click()
    frmCumpleaños.Show
End Sub

Private Sub mnudesconectar_Click()
    If DBConn.State = adStateOpen Then
        DBConn.Close
        
        DeshabilitarMenu Me
        
        Me.mnuArchivos.Enabled = True
        Me.mnuconectar.Enabled = True
        Me.mnuArcSal.Enabled = True
        Me.mnudesconectar.Enabled = False
        
        Me.Caption = TituloPrincipal & " - (No conectado)"
        FrmInicio.Show vbModal
        
        Me.Caption = TituloPrincipal & " - (Usuario " & UCase(mNomUser) & " conectado a " & UCase(SERVIDOR) & " - " & UCase(BASEDATO) & ")"
    End If
End Sub

Private Sub mnuGrupos_Click()
Dim cSQL As String
    
    mOrigen = True
        
    Set vABMGrupos = New CListaBaseABM
    
    With vABMGrupos
        .Caption = "Actualizacion de Grupos"
        .sql = "SELECT GRU_DESCRI, GRU_CODIGO " & _
               " FROM GRUPOS "
        .HeaderSQL = "Descripcion, Código"
        .FieldID = "GRU_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormGrupos
        Set .FormDatos = ABMGrupos
    End With
    
    Set auxDllActiva = vABMGrupos
    
    vABMGrupos.Show

End Sub

Private Sub mnuLabClinicos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMLabClinicos = New CListaBaseABM
    
    With vABMLabClinicos
        .Caption = "Actualizacion de Laboratorios Clinicos"
        .sql = "SELECT C.LAC_NOMBRE, C.LAC_CODIGO, C.LAC_DOMICI,L.LOC_DESCRI," & _
               "  C.LAC_TELEFONO  " & _
               " FROM LAB_CLINICOS C, LOCALIDAD L " & _
               " WHERE C.LOC_CODIGO = L.LOC_CODIGO "
        .HeaderSQL = "Nombre, Código,Domicilio, Localidad, Teléfono"
        .FieldID = "LAC_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormLabClinicos
        Set .FormDatos = ABMLabClinicos
    End With
    
    Set auxDllActiva = vABMLabClinicos
    
    vABMLabClinicos.Show

End Sub

Private Sub mnuLabDentales_Click()
'    Dim cSQL As String
'
'    mOrigen = True
'
'    Set vABMLabDentales = New CListaBaseABM
'
'    With vABMLabDentales
'        .Caption = "Actualizacion de Laboratorios Dentales"
'        .sql = "SELECT C.LAD_NOMBRE, C.LAD_CODIGO, C.LAD_DOMICI,L.LOC_DESCRI," & _
'               "  C.LAD_TELEFONO  " & _
'               " FROM LAB_DENTALES C, LOCALIDAD L " & _
'               " WHERE C.LOC_CODIGO = L.LOC_CODIGO "
'        .HeaderSQL = "Nombre, Código, Domicilio, Localidad, Teléfono"
'        .FieldID = "LAD_CODIGO"
'        '.Report = RptPath & "tipocomp.rpt"
'        Set .FormBase = vFormLabDentales
'        Set .FormDatos = ABMLabDentales
'    End With
'
'    Set auxDllActiva = vABMLabDentales
'
'    vABMLabDentales.Show
End Sub

Private Sub mnuMedicamentos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMMedicamentos = New CListaBaseABM
    
    With vABMMedicamentos
        .Caption = "Actualizacion de Medicamentos"
        .sql = "SELECT M.MED_NOMBRE, M.MED_CODIGO,M.MED_PRESENTACION, M.MED_DOSIFICACION,G.GRU_DESCRI," & _
               "  M.MED_EDAD  " & _
               " FROM MEDICAMENTOS M, GRUPOS G " & _
               " WHERE M.GRU_CODIGO = G.GRU_CODIGO "
        .HeaderSQL = "Nombre, Código,Prsentacion, Dosificacion, Grupo, Edad"
        .FieldID = "MED_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormMedicamentos
        Set .FormDatos = ABMMedicamentos
    End With
    
    Set auxDllActiva = vABMMedicamentos
    
    vABMMedicamentos.Show
End Sub

Private Sub mnuNotas_Click()
    On Error Resume Next
    Shell "C:\WINDOWS\system32\notepad.exe", vbNormalFocus
    
End Sub

Private Sub mnuObrasSociales_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMObraSocial = New CListaBaseABM
    
    With vABMObraSocial
        .Caption = "Actualización de Obras Sociales"
        .sql = "SELECT OS_NUMERO, OS_NOMBRE,OS_DOMICI, OS_TELEFONO FROM OBRA_SOCIAL"
        .HeaderSQL = "Numero, Nombre, Domicilio, Teléfono"
        .FieldID = "OS_NUMERO"
        '.Report = RPTPATH & "tarjeta_credito.rpt"
        Set .FormBase = vFormObraSocial
        Set .FormDatos = ABMObraSocial
    End With
    
    Set auxDllActiva = vABMObraSocial
    
    vABMObraSocial.Show
End Sub

Private Sub mnuPacientes_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMClientes = New CListaBaseABM
    
    With vABMClientes
        .Caption = "Actualizacion de Pacientes"
        .sql = "SELECT C.CLI_RAZSOC,C.CLI_NRODOC, C.CLI_CODIGO, C.CLI_DOMICI,C.CLI_CODPOS, L.LOC_DESCRI," & _
               "  C.CLI_TELEFONO, C.CLI_EDAD,C.CLI_OCUPACION " & _
               " FROM CLIENTE C, LOCALIDAD L " & _
               " WHERE C.LOC_CODIGO = L.LOC_CODIGO"
        .HeaderSQL = "Nombre,Documento, Código, Domicilio, Código Postal,Localidad,Teléfono, Edad, Ocupacion "
        .FieldID = "CLI_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormClientes
        Set .FormDatos = ABMClientes
    End With
    
    Set auxDllActiva = vABMClientes
    
    vABMClientes.Show
    
End Sub

Private Sub mnuParametros_Click()
    frmParametros.Show
End Sub

Private Sub mnuPermisos_Click()
    FrmPermisos.Show
End Sub

Private Sub mnuPersonal_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMVendedor = New CListaBaseABM
    
    With vABMVendedor
        .Caption = "Actualizacion de Personal"
        .sql = "SELECT C.VEN_NOMBRE, C.VEN_CODIGO,P.PR_DESCRI, C.VEN_DOMICI,L.LOC_DESCRI," & _
               "  C.VEN_TELEFONO  " & _
               " FROM VENDEDOR C, LOCALIDAD L , PROFESION P" & _
               " WHERE C.LOC_CODIGO = L.LOC_CODIGO AND P.PR_CODIGO = C.PR_CODIGO "
        .HeaderSQL = "Nombre, Código,Ocupacion, Domicilio, Localidad, Teléfono"
        .FieldID = "VEN_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormVendedor
        Set .FormDatos = ABMVendedor
    End With
    
    Set auxDllActiva = vABMVendedor
    
    vABMVendedor.Show
End Sub

Private Sub mnuProfesiones_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMProfesiones = New CListaBaseABM
    
    With vABMProfesiones
        .Caption = "Actualización de Profesiones"
        .sql = "SELECT PR_DESCRI,PR_CODIGO FROM PROFESION"
        .HeaderSQL = "Descripcion, Codigo"
        .FieldID = "PR_CODIGO"
        '.Report = RPTPATH & "tarjeta_credito.rpt"
        Set .FormBase = vFormObraSocial
        Set .FormDatos = ABMProfesion
    End With
    
    Set auxDllActiva = vABMProfesiones
    
    vABMProfesiones.Show
End Sub

Private Sub mnuProtocolos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMProtocolos = New CListaBaseABM
    
    With vABMProtocolos
        .Caption = "Actualizacion de Protocolos"
        .sql = "SELECT TIP_NOMBRE,TIP_CODIGO FROM TIPO_IMAGEN "
        .HeaderSQL = "Nombre,Código"
        .FieldID = "TIP_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormProtocolos
        Set .FormDatos = ABMProtocolos
    End With
    
    Set auxDllActiva = vABMProtocolos
    
    vABMProtocolos.Show
End Sub

Private Sub mnuRestArchivos_Click()
    With frmRestaurarBD
        .Caption = "Restaurar Archivos"
        .optCopiarDesde.Value = True
        .Label1 = "Restaurar desde:"
        .Show
    End With
End Sub
Private Sub mnuAyuda_Click()
    'On Error Resume Next
    'LeoIni
   
   ' Open Ayuda & "HC000.HTM" For Input As #1
End Sub

Private Sub mnuTratamientos_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMTratamiento = New CListaBaseABM
    
    With vABMTratamiento
        .Caption = "Actualizacion de Tratamiento"
        .sql = "SELECT TR_CODNUE,TR_DESCRI, TR_CODIGO,TR_PRECIO" & _
               " FROM TRATAMIENTO"
               '" WHERE C.LOC_CODIGO = L.LOC_CODIGO AND P.PR_CODIGO = C.PR_CODIGO "
        .HeaderSQL = "Codigo,Tratamiento, Id,Precio"
        .FieldID = "TR_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormTratamiento
        Set .FormDatos = ABMTratamiento
    End With
    
    Set auxDllActiva = vABMTratamiento
    
    vABMTratamiento.Show
End Sub

Private Sub mnuUsuario_Click()
    FrmUsuarios.Show vbModal
End Sub

Private Sub Motivo_Click()
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMMotivo = New CListaBaseABM
    
    With vABMMotivo
        .Caption = "Actualización de Motivos de turno"
        .sql = "SELECT MOT_CODIGO, MOT_DESCRI FROM MOTIVO"
        .HeaderSQL = "Numero, Nombre"
        .FieldID = "MOT_CODIGO"
        '.Report = RPTPATH & "tarjeta_credito.rpt"
        Set .FormBase = vFormMotivo
        Set .FormDatos = ABMMotivo
    End With
    
    Set auxDllActiva = vABMMotivo
    
    vABMMotivo.Show
End Sub

Private Sub tbrPrincipal_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 2: Call mnuArcSal_Click
        Case 3: Call mnuAsignarTurnos_Click
        Case 4: Call mnuArchivoActualizaciones_Click
        Case 6: Call mnuTratamientos_Click
        Case 7: Call mnuMedicamentos_Click
        Case 9: Call mnuCumpleaños_Click
        Case 10: Call mnuControlVisitas_Click
        
    End Select
End Sub


Private Sub txtPaciente_GotFocus()
    SelecTexto txtPaciente
End Sub

Private Sub txtPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then
        'txtPaciente_LostFocus
    End If
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtPaciente", "CODIGO"
    End If
End Sub

Private Sub txtPaciente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)

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
        
        hSQL = "Nombre, Código, DNI "
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .Field = "CLI_DNI"
        campo3 = .Field
        
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtcodCli" Then
                'txtCliente.Text = .ResultFields(2)
                'txtCliente_LostFocus
            Else
                txtPaciente.Text = .ResultFields(3)
                txtPaciente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub txtPaciente_LostFocus()
    Set Rec2 = New ADODB.Recordset
    
    If txtPaciente.Text <> "" Then
        Set Rec2 = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtPaciente.Text <> "" Then
            sql = sql & " CLI_NRODOC=" & XN(txtPaciente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtPaciente) & "%'"
        End If
        Rec2.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If Rec2.EOF = False Then
            ' aca entra al formulario que tiene que contener
            ' la HC, Turnos, Presupuestos del cliente
            ' Pagos y Cobros
            frmDatosClientes.txtDNI.Text = Rec2!CLI_NRODOC
            frmDatosClientes.txtDNI.ToolTipText = Rec2!CLI_CODIGO
            frmDatosClientes.Caption = "Paciente: " & Rec2!CLI_RAZSOC
            frmDatosClientes.lblPaciente = Rec2!CLI_RAZSOC
            If frmDatosClientes.Visible = False Then
                frmDatosClientes.Show vbModal
                txtPaciente.Text = ""
            End If
            'txtDesCli.Text = rec!CLI_RAZSOC
            'txtcodigo.Text = rec!CLI_CODIGO
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
            txtPaciente.SetFocus
            txtPaciente.Text = ""
        End If
        If Rec2.State = 1 Then
            Rec2.Close
        End If
    End If
End Sub

Private Function buscarcumples(hoy As Date)
Dim nHayCumple As Integer '1 hay 0 no hay
    sql = "SELECT CLI_CUMPLE,CLI_RAZSOC,CLI_EDAD,CLI_TELEFONO,CLI_CODIGO,CLI_MAIL "
    sql = sql & " FROM CLIENTE "
    'sql = sql & " WHERE DatePart('dd',CLI_CUMPLE) =  " & day(hoy)
    'sql = sql & " AND DatePart('mm',CLI_CUMPLE) =  " & Month(hoy)
    'sql = sql & " to_date(to_char(CLI_CUMPLE, 'DD/MM') || '/2001', 'DD/MM/YYYY') >= fecha_ini and"
    'to_date(to_char(fecha_nac, 'DD/MM') || '/2001', 'DD/MM/YYYY') <= fecha_fin

    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    nHayCumple = 0
    If rec.EOF = False Then
        Do While rec.EOF = False
            'CALCULAR Y GUARDAR NUEVA EDAD DEL PACIENTE
            If Day(Chk0(rec!CLI_CUMPLE)) = Day(hoy) And Month(Chk0(rec!CLI_CUMPLE)) = Month(hoy) Then
                nHayCumple = 1
            End If
            rec.MoveNext
        Loop
    End If
    rec.Close
    If nHayCumple = 1 Then
        frmCumpleaños.Show
    End If
End Function
Private Function buscarRecordatorio(hoy As Date)
    Dim nHayCClinico As Integer '1 hay 0 no hay
    sql = "SELECT CL.CCL_FECPC,P.CLI_RAZSOC,P.CLI_TELEFONO,CL.CLI_CODIGO,P.CLI_MAIL, "
    sql = sql & " T.TR_DESCRI"
    sql = sql & " FROM CLIENTE P, CCLINICO CL, TRATAMIENTO T"
    sql = sql & " WHERE P.CLI_CODIGO = CL.CLI_CODIGO "
    sql = sql & " AND T.TR_CODIGO = CL.TR_CODIGO "
    sql = sql & " AND CL.CCL_FECPC = " & XDQ(hoy)
   
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic

    nHayCClinico = 0
    If rec.EOF = False Then
        nHayCClinico = 1
        'Do While rec.EOF = False
        '    'CALCULAR Y GUARDAR NUEVA EDAD DEL PACIENTE
        '    If day(Chk0(rec!CCL_FECPC)) = day(hoy) And Month(Chk0(rec!CCL_FECPC)) = Month(hoy) Then
        '        nHayCClinico = 1
        '    End If
        '    rec.MoveNext
        'Loop
    End If
    rec.Close
    If nHayCClinico = 1 Then
        frmControlVisitas.Show
    End If
End Function


Private Function configuroLetrero(Fecha As Date) As String
    
    Dim DIA As Integer
    Dim Doc As Integer
    DIA = Weekday(Fecha, vbMonday)
    Doc = 0
    If User <> 99 Then
        Doc = XN(User)
    End If
    
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
        
    
    'busco turnos
    sql = "SELECT T.*,V.VEN_NOMBRE,C.CLI_RAZSOC,C.CLI_DNI"
    sql = sql & " FROM TURNOS T, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND T.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND T.TUR_FECHA = " & XDQ(Fecha)
    If Doc <> 0 Then
        sql = sql & " AND T.VEN_CODIGO = " & Doc
    End If
    sql = sql & " ORDER BY V.VEN_CODIGO,T.TUR_HORAD"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        If Doc = 0 Then
            Doc = rec!VEN_CODIGO
        End If
        Letrero = "Turnos del día " & WeekdayName(DIA, False) & " " & Day(Fecha) & " de " & MonthName(Month(Fecha), False) & " de " & Year(Fecha) & " del Dr: " & rec!VEN_NOMBRE & ": "
        Do While rec.EOF = False
            If Doc <> rec!VEN_CODIGO Then
                Letrero = Letrero & ". Turnos del Dr: " & rec!VEN_NOMBRE & ": "
            End If
            Letrero = Letrero & Format(rec!TUR_HORAD, "hh:mm") & " Hs - " & _
                          rec!CLI_RAZSOC & " , " & rec!TUR_MOTIVO & " // "
            
            rec.MoveNext
        Loop
    End If
    rec.Close
    Letrero = Letrero
End Function
Private Sub Timer1_Timer()
    
    Static Anterior As Boolean
    Static tamañoLetrero As Single
    Static X As Single
    If Not Anterior Then
        tamañoLetrero = Menu.Picture2.TextWidth(Letrero)
        Anterior = True
        X = Menu.Picture2.ScaleWidth
    End If
    Menu.Picture2.Cls
    Menu.Picture2.CurrentX = X
    Menu.Picture2.CurrentY = 100
'Para cambiar el tipo de letra
    Menu.Picture2.FontName = "Arial"
    Menu.Picture2.FontBold = True
    Menu.Picture2.Print Letrero
    X = X - 25
    If X < -tamañoLetrero Then X = Menu.Picture2.ScaleWidth
End Sub

