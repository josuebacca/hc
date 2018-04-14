VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmhistoriaclinica 
   Caption         =   "Historia Clinica"
   ClientHeight    =   11205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19935
   LinkTopic       =   "Form1"
   ScaleHeight     =   11205
   ScaleWidth      =   19935
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabhc 
      Height          =   9735
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Curso Clinico"
      TabPicture(0)   =   "frmhistoriaclinica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSFlexGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAgregarPedido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdImprimirPedido"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEliminarImag"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDactual"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboDia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboMes"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboAño"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdBuscar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Ecografias"
      TabPicture(1)   =   "frmhistoriaclinica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdEliminarEco"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAgregarEco"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdVer"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdVerEstudio"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Laboratorio"
      TabPicture(2)   =   "frmhistoriaclinica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Ginecologia"
      TabPicture(3)   =   "frmhistoriaclinica.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "optTipoEst"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "optFechaGine"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboTipoEstGine"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame3"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdAgregarEstGine"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdEliminarEstGine"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdImprimirEstGine"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.CommandButton cmdImprimirEstGine 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -57480
         TabIndex        =   41
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEstGine 
         Caption         =   "Eliminar Estudio"
         Height          =   375
         Left            =   -59400
         TabIndex        =   40
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CommandButton cmdAgregarEstGine 
         Caption         =   "Agregar Estudio"
         Height          =   375
         Left            =   -61080
         TabIndex        =   39
         Top             =   9000
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estudios"
         Height          =   7095
         Left            =   -74880
         TabIndex        =   37
         Top             =   1680
         Width           =   19095
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   6615
            Left            =   120
            TabIndex        =   38
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
         TabIndex        =   36
         Text            =   "Combo2"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton optFechaGine 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   -68400
         TabIndex        =   35
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optTipoEst 
         Caption         =   "Tipo Estudio"
         Height          =   375
         Left            =   -73320
         TabIndex        =   34
         Top             =   600
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdVerEstudio 
         Caption         =   "Ver Estudio(VA ACA?)"
         Height          =   375
         Left            =   -60360
         TabIndex        =   32
         Top             =   8760
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   7800
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboAño 
         Height          =   315
         Left            =   5400
         TabIndex        =   29
         Text            =   "Año"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   3600
         TabIndex        =   28
         Text            =   "Mes"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboDia 
         Height          =   315
         Left            =   1800
         TabIndex        =   27
         Text            =   "Dia"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -61800
         TabIndex        =   26
         Top             =   8760
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregarEco 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -58800
         TabIndex        =   25
         Top             =   8760
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminarEco 
         Caption         =   "Elinimar"
         Height          =   375
         Left            =   -57000
         TabIndex        =   24
         Top             =   8760
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ecografías"
         Height          =   8055
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   19095
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   6255
            Left            =   120
            TabIndex        =   23
            Top             =   1680
            Width           =   18855
            _ExtentX        =   33258
            _ExtentY        =   11033
            _Version        =   393216
            Rows            =   4
            Cols            =   5
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   11880
            TabIndex        =   22
            Text            =   "Doctor"
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   8760
            TabIndex        =   21
            Text            =   "Combo1"
            Top             =   840
            Width           =   2295
         End
         Begin VB.ComboBox cboEspecialidad 
            Height          =   315
            Left            =   4800
            TabIndex        =   20
            Text            =   "Especialidad"
            Top             =   840
            Width           =   2655
         End
         Begin VB.OptionButton optEmpleado 
            Caption         =   "Doctor"
            Height          =   255
            Left            =   11880
            TabIndex        =   19
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton optFecha 
            Caption         =   "Fecha"
            Height          =   375
            Left            =   8760
            TabIndex        =   18
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optEspecialidad 
            Caption         =   "Especialidad"
            Height          =   255
            Left            =   4800
            TabIndex        =   17
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Buscar por:"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtDactual 
         Height          =   495
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   9120
         Width           =   4815
      End
      Begin VB.CommandButton cmdEliminarImag 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   16800
         TabIndex        =   12
         Top             =   9120
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimirPedido 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   18120
         TabIndex        =   11
         Top             =   9120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarPedido 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   15480
         TabIndex        =   10
         Top             =   9120
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7695
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   19215
         _ExtentX        =   33893
         _ExtentY        =   13573
         _Version        =   393216
         Cols            =   4
      End
      Begin VB.Label Label8 
         Caption         =   "Buscar por:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Buscar por fecha:"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Diagnóstico actual:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   9120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Paciente"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19455
      Begin VB.TextBox txtObra 
         Height          =   285
         Left            =   7200
         TabIndex        =   7
         Text            =   "Obra Social"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDNI 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbNombre 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Text            =   "Nombre"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Buscar por fecha:"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Obra Social:"
         Height          =   255
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "D.N.I.:"
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre y apellido:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmhistoriaclinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub optEspecialidad_Click()

End Sub

Private Sub Option1_Click()

End Sub
