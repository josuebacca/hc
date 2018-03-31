VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MenuCore 
   BorderStyle     =   0  'None
   Caption         =   "Turnos del Dia"
   ClientHeight    =   6585
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H000000FF&
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   -120
      Width           =   11415
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turnos del Dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3735
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "VER"
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
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrillaAplicar 
      Height          =   9975
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   17595
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   300
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":0E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":1982
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":1C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":24B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":27D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":2AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":2E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":311E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":3438
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":3752
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":3A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":3D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":40A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":43BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":46D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":49EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MenuCore.frx":4D08
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MenuCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuPacientes_Click()
    Dim cSQL As String
    
    mOrigen = True
    
    Set vABMClientes = New CListaBaseABM
    
    With vABMClientes
        .Caption = "Actualizacion de Pacientes"
        .SQL = "SELECT C.CLI_RAZSOC, C.CLI_CODIGO, C.CLI_DOMICI,C.CLI_CODPOS, L.LOC_DESCRI," & _
        "  C.CLI_NRODOC, C.CLI_TELEFONO, C.CLI_EDAD,C.CLI_OCUPACION " & _
        " FROM CLIENTE C, LOCALIDAD L " & _
        " WHERE C.LOC_CODIGO = L.LOC_CODIGO"
        .HeaderSQL = "Nombre, Código, Domicilio, Código Postal,Localidad,Documento,Teléfono, Edad, Ocupacion "
        .FieldID = "CLI_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
    Set .FormBase = vFormClientes
    Set .FormDatos = ABMClientes
    End With
    
    Set auxDllActiva = vABMClientes
    
    vABMClientes.Show
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

