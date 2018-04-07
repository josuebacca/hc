VERSION 5.00
Begin VB.Form frmDatosClientes 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   Icon            =   "frmDatosClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDatosClientes.frx":08CA
   ScaleHeight     =   4080
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   6315
      TabIndex        =   6
      Top             =   3240
      Width           =   6375
      Begin VB.Label lblPaciente 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Información centralizada del Paciente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.TextBox txtDNI 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      Picture         =   "frmDatosClientes.frx":5E08
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Picture         =   "frmDatosClientes.frx":6E4A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Presupuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      Picture         =   "frmDatosClientes.frx":7714
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdTurnos 
      Caption         =   "Turnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      Picture         =   "frmDatosClientes.frx":7A1E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdHClinica 
      Caption         =   "Historia Clinica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      MaskColor       =   &H80000010&
      Picture         =   "frmDatosClientes.frx":7D28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmDatosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdHClinica_Click()
    'POR AHORA HACEMOS ESTO, DESPUES VEMOS SI ENTRAMOS A CADA SOLAPA EN PARTICULAR
    Dim cSQL As String
    
    mOrigen = True
        
    Set vABMClientes = New CListaBaseABM
    
    With vABMClientes
        .Caption = "Actualizacion de Pacientes"
        .sql = "SELECT C.CLI_RAZSOC,C.CLI_DNI, C.CLI_CODIGO, C.CLI_DOMICI,C.CLI_CODPOS, L.LOC_DESCRI," & _
               "  C.CLI_NRODOC, C.CLI_TELEFONO, C.CLI_EDAD,C.CLI_OCUPACION " & _
               " FROM CLIENTE C, LOCALIDAD L " & _
               " WHERE C.LOC_CODIGO = L.LOC_CODIGO"
        .HeaderSQL = "Nombre,DNI, Código, Domicilio, Código Postal,Localidad,Documento,Teléfono, Edad, Ocupacion "
        .FieldID = "CLI_CODIGO"
        '.Report = RptPath & "tipocomp.rpt"
        Set .FormBase = vFormClientes
        Set .FormDatos = ABMClientes
    End With
    
    Set auxDllActiva = vABMClientes
    
    vABMClientes.Show
End Sub

Private Sub cmdTurnos_Click()
    frmBuscarTurnos.txtCliente = txtDNI.Text
    frmBuscarTurnos.Show vbModal
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command6_Click()
    'Set rec = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call Centrar_pantalla(Me)
End Sub
