VERSION 5.00
Begin VB.Form frmIngresarClave 
   Caption         =   "Ingresar clave"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      DisabledPicture =   "frmIngresarClave.frx":0000
      Height          =   750
      Left            =   1560
      Picture         =   "frmIngresarClave.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      DisabledPicture =   "frmIngresarClave.frx":0614
      Height          =   750
      Left            =   2775
      Picture         =   "frmIngresarClave.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox txtClave 
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
      Left            =   360
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese su clave personal para continuar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmIngresarClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ClaveIngresada As String

Private Sub Form_Load()
    txtClave.PasswordChar = "*"
End Sub

Private Sub cmdAceptar_Click()
    ClaveIngresada = txtClave.text
    Me.Hide
End Sub

Private Sub CmdCancelar_Click()
    ClaveIngresada = ""
    Me.Hide
End Sub
Private Sub Form_Activate()
    txtClave.SetFocus
End Sub
