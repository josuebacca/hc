VERSION 5.00
Begin VB.Form frmAuditoriaTurno 
   Caption         =   "Auditoría Turno"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblUpdatedAt 
      Caption         =   "At"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblActualizadoPor 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Última actualización:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblCreatedAt 
      Caption         =   "At"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblCreadoPor 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   540
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Creación:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "frmAuditoriaTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblAt_Click()

End Sub
Public Sub CargarDatos(creador As String, fCreacion As Variant, actualizador As String, fActualizacion As Variant)

    lblCreadoPor.Caption = creador

    If IsDate(fCreacion) Then
        lblCreatedAt.Caption = Format(CDate(fCreacion), "dd/mm/yyyy hh:mm")
    Else
        lblCreatedAt.Caption = "-"
    End If

    lblActualizadoPor.Caption = actualizador

    If IsDate(fActualizacion) Then
        lblUpdatedAt.Caption = Format(CDate(fActualizacion), "dd/mm/yyyy hh:mm")
    Else
        lblUpdatedAt.Caption = "-"
    End If

End Sub
