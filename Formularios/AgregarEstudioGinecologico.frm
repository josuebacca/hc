VERSION 5.00
Begin VB.Form AgregarEstudioGinecologico 
   Caption         =   "Agregar Estudio Ginecologico"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdCancelarGine 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdAceptarGine 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cboGinecologos 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Text            =   "Ginecólogos"
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cboEstudiosGinecologicos 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Text            =   "Estudios"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Ginecólogo:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Estudio Ginecológico:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "AgregarEstudioGinecologico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
