VERSION 5.00
Begin VB.Form Ecografía 
   Caption         =   "Ecografía"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   11415
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   11535
      Begin VB.CommandButton cmdVerTG 
         Caption         =   "Ver Tamaño Grande"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   9600
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   7560
         TabIndex        =   3
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Text            =   "Observaciones"
         Top             =   480
         Width           =   11295
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Ecografía"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
