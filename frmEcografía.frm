VERSION 5.00
Begin VB.Form frmEcografía 
   Caption         =   "Ecografía"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton cmdCerrarEco 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton cmdImprimirEco 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   5520
         Width           =   1935
      End
      Begin VB.CommandButton cmdVerTG 
         Caption         =   "Ver Tamaño Grande"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox txtObservEco 
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   4200
         Width           =   7935
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   3840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEcografía"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVerTG_Click()

End Sub
