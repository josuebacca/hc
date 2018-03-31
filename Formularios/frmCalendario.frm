VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCalendario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendario"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      MonthBackColor  =   15925247
      StartOfWeek     =   60686337
      CurrentDate     =   37455
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    MonthView1.Value = Date
End Sub

