VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRestaurarBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restaurar Base de Datos"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdSelArch 
         Caption         =   "..."
         Height          =   325
         Left            =   3165
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   425
      End
      Begin VB.ComboBox cboUnidad 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin VB.OptionButton optCopiarDesde 
         Caption         =   "Copiar Base de Datos Desde"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton optCopiarA 
         Caption         =   "Copiar Base de Datos A"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Archivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Disco:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1230
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "frmRestaurarBD.frx":0000
      Height          =   720
      Left            =   2520
      Picture         =   "frmRestaurarBD.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   915
   End
   Begin MSComDlg.CommonDialog ComD1 
      Left            =   240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "frmRestaurarBD.frx":0614
      Height          =   720
      Left            =   1560
      Picture         =   "frmRestaurarBD.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   915
   End
End
Attribute VB_Name = "frmRestaurarBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
    Dim origen As String
    Dim destino As String
    Dim objeto As FileSystemObject
    Set objeto = New FileSystemObject
    
    Dim resp As String
    If cboUnidad.ListIndex > 0 Then
        If optCopiarA.Value = True Then
            resp = MsgBox("Esta a punto de hacer una Backup(copia) de la Base de Datos del Sistema en la unidad de Disco: " & cboUnidad.Text & " Confirma la accion?", 36, "Copiar A:")
            If resp <> 6 Then Exit Sub
            Screen.MousePointer = vbHourglass
            LeoIni
            
            origen = DirBkp & BASEDATO & ".MDB"
            destino = cboUnidad.Text
            objeto.CopyFile origen, destino, True
            Screen.MousePointer = vbNormal
            MsgBox "La copia en la Unidad de Disco ha sido exitosa", vbInformation, TIT_MSGBOX
            
        Else
            If optCopiarDesde.Value = True Then
                resp = MsgBox("Esta a punto de Restaurar la Base de Datos en el Sistema desde la Unidad: " & cboUnidad.Text & " Esto puede provocar que se pierda datos. Confirma la accion?", 36, "Copiar Desde:")
                If resp <> 6 Then Exit Sub
                Screen.MousePointer = vbHourglass
                origen = cboUnidad.Text & BASEDATO & ".MDB"
                destino = DirBkp
                objeto.CopyFile origen, destino, True
                Screen.MousePointer = vbNormal
                MsgBox "La Restauracion en el Sistema ha sido exitosa", vbInformation, TIT_MSGBOX
            End If
            
        End If
    Else
        MsgBox "Debe seleccionar una Unidad de Disco", vbExclamation, TIT_MSGBOX
        Screen.MousePointer = vbNormal
        cboUnidad.SetFocus
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelArch_Click()
On Error Resume Next
ComD1.CancelError = True
ComD1.DialogTitle = "Abrir"
'ComD1.Filter = "*.xls"

ComD1.ShowOpen
txtArchivo.Text = ComD1.FileName

End Sub

Private Sub Form_Load()
    Centrar_pantalla Me
    cargarCombo
End Sub

Private Function cargarCombo()
    cboUnidad.AddItem ""
    cboUnidad.AddItem "C:\"
    cboUnidad.AddItem "D:\"
    cboUnidad.AddItem "E:\"
    cboUnidad.AddItem "F:\"
    cboUnidad.AddItem "G:\"
End Function

