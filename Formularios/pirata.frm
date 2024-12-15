VERSION 5.00
Begin VB.Form frmPirata 
   Caption         =   "Registración"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   75
      TabIndex        =   4
      Top             =   945
      Width           =   3810
      Begin VB.TextBox txtFija 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   300
         TabIndex        =   5
         Top             =   675
         Width           =   1440
      End
      Begin VB.TextBox txtVariable 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1845
         TabIndex        =   0
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. de Serie:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         TabIndex        =   6
         Top             =   345
         Width           =   2490
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   450
      Left            =   2370
      TabIndex        =   2
      Top             =   2565
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   795
      TabIndex        =   1
      Top             =   2565
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "para obtener la Clave de Registración."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   705
      Width           =   3825
   End
   Begin VB.Label Label3 
      Caption         =   "oficinasdigitales@yahoo.com.ar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   105
      TabIndex        =   7
      Top             =   375
      Width           =   2805
   End
   Begin VB.Label Label1 
      Caption         =   "Envie un e-mail a: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   90
      Width           =   3975
   End
End
Attribute VB_Name = "frmPirata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    mVer = Verifico_Clave
    If mVer Then
        'grabo el archivito para que no pida mas clave
        Open "c:\windows\system\W95INF128.DLL" For Output As #1
            Print #1, "."
        Close #1
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    mVer = False
    Unload Me
End Sub

Private Sub Form_Load()
    Randomize
    Dim mClaveFija As Integer
    mClaveFija = ((7199 - 7100 + 1) * Rnd + 7100)
    txtFija.Text = mClaveFija
    
'    Open "c:\CLAVES.TXT" For Output As #1
'    For i = 1 To 100
'        MCARMEN1 = Chr(Int((90 - 65 + 1) * Rnd + 65))
'        MCARMEN2 = Int((9 - 1 + 1) * Rnd + 1)
'        MCARMEN3 = Chr(Int((90 - 65 + 1) * Rnd + 65))
'        MCARMEN4 = Int((9 - 1 + 1) * Rnd + 1)
'        MCARMEN5 = Chr(Int((90 - 65 + 1) * Rnd + 65))
'        MCARMEN6 = MCARMEN1 & MCARMEN2 & MCARMEN3 & MCARMEN4 & MCARMEN5
'
'        Print #1, MCARMEN6
'    Next
'    Close #1
    
    
End Sub
Public Function Verifico_Clave() As Boolean
    Dim mBien As Boolean
    mBien = False
    Select Case txtFija.Text
        Case "7100"
            If txtVariable.Text = "N6H3U" Then mBien = True
        Case "7101"
            If txtVariable.Text = "A7V7B" Then mBien = True
        Case "7102"
            If txtVariable.Text = "K8U4Z" Then mBien = True
        Case "7103"
            If txtVariable.Text = "W1Y4N" Then mBien = True
        Case "7104"
            If txtVariable.Text = "T1P5H" Then mBien = True
        Case "7105"
            If txtVariable.Text = "Q6G3V" Then mBien = True
        Case "7106"
            If txtVariable.Text = "V6Z9F" Then mBien = True
        Case "7107"
            If txtVariable.Text = "S9G5C" Then mBien = True
        Case "7108"
            If txtVariable.Text = "Z7A6C" Then mBien = True
        Case "7109"
            If txtVariable.Text = "C8H1H" Then mBien = True
        Case "7110"
            If txtVariable.Text = "J3Y9K" Then mBien = True
        Case "7111"
            If txtVariable.Text = "H2E6K" Then mBien = True
        Case "7112"
            If txtVariable.Text = "K7I6F" Then mBien = True
        Case "7113"
            If txtVariable.Text = "E6C5X" Then mBien = True
        Case "7114"
            If txtVariable.Text = "G8J3X" Then mBien = True
        Case "7115"
            If txtVariable.Text = "Q6L1O" Then mBien = True
        Case "7116"
            If txtVariable.Text = "S9V1O" Then mBien = True
        Case "7117"
            If txtVariable.Text = "X4R5N" Then mBien = True
        Case "7118"
            If txtVariable.Text = "M4K3B" Then mBien = True
        Case "7119"
            If txtVariable.Text = "G9B4J" Then mBien = True
        Case "7120"
            If txtVariable.Text = "M2M3Q" Then mBien = True
        Case "7121"
            If txtVariable.Text = "O2Y6N" Then mBien = True
        Case "7122"
            If txtVariable.Text = "K1U5T" Then mBien = True
        Case "7123"
            If txtVariable.Text = "P8A2B" Then mBien = True
        Case "7124"
            If txtVariable.Text = "C3D1N" Then mBien = True
        Case "7125"
            If txtVariable.Text = "R5V1E" Then mBien = True
        Case "7126"
            If txtVariable.Text = "R5J2S" Then mBien = True
        Case "7127"
            If txtVariable.Text = "Y5C7K" Then mBien = True
        Case "7128"
            If txtVariable.Text = "M5F3C" Then mBien = True
        Case "7129"
            If txtVariable.Text = "P2Y1L" Then mBien = True
        Case "7130"
            If txtVariable.Text = "H8T3R" Then mBien = True
        Case "7131"
            If txtVariable.Text = "G1A3U" Then mBien = True
        Case "7132"
            If txtVariable.Text = "H3M3I" Then mBien = True
        Case "7133"
            If txtVariable.Text = "B5F8P" Then mBien = True
        Case "7134"
            If txtVariable.Text = "T9I5C" Then mBien = True
        Case "7135"
            If txtVariable.Text = "Q4Y2Y" Then mBien = True
        Case "7136"
            If txtVariable.Text = "Q4D5F" Then mBien = True
        Case "7137"
            If txtVariable.Text = "Z2A4O" Then mBien = True
        Case "7138"
            If txtVariable.Text = "X5K8V" Then mBien = True
        Case "7139"
            If txtVariable.Text = "R7Z4M" Then mBien = True
        Case "7140"
            If txtVariable.Text = "K7E4O" Then mBien = True
        Case "7141"
            If txtVariable.Text = "V5L5F" Then mBien = True
        Case "7142"
            If txtVariable.Text = "Q5R8J" Then mBien = True
        Case "7143"
            If txtVariable.Text = "H3D5F" Then mBien = True
        Case "7144"
            If txtVariable.Text = "P4W5E" Then mBien = True
        Case "7145"
            If txtVariable.Text = "R7P8E" Then mBien = True
        Case "7146"
            If txtVariable.Text = "V2Y1B" Then mBien = True
        Case "7147"
            If txtVariable.Text = "U4M2D" Then mBien = True
        Case "7148"
            If txtVariable.Text = "E1S5O" Then mBien = True
        Case "7149"
            If txtVariable.Text = "F5T7K" Then mBien = True
        Case "7150"
            If txtVariable.Text = "X7C6S" Then mBien = True
        Case "7151"
            If txtVariable.Text = "A4K3Z" Then mBien = True
        Case "7152"
            If txtVariable.Text = "U7K7H" Then mBien = True
        Case "7153"
            If txtVariable.Text = "J4Y2Q" Then mBien = True
        Case "7154"
            If txtVariable.Text = "J1E1L" Then mBien = True
        Case "7155"
            If txtVariable.Text = "Y5M9F" Then mBien = True
        Case "7156"
            If txtVariable.Text = "J4H5D" Then mBien = True
        Case "7157"
            If txtVariable.Text = "N9O9R" Then mBien = True
        Case "7158"
            If txtVariable.Text = "L7B7S" Then mBien = True
        Case "7159"
            If txtVariable.Text = "M2F3U" Then mBien = True
        Case "7160"
            If txtVariable.Text = "B5T8I" Then mBien = True
        Case "7161"
            If txtVariable.Text = "Z8R9W" Then mBien = True
        Case "7162"
            If txtVariable.Text = "K2Y8S" Then mBien = True
        Case "7163"
            If txtVariable.Text = "K1E2N" Then mBien = True
        Case "7164"
            If txtVariable.Text = "K1H6W" Then mBien = True
        Case "7165"
            If txtVariable.Text = "M2X4I" Then mBien = True
        Case "7166"
            If txtVariable.Text = "U2L3W" Then mBien = True
        Case "7167"
            If txtVariable.Text = "P4K8P" Then mBien = True
        Case "7168"
            If txtVariable.Text = "Y5I8G" Then mBien = True
        Case "7169"
            If txtVariable.Text = "G2J1T" Then mBien = True
        Case "7170"
            If txtVariable.Text = "V3S4V" Then mBien = True
        Case "7171"
            If txtVariable.Text = "T4C4I" Then mBien = True
        Case "7172"
            If txtVariable.Text = "S3U2P" Then mBien = True
        Case "7173"
            If txtVariable.Text = "Y3Y2Z" Then mBien = True
        Case "7174"
            If txtVariable.Text = "Q6X6G" Then mBien = True
        Case "7175"
            If txtVariable.Text = "W1L7G" Then mBien = True
        Case "7176"
            If txtVariable.Text = "J4N3P" Then mBien = True
        Case "7177"
            If txtVariable.Text = "F1X2Q" Then mBien = True
        Case "7178"
            If txtVariable.Text = "X3Y8L" Then mBien = True
        Case "7179"
            If txtVariable.Text = "M7V4F" Then mBien = True
        Case "7180"
            If txtVariable.Text = "I4D6C" Then mBien = True
        Case "7181"
            If txtVariable.Text = "F7N2Z" Then mBien = True
        Case "7182"
            If txtVariable.Text = "M1L3T" Then mBien = True
        Case "7183"
            If txtVariable.Text = "X8E1A" Then mBien = True
        Case "7184"
            If txtVariable.Text = "D4E9W" Then mBien = True
        Case "7185"
            If txtVariable.Text = "P7D3A" Then mBien = True
        Case "7186"
            If txtVariable.Text = "V8U3L" Then mBien = True
        Case "7187"
            If txtVariable.Text = "O8O2F" Then mBien = True
        Case "7188"
            If txtVariable.Text = "J6I5L" Then mBien = True
        Case "7189"
            If txtVariable.Text = "O1O2Z" Then mBien = True
        Case "7190"
            If txtVariable.Text = "Y9V4U" Then mBien = True
        Case "7191"
            If txtVariable.Text = "Z3E7A" Then mBien = True
        Case "7192"
            If txtVariable.Text = "I7C1H" Then mBien = True
        Case "7193"
            If txtVariable.Text = "K6I3V" Then mBien = True
        Case "7194"
            If txtVariable.Text = "I4K8K" Then mBien = True
        Case "7195"
            If txtVariable.Text = "S6K9Q" Then mBien = True
        Case "7196"
            If txtVariable.Text = "N3P7L" Then mBien = True
        Case "7197"
            If txtVariable.Text = "X3R9H" Then mBien = True
        Case "7198"
            If txtVariable.Text = "J1H9T" Then mBien = True
        Case "7199"
            If txtVariable.Text = "P6G5U" Then mBien = True
    End Select
    Verifico_Clave = mBien
End Function

