VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCommunicator 
   BackColor       =   &H80000004&
   Caption         =   "Chat Interno"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   Picture         =   "frmCommunicator.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      Height          =   375
      Left            =   1800
      MaskColor       =   &H80000010&
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdGrilla 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   25
      Cols            =   1
      FixedCols       =   0
      RowHeightMin    =   290
      BackColor       =   16777215
      ForeColor       =   49152
      BackColorFixed  =   12582912
      ForeColorFixed  =   -2147483635
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483630
      BackColorBkg    =   16777215
      GridColor       =   -2147483634
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function configurogrilla()

    grdGrilla.FormatString = "Usuario|Cod Usuario|IP"
    grdGrilla.ColWidth(0) = 1800 'HORAS
    grdGrilla.ColWidth(0) = 2700 'PACIENTE
    grdGrilla.ColWidth(0) = 3900 'MOTIVO
    
    grdGrilla.Cols = 3
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To grdGrilla.Cols - 1
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &HC00000      'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
End Function
Private Function CargarPersonal()
    sql = "SELECT VEN_CODIGO,VEN_NOMBRE,VEN_IP"
    sql = sql & " FROM VENDEDOR ORDER BY VEN_NOMBRE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    'grdGrilla.Rows = 1
    If rec.EOF = False Then
        For i = 1 To rec.RecordCount
            grdGrilla.TextMatrix(i, 0) = rec!VEN_NOMBRE
            grdGrilla.TextMatrix(i, 1) = rec!VEN_CODIGO
            grdGrilla.TextMatrix(i, 2) = Chk0(rec!VEN_IP)
            rec.MoveNext
            
        Next i
    End If
    rec.Close
End Function

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Centrar_pantalla Me
    configurogrilla
    CargarPersonal
End Sub

Private Sub grdGrilla_Click()
    Shell "C:\Documents and Settings\Administrador\Mis documentos\Visual Basic\CORE\Tests\ws_server\ws_server.exe", vbNormalFocus
End Sub
