VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmControlVisitas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Visitas"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9585
   Icon            =   "frmControlVisitas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   120
      Picture         =   "frmControlVisitas.frx":1CFA
      ScaleHeight     =   375
      ScaleWidth      =   585
      TabIndex        =   2
      Top             =   70
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   0
      Top             =   4800
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   8017
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   16777215
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   400
      Left            =   730
      ScaleHeight     =   375
      ScaleWidth      =   8730
      TabIndex        =   1
      Top             =   70
      Width           =   8760
   End
End
Attribute VB_Name = "frmControlVisitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Letrero As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Call Centrar_pantalla(Me)
    configurogrilla
    configurodia Date
    buscarcumples Date
End Sub
Private Function configurodia(Fecha As Date) As String
    
    Dim DIA As Integer
    DIA = Weekday(Fecha, vbMonday)
    Letrero = "Control de Visitas del día " & WeekdayName(DIA, False) & " " & day(Fecha) & " de " & MonthName(Month(Fecha), False) & " de " & Year(Fecha)
End Function

Private Function configurogrilla()
'GRILLA BUSQUEDA
    GrdModulos.FormatString = "Fecha|Paciente|Motivo|Telefono|CodPaciente|mail|Doctor"
    GrdModulos.ColWidth(0) = 1200 'FECHA
    GrdModulos.ColWidth(1) = 2400 'Paciente
    GrdModulos.ColWidth(2) = 2400 'Motivo
    GrdModulos.ColWidth(3) = 1500 'Telefono
    GrdModulos.ColWidth(4) = 0 'COD PACIENTE
    GrdModulos.ColWidth(5) = 0 'MAIL PACIENTE
    GrdModulos.ColWidth(6) = 2000 'doctor
    
    GrdModulos.Cols = 7
    GrdModulos.Rows = 1
    GrdModulos.HighLight = flexHighlightNever
    GrdModulos.BorderStyle = flexBorderNone
    GrdModulos.row = 0
    For i = 0 To GrdModulos.Cols - 1
        GrdModulos.Col = i
        GrdModulos.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        GrdModulos.CellBackColor = &H808080    'GRIS OSCURO
        GrdModulos.CellFontBold = True
    Next
    
End Function
Private Function buscarcumples(hoy As Date)
    sql = "SELECT CL.CCL_FECPC,P.CLI_RAZSOC,P.CLI_TELEFONO,CL.CLI_CODIGO,P.CLI_MAIL, "
    sql = sql & " T.TR_DESCRI,CL.CCL_INDICA,V.VEN_NOMBRE"
    sql = sql & " FROM CLIENTE P, CCLINICO CL, TRATAMIENTO T,VENDEDOR V"
    sql = sql & " WHERE P.CLI_CODIGO = CL.CLI_CODIGO "
    sql = sql & " AND T.TR_CODIGO = CL.TR_CODIGO "
    sql = sql & " AND V.VEN_CODIGO = CL.VEN_CODIGO "
    sql = sql & " AND CL.CCL_FECPC = " & XDQ(Date)
   

    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            
            GrdModulos.AddItem rec!CCL_FECPC & Chr(9) & rec!CLI_RAZSOC & Chr(9) & _
                               rec!CCL_INDICA & Chr(9) & rec!CLI_TELEFONO & Chr(9) & _
                               rec!CLI_CODIGO & Chr(9) & rec!CLI_MAIL & Chr(9) & rec!VEN_NOMBRE

            rec.MoveNext
        Loop
    End If
    rec.Close
End Function

Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
    
    Static Anterior As Boolean
    Static tamañoLetrero As Single
    Static X As Single
    If Not Anterior Then
        tamañoLetrero = Picture1.TextWidth(Letrero)
        Anterior = True
        X = Picture1.ScaleWidth
    End If
    Picture1.Cls
    Picture1.CurrentX = X
    Picture1.CurrentY = 0
'Para cambiar el tipo de letra
    Picture1.FontName = "Arial"
    Picture1.FontBold = True
    Picture1.Print Letrero
    X = X - 30
    If X < -tamañoLetrero Then X = Picture1.ScaleWidth
End Sub
