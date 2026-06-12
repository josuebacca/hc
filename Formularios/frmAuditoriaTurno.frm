VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAuditoriaTurno 
   Caption         =   "Historial del turno"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   16080
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdHistorial 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Turno"
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   15615
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   6600
         TabIndex        =   5
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lblDoctor 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblPaciente 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label lblFecha 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAuditoriaTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Variable privada del formulario (en lugar de las Public)
Private mIdTurno As Long

Public Sub CargarDatos(ByVal nIdTurno As Long, ByVal sNombrePaciente As String, ByVal sFechaTurno As String, ByVal sDoctor As String)
    mIdTurno = nIdTurno
    lblPaciente.Caption = sNombrePaciente
    lblFecha.Caption = sFechaTurno
    lblDoctor.Caption = sDoctor
    CargarHistorial
End Sub


Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargarHistorial
End Sub

Private Sub CargarHistorial()
    Dim rec As ADODB.Recordset
    Dim Fila As Integer
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    
    ' Setup de columnas de la grilla
    grdHistorial.FormatString = "^Horario|<Importe|<Motivo|<Orden|<Estado|<Usuario|<Acción|<Fecha Acción|<Llave"
    
    grdHistorial.Font.Size = 10  ' default suele ser 8

    grdHistorial.ColWidth(0) = 1200 'HORARIO
    grdHistorial.ColWidth(1) = 1600 'IMPORTE
    grdHistorial.ColWidth(2) = 3400 'MOTIVO
    grdHistorial.ColWidth(3) = 800 'ORDEN
    grdHistorial.ColWidth(4) = 1100 'Estado
    grdHistorial.ColWidth(5) = 2200 'USUARIO
    grdHistorial.ColWidth(6) = 1500 'ACCION
    grdHistorial.ColWidth(7) = 1800 'FECHA ACCION
    grdHistorial.ColWidth(8) = 1400 'LLAVE
    
    grdHistorial.Cols = 9
    grdHistorial.rows = 1
    grdHistorial.BorderStyle = flexBorderNone

    grdHistorial.row = 0
    For i = 0 To grdHistorial.Cols - 1
        grdHistorial.Col = i
        grdHistorial.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdHistorial.CellBackColor = &H808080    'GRIS OSCURO
        grdHistorial.CellFontBold = True
    Next
    
    
    ' Armo el SQL con JOIN a LLAVE_USUARIO y descripción de acción con CASE
    sql = "SELECT"
    sql = sql & "  HT.HORA_DESDE,"
    sql = sql & "  HT.HORA_HASTA,"
    sql = sql & "  HT.IMPORTE,"
    sql = sql & "  AT.ACCION_DESCRI,"
    sql = sql & "  HT.ACCION_FECHA,"
    sql = sql & "  LU.LLA_USUARIO,"
    sql = sql & "  HT.ORDEN,"
    sql = sql & "  HT.MOTIVO,"
    sql = sql & "  ET.ESTADO_DESCRI,"
    sql = sql & "  HT.USU_NOMBRE"
    sql = sql & " FROM HISTORICO_TURNO HT"
    sql = sql & " LEFT JOIN LLAVE_USUARIO LU ON LU.LLA_CODIGO = HT.LLA_CODIGO"
    sql = sql & " LEFT JOIN ACCION_TURNO AT ON AT.ACCION_CODIGO = HT.ACCION_CODIGO"
    sql = sql & " LEFT JOIN ESTADO_TURNO ET ON ET.ESTADO_CODIGO = HT.ESTADO_CODIGO"
    sql = sql & " WHERE HT.ID_TURNO = " & mIdTurno
    sql = sql & " ORDER BY HT.ACCION_FECHA DESC"
    
    Set rec = New ADODB.Recordset
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Not rec.EOF Then
        Do While Not rec.EOF
            grdHistorial.AddItem ""
            Fila = grdHistorial.rows - 1
            
            ' Horario: muestro solo la parte de la hora de los campos datetime
            sHoraDesde = ""
            sHoraHasta = ""
            If Not IsNull(rec!HORA_DESDE) Then
                sHoraDesde = Format(CDate(rec!HORA_DESDE), "HH:MM")
            End If
            If Not IsNull(rec!HORA_HASTA) Then
                sHoraHasta = Format(CDate(rec!HORA_HASTA), "HH:MM")
            End If
            grdHistorial.TextMatrix(Fila, 0) = sHoraDesde & " - " & sHoraHasta
            
            ' Importe
            If Not IsNull(rec!Importe) Then
                grdHistorial.TextMatrix(Fila, 1) = Format(rec!Importe, "0.00")
            End If
            
            ' Motivo
            If Not IsNull(rec!ACCION_DESCRI) Then
                grdHistorial.TextMatrix(Fila, 2) = ChkNull(rec!Motivo)
            End If
            
            ' Orden
            If Not IsNull(rec!ACCION_DESCRI) Then
                grdHistorial.TextMatrix(Fila, 3) = ChkNull(rec!orden)
            End If
            
            ' Estado
            If Not IsNull(rec!ESTADO_DESCRI) Then
                grdHistorial.TextMatrix(Fila, 4) = ChkNull(rec!ESTADO_DESCRI)
            End If
            
            ' Usuario
            If Not IsNull(rec!USU_NOMBRE) Then
                grdHistorial.TextMatrix(Fila, 5) = ChkNull(rec!USU_NOMBRE)
            End If
            
            ' Acción
            If Not IsNull(rec!ACCION_DESCRI) Then
                grdHistorial.TextMatrix(Fila, 6) = rec!ACCION_DESCRI
            End If
            
            ' Fecha y hora de la acción
            If Not IsNull(rec!ACCION_FECHA) Then
                grdHistorial.TextMatrix(Fila, 7) = Format(CDate(rec!ACCION_FECHA), "DD/MM/YYYY HH:MM")
            End If
            
            ' Llave
            If Not IsNull(rec!LLA_USUARIO) Then
                grdHistorial.TextMatrix(Fila, 8) = rec!LLA_USUARIO
            End If
            
            rec.MoveNext
        Loop
    Else
        MsgBox "No hay historial para este turno.", vbInformation, TIT_MSGBOX
    End If
    
    rec.Close
    Set rec = Nothing
End Sub

