VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAuditoriaTurno 
   Caption         =   "Historial del turno"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdHistorial 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Turno"
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13695
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   495
         Left            =   5760
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
         Left            =   9960
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
    Dim fila As Integer
    Dim sHoraDesde As String
    Dim sHoraHasta As String
    
    ' Setup de columnas de la grilla
    grdHistorial.FormatString = "^Horario|<Importe|<Acción|<Fecha Acción|<Usuario"
    
    grdHistorial.Font.Size = 10  ' default suele ser 8

    grdHistorial.ColWidth(0) = 1200 'HORAS
    grdHistorial.ColWidth(1) = 2000 'PACIENTE
    grdHistorial.ColWidth(2) = 2000 'EDAD
    grdHistorial.ColWidth(3) = 2000 'CELULAR/TELEFONO
    grdHistorial.ColWidth(4) = 1500 'CELULAR
    
    grdHistorial.Cols = 5
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
    sql = sql & "  LU.LLA_USUARIO"
    sql = sql & " FROM HISTORICO_TURNO HT"
    sql = sql & " LEFT JOIN LLAVE_USUARIO LU ON LU.LLA_CODIGO = HT.LLA_CODIGO"
    sql = sql & " LEFT JOIN ACCION_TURNO AT ON AT.ACCION_CODIGO = HT.ACCION_CODIGO"
    sql = sql & " WHERE HT.ID_TURNO = " & mIdTurno
    sql = sql & " ORDER BY HT.ACCION_FECHA DESC"
    
    Set rec = New ADODB.Recordset
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If Not rec.EOF Then
        Do While Not rec.EOF
            grdHistorial.AddItem ""
            fila = grdHistorial.rows - 1
            
            ' Horario: muestro solo la parte de la hora de los campos datetime
            sHoraDesde = ""
            sHoraHasta = ""
            If Not IsNull(rec!HORA_DESDE) Then
                sHoraDesde = Format(CDate(rec!HORA_DESDE), "HH:MM")
            End If
            If Not IsNull(rec!HORA_HASTA) Then
                sHoraHasta = Format(CDate(rec!HORA_HASTA), "HH:MM")
            End If
            grdHistorial.TextMatrix(fila, 0) = sHoraDesde & " - " & sHoraHasta
            
            ' Importe
            If Not IsNull(rec!Importe) Then
                grdHistorial.TextMatrix(fila, 1) = Format(rec!Importe, "0.00")
            End If
            
            ' Acción
            If Not IsNull(rec!ACCION_DESCRI) Then
                grdHistorial.TextMatrix(fila, 2) = rec!ACCION_DESCRI
            End If
            
            ' Fecha y hora de la acción
            If Not IsNull(rec!ACCION_FECHA) Then
                grdHistorial.TextMatrix(fila, 3) = Format(CDate(rec!ACCION_FECHA), "DD/MM/YYYY HH:MM")
            End If
            
            ' Usuario
            If Not IsNull(rec!LLA_USUARIO) Then
                grdHistorial.TextMatrix(fila, 4) = rec!LLA_USUARIO
            End If
            
            rec.MoveNext
        Loop
    Else
        MsgBox "No hay historial para este turno.", vbInformation, TIT_MSGBOX
    End If
    
    rec.Close
    Set rec = Nothing
End Sub

