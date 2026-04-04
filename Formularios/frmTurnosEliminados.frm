VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTurnosEliminados 
   Caption         =   "Turnos Eliminados"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   17535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      DragIcon        =   "frmTurnosEliminados.frx":0000
      Height          =   375
      Left            =   5760
      Picture         =   "frmTurnosEliminados.frx":628A
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid grdTurnos 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   151388161
      CurrentDate     =   46104
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   151388161
      CurrentDate     =   46104
   End
   Begin VB.Label lblFIltro 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmTurnosEliminados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dictLlaves As Object
Private Function ObtenerNombreUsuario(Codigo As Variant) As String

    Dim obj As ClsLLave

    If IsNull(Codigo) Or Codigo = "" Then
        ObtenerNombreUsuario = "-"
        Exit Function
    End If

    If dictLlaves.Exists(CLng(Codigo)) Then
        Set obj = dictLlaves(CLng(Codigo))
        ObtenerNombreUsuario = obj.USUARIO
    Else
        ObtenerNombreUsuario = "-"
    End If

End Function
Private Sub ValidarYCargar()

    If dtpDesde.Value > dtpHasta.Value Then
        MsgBox "La fecha desde no puede ser mayor a la fecha hasta.", vbExclamation
        Exit Sub
    End If

    Call CargarGrilla

End Sub
Private Sub cmdActualizar_Click()
    Call ValidarYCargar
End Sub
Public Sub CargarLlavesUsuarios()

    Dim obj As ClsLLave
    Dim sql As String

    Set dictLlaves = CreateObject("Scripting.Dictionary")

    sql = "SELECT LLA_CODIGO, LLA_VALOR, LLA_USUARIO FROM LLAVE_USUARIO"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic

    Do While Not Rec1.EOF

        Set obj = New ClsLLave

        obj.Codigo = Rec1!LLA_CODIGO
        obj.Valor = Trim(CStr(Rec1!LLA_VALOR))
        obj.USUARIO = Rec1!LLA_USUARIO

        dictLlaves.Add obj.Codigo, obj

        Rec1.MoveNext
    Loop

    Rec1.Close

End Sub
Private Sub Form_Load()

    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    ' Fechas por defecto
    dtpDesde.Value = DateAdd("d", -7, Date)
    dtpHasta.Value = Date + 1
    
    CargarLlavesUsuarios
    Call CargarGrilla

End Sub


Private Sub CmdBuscar_Click()
    Call CargarGrilla
End Sub


Private Sub CargarGrilla()

    Dim sql As String
    Dim FechaDesde As String
    Dim FechaHasta As String

    Dim Fila As Long
    Dim obj As ClsLLave

    ' Formateo fechas seguro
    FechaDesde = "'" & Format(dtpDesde.Value, "dd/mm/yyyy") & "'"
    FechaHasta = "'" & Format(dtpHasta.Value, "dd/mm/yyyy") & "'"

    ' Query
    sql = "SELECT TOP 200 " & _
          "C.CLI_NRODOC, " & _
          "C.CLI_RAZSOC, " & _
          "T.TUR_FECHA, " & _
          "T.TUR_HORAD, " & _
          "T.TUR_HORAH, " & _
          "V.VEN_NOMBRE, " & _
          "T.CREATED_AT, T.CREADO_POR, " & _
          "T.UPDATED_AT, T.ACTUALIZADO_POR, " & _
          "T.DELETED_AT, T.BORRADO_POR " & _
          "FROM TURNOS T " & _
          "INNER JOIN CLIENTE C ON T.CLI_CODIGO = C.CLI_CODIGO " & _
          "INNER JOIN VENDEDOR V ON T.VEN_CODIGO = V.VEN_CODIGO " & _
          "WHERE T.DELETED_AT IS NOT NULL " & _
          "AND T.TUR_FECHA BETWEEN " & FechaDesde & " AND " & FechaHasta & " " & _
          "ORDER BY T.DELETED_AT DESC"

    'rec.Close
    rec.Open sql, DBConn, adOpenStatic, adLockReadOnly

    ' Configurar grilla
    grdTurnos.rows = 1
    grdTurnos.Cols = 12

    grdTurnos.TextMatrix(0, 0) = "DNI"
    grdTurnos.TextMatrix(0, 1) = "Paciente"
    grdTurnos.TextMatrix(0, 2) = "Fecha Turno"
    grdTurnos.TextMatrix(0, 3) = "Desde"
    grdTurnos.TextMatrix(0, 4) = "Hasta"
    grdTurnos.TextMatrix(0, 5) = "Doctor"
    grdTurnos.TextMatrix(0, 6) = "Creado"
    grdTurnos.TextMatrix(0, 7) = "Creado por"
    grdTurnos.TextMatrix(0, 8) = "Ult. actualización"
    grdTurnos.TextMatrix(0, 9) = "Actualizado por"
    grdTurnos.TextMatrix(0, 10) = "Eliminado"
    grdTurnos.TextMatrix(0, 11) = "Eliminado por"
    
    grdTurnos.ColWidth(0) = 1000   ' DNI
    grdTurnos.ColWidth(1) = 3000   ' Nombre paciente
    grdTurnos.ColWidth(2) = 1200   ' Fecha turno
    grdTurnos.ColWidth(3) = 800   ' Hora desde
    grdTurnos.ColWidth(4) = 800   ' Hora hasta
    grdTurnos.ColWidth(5) = 2000   ' Doctor
    grdTurnos.ColWidth(6) = 1500   ' Created At
    grdTurnos.ColWidth(7) = 1100   ' Creado por
    grdTurnos.ColWidth(8) = 1800   ' Updated At
    grdTurnos.ColWidth(9) = 1200   ' Actualizado por
    grdTurnos.ColWidth(10) = 1500  ' Deleted At
    grdTurnos.ColWidth(11) = 1200  ' Eliminado por

    Fila = 1

    Do While Not rec.EOF

        grdTurnos.rows = grdTurnos.rows + 1

        grdTurnos.TextMatrix(Fila, 0) = rec!CLI_NRODOC
        grdTurnos.TextMatrix(Fila, 1) = rec!CLI_RAZSOC
        grdTurnos.TextMatrix(Fila, 2) = Format(rec!TUR_FECHA, "dd/mm/yyyy")
        grdTurnos.TextMatrix(Fila, 3) = Format(rec!TUR_HORAD, "hh:nn")
        grdTurnos.TextMatrix(Fila, 4) = Format(rec!TUR_HORAH, "hh:nn")
        grdTurnos.TextMatrix(Fila, 5) = rec!VEN_NOMBRE

        ' CREATED
        grdTurnos.TextMatrix(Fila, 6) = IIf(IsNull(rec!CREATED_AT), "-", Format(rec!CREATED_AT, "dd/mm/yyyy hh:nn"))
        grdTurnos.TextMatrix(Fila, 7) = ObtenerNombreUsuario(rec!CREADO_POR)

        ' UPDATED
        grdTurnos.TextMatrix(Fila, 8) = IIf(IsNull(rec!UPDATED_AT), "-", Format(rec!UPDATED_AT, "dd/mm/yyyy hh:nn"))
        grdTurnos.TextMatrix(Fila, 9) = ObtenerNombreUsuario(rec!ACTUALIZADO_POR)

        ' DELETED
        grdTurnos.TextMatrix(Fila, 10) = IIf(IsNull(rec!DELETED_AT), "-", Format(rec!DELETED_AT, "dd/mm/yyyy hh:nn"))
        grdTurnos.TextMatrix(Fila, 11) = ObtenerNombreUsuario(rec!BORRADO_POR)

        Fila = Fila + 1
        rec.MoveNext

    Loop
    
    rec.Close
    
    lblFIltro.Caption = "Mostrando eliminados con fecha de turno desde " & _
    Format(dtpDesde.Value, "dd/mm/yyyy") & _
    " hasta " & Format(dtpHasta.Value, "dd/mm/yyyy")

End Sub

