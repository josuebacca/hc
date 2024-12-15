Attribute VB_Name = "Module4"

Public Function CalculoEdad(FecNaci As Date) As Integer
    Dim DiaNac As String
    Dim DiaAct As String
    Dim MesNac As String
    Dim MesAct As String
    Dim AnoNac As String
    Dim AnoAct As String
    Dim FecAct As String
    Dim FecNac As String
    
    FecNac = Format(FecNaci, "DD/MM/yyyy")
    DiaNac = Left(FecNac, 2)
    MesNac = Mid(FecNac, 4, 2)
    AnoNac = Right(FecNac, 4)
    
    FecAct = Format(Date, "DD/MM/yyyy")
    DiaAct = Left(FecAct, 2)
    MesAct = Mid(FecAct, 4, 2)
    AnoAct = Right(FecAct, 4)
    
    CalculoEdad = Val(AnoAct) - Val(AnoNac)
    If Val(MesAct) = Val(MesNac) Then
        If Val(DiaAct) < Val(DiaNac) Then
            CalculoEdad = CalculoEdad - 1
        End If
    ElseIf Val(MesAct) < Val(MesNac) Then
        CalculoEdad = CalculoEdad - 1
    End If
    If CalculoEdad < 1 Then
        Screen.MousePointer = vbNormal
        MsgBox "La edad es menor que uno. Verifique que esté correctamente cargada la fecha de nacimiento", 64, AppName
    End If
End Function

Public Sub CascadaW()
    MainMDI.Arrange 0
End Sub

Public Function Tipo(T As Integer)
    Select Case T
        Case 8
            Tipo = "D"
        Case 2, 3, 4, 5, 6, 7, 11
            Tipo = "N"
        Case 1
            Tipo = "B"
        Case 10, 12
            Tipo = "S"
        Case Else
            Tipo = "S"
            MsgBox "Ocurrió un error en la determinación de un tipo de dato en el Listado. Las condiciones de filtrado pueden no funcionar debidamente.", 48, AppName
    End Select
End Function


Public Sub CallFrm(FrmLlamado As Form, ByVal FrmDestino As Form)
    Set frmDest = FrmDestino
    FrmLlamado.Show
    Set FrmLlamado = Nothing
End Sub

Public Sub ErrManager(ErrNum As Integer)
    Screen.MousePointer = vbNormal
    Select Case ErrNum
        Case 3022
            MsgBox "Ya existe un registro que contiene datos identificatorios únicos. La información está incorrecta o bien ya ha sido ingresada.", 48, AppName
        Case 3061
            MsgBox "Error en la transacción se esperaban más datos que los devueltos.", 48, AppName
        Case 3189
            MsgBox "La operación no puede ser realizada debido a que la tabla esta siendo bloqueada por otro usuario.", 48, AppName
        Case 3200
            MsgBox "No se puede eliminar el registro debido a que esta información está siendo usada en otras áreas del sistema. Elimínela de allí y reintente desde aquí.", 48, AppName
        Case 20500
            MsgBox "No se puede ejecutar el reporte por conflictos de memoria o formato del mismo.", 48, AppName
        Case 3201
            MsgBox "No se puede AGREGAR o ACTUALIZAR este registro porque se necesita que exista un registro en otra tabla padre.", 48, AppName
        Case Else
            MsgBox "Condicion Inesperada en la transacción tipo (" & Err.Number & ") " & Err.Description & ".", 48, AppName
    End Select
    Screen.MousePointer = vbNormal
End Sub

    
Public Function TraeUltNum(Tabla As String)
    Dim Result As Integer
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    sSQL = "SELECT NUMERO FROM NUMEROS WHERE TABLA=" & XS(Tabla)
    Result = QRY(sSQL, snp)
    Select Case Result
        Case 1
            TraeUltNum = snp.Fields(0)
        Case -1, -2, 0
            TraeUltNum = Result
    End Select
    snp.Close
    Set snp = Nothing
End Function

Public Function ActUltNum(Tabla As String)
    Dim Result As Integer
    Dim snp As ADODB.Recordset
    Dim sSQL As String
    On Error GoTo ErrorTran
    sSQL = "UPDATE NUMEROS SET NUMERO=NUMERO+1 WHERE TABLA=" & XS(Tabla)
    DBConn.Execute sSQL
    ActUltNum = 1
    Exit Function
ErrorTran:
    ActUltNum = -1
End Function


Public Function XDO_REP(Valor As Variant)
    If Valor = "__/__/____" Then
        XDO_REP = "Null"
    Else
        XDO_REP = "DATE("
        XDO_REP = XDO_REP + Mid(Valor, 7, 4)
        XDO_REP = XDO_REP & ","
        XDO_REP = XDO_REP + Mid(Valor, 4, 2)
        XDO_REP = XDO_REP & ","
        XDO_REP = XDO_REP + Mid(Valor, 1, 2)
        XDO_REP = XDO_REP & ")"
    End If
End Function

Public Function Agregando() As Boolean
    If Frm.Modo.BackColor = QBColor(14) Then
        Agregando = True
    Else
        Agregando = False
    End If
End Function

Public Sub MaskFecha(ByRef Ctrl As TextBox)
    Dim Dig(20) As String
    Fec = Trim(Ctrl.Text)
    Num = Len(Fec)
    DIA = ""
    MES = ""
    año = ""
    Fecha = ""
    If Num = 0 Then Exit Sub
    For i = 1 To Num
        Dig(i) = Mid(Fec, i, 1)
        If Not IsNumeric(Dig(i)) Then
            Dig(i) = "/"
        End If
        Fecha = Fecha + Dig(i)
    Next
    If IsNumeric(Fec) = True Then
        If Num = 7 Or Num > 8 Then
            flag = "Error"
        ElseIf Num = 4 Then
            Dato = Left(Fec, 2) & "/"
            Dato = Dato + Right(Fec, 2)
        ElseIf Num = 6 Then
            Dato = Left(Fec, 2) & "/"
            Dato = Dato + Mid(Fec, 3, 2) & "/"
            Dato = Dato + Right(Fec, 2)
        ElseIf Num = 8 Then
            Dato = Left(Fec, 2) & "/"
            Dato = Dato + Mid(Fec, 3, 2) & "/"
            Dato = Dato + Right(Fec, 4)
        ElseIf Num < 6 Then
            flag = "Error"
        Else
            Dato = Fec
        End If
    Else
            Dato = Fec
    End If
    Ctrl.Text = Format(Dato, "dd/mm/yyyy")
    If IsDate(Ctrl.Text) = False Then
        flag = "Error"
    End If
    If flag = "Error" Then
        Ctrl.Text = ""
        MsgBox ("ERROR EN FECHA, ingrese un formato válido." & Chr(13) & "Ej: Para 3 de Agosto de 1999 ingrese 030899")
        Ctrl.SetFocus
    End If
End Sub


Public Function MaskHor(Ctrl As TextBox)
    hor = Ctrl.Text
    Num = Len(hor)
    If Num = 0 Then Exit Function
    If IsNumeric(hor) = True Then
        If Num > 4 Then
            flag = "Error"
        ElseIf Num < 3 Then
            Dato = hor & ":00"
        ElseIf Num = 4 Then
            Dato = Left(hor, 2) & ":"
            Dato = Dato + Right(hor, 2)
        ElseIf Num = 3 Then
            Dato = Left(hor, 1) & ":"
            Dato = Dato + Right(hor, 2)
        Else
            Dato = hor
        End If
    Else
            Dato = hor
    End If
    MaskHor = Format(Dato, "hh:mm")
    If IsDate(MaskHor) = False Then
        flag = "Error"
    End If
    If flag = "Error" Then
        MaskHor = ""
        MsgBox ("Error en hora, ingrese un formato válido." & Chr(13) & "Ej: Para 01:09 ingrese 109 o 0109 o 01:09")
        Ctrl.SetFocus
    End If
End Function

Public Sub Hora()
    Dim Frm As Form
    On Error GoTo nada
    For Each Frm In Forms
        If Frm.Name <> "MainMDI" Then
            Frm.Info.Caption = Format(Time, "HH:MM") & "-" & Format(Date, "dd/mm/yyyy")
        End If
    Next
    Exit Sub
nada:
End Sub


Public Sub MSJ(ByRef Frm As Form, ByRef Mensaje As String)
    Frm.BE.Caption = " " & Mensaje
    Frm.BE.Refresh
End Sub


Public Function QRY(Cad As String, ByRef snp As ADODB.Recordset) As Integer
On Error GoTo ErroCom
    Set snp = New ADODB.Recordset
    snp.Open Cad$, DBConn, adOpenStatic, adLockOptimistic
    QRY = snp.RecordCount
Exit Function
ErroCom:
    QRY = -2
    ErrManager Err
End Function




Public Function NoNum(Ctrl As TextBox, Nombre As String)
  NoNum = False
  If Trim(Ctrl.Text) <> "" Then
    If Not IsNumeric(Ctrl.Text) Or Val(Ctrl.Text) < 0 Then
      Screen.MousePointer = vbHourglass
      MsgBox Nombre & " debe se numerico positivo", 48
      Ctrl = ""
      Screen.MousePointer = vbNormal
      Ctrl.SetFocus
      NoNum = True
      Resu_Vali = True
    End If
  End If
End Function
Public Function NoVacio(Ctrl As TextBox, Nombre As String)
  NoVacio = False
  If Trim(Ctrl.Text) = "" Then
    Screen.MousePointer = vbNormal
    MsgBox Nombre & " no puede estar vacío", 48
    Ctrl.SetFocus
    NoVacio = True
  End If
End Function


Public Sub BTActualizar(ByRef Frm As Form)
    If Frm.cActualizar.Visible Then
       Frm.cNuevo.Enabled = True
       Frm.cAgregar.Enabled = False
       Frm.cEliminar.Enabled = True
       Frm.cActualizar.Enabled = True
       Frm.cImprimir.Enabled = True
       MSJ Frm, "Modo Actualizar/Eliminar"
       Frm.Modo.BackColor = QBColor(12)
    End If
End Sub
Public Sub BTActualizarR(ByRef Frm As Form)
'       frm.cActualizar.Visible
'       frm.cNuevo.Enabled = True
'       frm.cAgregar.Enabled = False
'       frm.cEliminar.Enabled = True
'       frm.cActualizar.Enabled = True
'       frm.cImprimir.Enabled = True
       MSJ Frm, "Modo Actualizar/Eliminar"
       Frm.Modo.BackColor = QBColor(12)
End Sub
Public Sub BTConsultar(ByRef Frm As Form)
   MSJ Frm, "Modo Consulta"
   Frm.cAgregar.Enabled = False
   Frm.cActualizar.Enabled = False
   Frm.cEliminar.Enabled = False
   Frm.cImprimir.Enabled = True
   Frm.Modo.BackColor = QBColor(10)
End Sub
Public Sub BTConsultarR(ByRef Frm As Form)
   MSJ Frm, "Modo Consulta"
'   frm.cAgregar.Enabled = False
'   frm.cActualizar.Enabled = False
'   frm.cEliminar.Enabled = False
'   frm.cImprimir.Enabled = True
   Frm.Modo.BackColor = QBColor(10)
End Sub

Public Sub BTAgregar(Frm As Form)
If Frm.cAgregar.Visible Then
  Frm.cAgregar.Enabled = True
  Frm.cEliminar.Enabled = False
  Frm.cActualizar.Enabled = False
  Frm.cImprimir.Enabled = True
  MSJ Frm, "Modo Agregar"
  Frm.Modo.BackColor = QBColor(14)
End If
End Sub
Public Sub BTAgregarR(Frm As Form)
'  frm.cAgregar.Visible
'  frm.cAgregar.Enabled = True
'  frm.cEliminar.Enabled = False
'  frm.cActualizar.Enabled = False
'  frm.cImprimir.Enabled = True
  MSJ Frm, "Modo Agregar"
  Frm.Modo.BackColor = QBColor(14)
End Sub
Public Sub BTNada(ByRef Frm As Form)
  Frm.cNuevo.Enabled = False
  Frm.cAgregar.Enabled = False
  Frm.cEliminar.Enabled = False
  Frm.cActualizar.Enabled = False
  Frm.cImprimir.Enabled = False
End Sub

Public Function DateNull(Valor As Variant) As String
    If IsNull(Valor) Then
        DateNull = "__/__/____"
    Else
        DateNull = Format(Valor, "dd/mm/yyyy")
    End If
End Function

Public Function Sentido(XX As String) As String
    If XX = "é" Then
        Sentido = "+"
    Else
        Sentido = "-"
    End If
End Function

Public Function SentidoSQL(XX As String) As String
    If XX = "é" Then
        SentidoSQL = "ASC"
    Else
        SentidoSQL = "DESC"
    End If
End Function

Public Function XDT(Valor As Variant) As Date
    If Valor = "__/__/____" Then
        XDT = "Null"
    Else
        XDT = "'" & Format(Date, "mm/dd/yyyy hh:mm:ss") & "'"
    End If
End Function

Public Function XDO(Valor As Variant)
    If Valor = "__/__/____" Then
        XDO = "Null"
    Else
        XDO = " TO_DATE('" & Valor & "','DD/MM/YYYY')"
    End If
End Function






