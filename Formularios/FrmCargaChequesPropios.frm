VERSION 5.00
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form FrmCargaChequesPropios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Cheques Propios"
   ClientHeight    =   3600
   ClientLeft      =   2535
   ClientTop       =   1005
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCargaChequesPropios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FechaCtl.Fecha TxtCheFecEnt 
      Height          =   315
      Left            =   1305
      TabIndex        =   0
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.TextBox TxtCheImport 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      TabIndex        =   8
      Top             =   2250
      Width           =   1125
   End
   Begin VB.TextBox TxtCheObserv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   60
      TabIndex        =   9
      Top             =   2610
      Width           =   5040
   End
   Begin VB.TextBox TxtCheNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1200
      Width           =   5040
   End
   Begin VB.TextBox TxtCheMotivo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1560
      Width           =   5040
   End
   Begin VB.TextBox TxtCheNumero 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4965
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1380
   End
   Begin VB.ComboBox cboBanco 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   5040
   End
   Begin VB.ComboBox cboCtaBancaria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2100
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5505
      TabIndex        =   13
      Top             =   3165
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   4590
      TabIndex        =   12
      Top             =   3165
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   3675
      TabIndex        =   11
      Top             =   3165
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   3165
      Width           =   900
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   30
      TabIndex        =   14
      Top             =   3030
      Width           =   6405
   End
   Begin FechaCtl.Fecha TxtCheFecEmi 
      Height          =   315
      Left            =   1305
      TabIndex        =   6
      Top             =   1920
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin FechaCtl.Fecha TxtCheFecVto 
      Height          =   315
      Left            =   5205
      TabIndex        =   7
      Top             =   1920
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Pago:"
      Height          =   195
      Index           =   3
      Left            =   4035
      TabIndex        =   25
      Top             =   1980
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Emisión:"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   24
      Top             =   1950
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   23
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   22
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Index           =   6
      Left            =   90
      TabIndex        =   21
      Top             =   2640
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cheque:"
      Height          =   195
      Index           =   7
      Left            =   3960
      TabIndex        =   20
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Responsable:"
      Height          =   195
      Index           =   9
      Left            =   90
      TabIndex        =   19
      Top             =   1245
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Concepto:"
      Height          =   195
      Index           =   10
      Left            =   90
      TabIndex        =   18
      Top             =   1605
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cuenta:"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   885
      Width           =   885
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   3150
      Width           =   660
   End
End
Attribute VB_Name = "FrmCargaChequesPropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As ADODB.Recordset
Dim Rec1 As ADODB.Recordset
Dim sql As String
Dim ImporteCheque As String

Function Validar() As Boolean
   If Trim(TxtCheNumero.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Número de Cheque.", 16, TIT_MSGBOX
        TxtCheNumero.SetFocus
        Exit Function
        
   ElseIf CboBanco.ListIndex = -1 Then
        Validar = False
        MsgBox "Ingrese el Banco.", 16, TIT_MSGBOX
        CboBanco.SetFocus
        Exit Function
                 
   ElseIf cboCtaBancaria.ListIndex = -1 Then
        Validar = False
        MsgBox "Ingrese la Cta Bancaria.", 16, TIT_MSGBOX
        cboCtaBancaria.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheNombre.Text) = "" Then
        Validar = False
        MsgBox "Debe ingresar la Persona responsable.", 16, TIT_MSGBOX
        TxtCheNombre.SetFocus
        Exit Function
   
   ElseIf Trim(TxtCheMotivo.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Concepto del Cheque.", 16, TIT_MSGBOX
        TxtCheMotivo.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheFecEmi.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Fecha de Emisión.", 16, TIT_MSGBOX
        TxtCheFecEmi.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheFecVto.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Fecha de Vencimiento.", 16, TIT_MSGBOX
        TxtCheFecVto.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheImport.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Importe del Cheque.", 16, TIT_MSGBOX
        TxtCheImport.SetFocus
        Exit Function
        
   End If
   
   Validar = True
End Function


Private Sub CboBanco_LostFocus()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Dim MtrObjetos As Variant
        
    If CboBanco.ListIndex <> -1 Then
    
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE_PROPIO " & _
              " WHERE CHEP_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(CboBanco.ItemData(CboBanco.ListIndex))
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then 'EXITE
            Me.TxtCheFecEnt.Text = rec!CHEP_FECENT
            Me.TxtCheNumero.Text = Trim(rec!CHEP_NUMERO)
            
            Me.TxtCheNombre.Text = ChkNull(rec!CHEP_NOMBRE)
            Me.TxtCheMotivo.Text = rec!CHEP_MOTIVO
            Me.TxtCheFecEmi.Text = rec!CHEP_FECEMI
            Me.TxtCheFecVto.Text = rec!CHEP_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!CHEP_IMPORT)
            ImporteCheque = rec!CHEP_IMPORT
            Me.TxtCheObserv.Text = ChkNull(rec!CHEP_OBSERV)
            Call CargoCtaBancaria(CStr(CboBanco.ItemData(CboBanco.ListIndex)))
            Call BuscaProx(Trim(rec!CTA_NROCTA), cboCtaBancaria)
            TxtCheNumero.Enabled = False
            CboBanco.Enabled = False
            MtrObjetos = Array(TxtCheNumero, CboBanco)
            Call CambiarColor(MtrObjetos, 2, &H80000018, "D")
        Else
            
           rec.Close
           Call CargoCtaBancaria(CStr(CboBanco.ItemData(CboBanco.ListIndex)))
           Exit Sub
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub CargoCtaBancaria(Banco As String)
    Set Rec1 = New ADODB.Recordset
    cboCtaBancaria.Clear
    sql = "SELECT CTA_NROCTA FROM CTA_BANCARIA"
    sql = sql & " WHERE BAN_CODINT=" & XN(Banco)
    sql = sql & " AND CTA_FECCIE IS NULL"
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
     Do While Rec1.EOF = False
         cboCtaBancaria.AddItem Trim(Rec1!CTA_NROCTA)
         Rec1.MoveNext
     Loop
     cboCtaBancaria.ListIndex = 0
    End If
    Rec1.Close
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtCheNumero.Text) <> "" And Me.CboBanco.ListIndex <> -1 Then
        resp = MsgBox("¿Seguro desea eliminar el Cheque Nro: " & Trim(Me.TxtCheNumero.Text) & "? ", 36, TIT_MSGBOX)
        If resp <> 6 Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        lblEstado.Caption = "Borrando..."
        DBConn.BeginTrans
        
        'ACTUALIZO EL SALDO DE LA CTA-BANCARIA
'        If ImporteCheque <> "" Then
'            sql = "UPDATE CTA_BANCARIA"
'            sql = sql & " SET CTA_SALACT = CTA_SALACT + " & XN(ImporteCheque)
'            sql = sql & " WHERE"
'            sql = sql & " CTA_NROCTA=" & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
'            sql = sql & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
'            DBConn.Execute sql
'        End If
        
        DBConn.Execute "DELETE FROM CHEQUE_ESTADOS WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.CboBanco.ItemData(CboBanco.ListIndex))
                       
        DBConn.Execute "DELETE FROM CHEQUE WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.CboBanco.ItemData(CboBanco.ListIndex))
        
        Screen.MousePointer = vbNormal
        lblEstado.Caption = ""
        DBConn.CommitTrans
        CmdNuevo_Click
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    If rec.State = 1 Then rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX

End Sub

Private Sub CmdGrabar_Click()
    
  If Validar = True Then
  
    On Error GoTo CLAVOSE
    
    DBConn.BeginTrans
    
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Guardando..."
    Me.Refresh
    
    sql = "SELECT * FROM CHEQUE_PROPIO WHERE CHEP_NUMERO = " & XS(TxtCheNumero.Text)
    sql = sql & " AND BAN_CODINT = " & XN(CboBanco.ItemData(CboBanco.ListIndex))
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = True Then
         sql = "INSERT INTO CHEQUE_PROPIO(CHEP_NUMERO,BAN_CODINT,CHEP_NOMBRE,CHEP_IMPORT,CHEP_FECEMI,"
         sql = sql & "CHEP_FECVTO,CHEP_FECENT,CHEP_MOTIVO,CHEP_OBSERV,CTA_NROCTA)"
         sql = sql & " VALUES (" & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XN(CboBanco.ItemData(CboBanco.ListIndex)) & ","
         sql = sql & XS(Me.TxtCheNombre.Text) & ","
         sql = sql & XN(Me.TxtCheImport.Text) & ","
         sql = sql & XDQ(Me.TxtCheFecEmi.Text) & ","
         sql = sql & XDQ(Me.TxtCheFecVto.Text) & ","
         sql = sql & XDQ(Me.TxtCheFecEnt.Text) & ","
         sql = sql & XS(Me.TxtCheMotivo.Text) & ","
         sql = sql & XS(Me.TxtCheObserv.Text) & ","
         sql = sql & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex)) & ")"
         DBConn.Execute sql
    Else
         sql = "UPDATE CHEQUE_PROPIO SET CHEP_NOMBRE = " & XS(Me.TxtCheNombre.Text)
         sql = sql & ",CHEP_IMPORT = " & XN(Me.TxtCheImport.Text)
         sql = sql & ",CHEP_FECEMI =" & XDQ(Me.TxtCheFecEmi.Text)
         sql = sql & ",CHEP_FECVTO =" & XDQ(Me.TxtCheFecVto.Text)
         sql = sql & ",CHEP_FECENT = " & XDQ(Me.TxtCheFecEnt.Text)
         sql = sql & ",CHEP_MOTIVO = " & XS(Me.TxtCheMotivo.Text)
         sql = sql & ",CHEP_OBSERV = " & XS(Me.TxtCheObserv.Text)
         sql = sql & ",CTA_NROCTA= " & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
         sql = sql & " WHERE CHEP_NUMERO = " & XS(Me.TxtCheNumero.Text)
         sql = sql & " AND BAN_CODINT = " & XN(CboBanco.ItemData(CboBanco.ListIndex))
         DBConn.Execute sql
    End If
    rec.Close
     
    'Insert en la Tabla de Estados de Cheques
    sql = "INSERT INTO CHEQUE_PROPIO_ESTADO (CHEP_NUMERO,BAN_CODINT,ECH_CODIGO,CPES_FECHA,CPES_DESCRI)"
    sql = sql & " VALUES ("
    sql = sql & XS(Me.TxtCheNumero.Text) & ","
    sql = sql & XN(CboBanco.ItemData(CboBanco.ListIndex)) & "," & XN(8) & ","
    sql = sql & XDQ(Date) & ",'CHEQUE LIBRADO')"
    DBConn.Execute sql
    
    'ACTUALIZO EL SALDO DE LA CTA-BANCARIA
'    If ImporteCheque <> "" Then
'        sql = "UPDATE CTA_BANCARIA"
'        sql = sql & " SET CTA_SALACT = CTA_SALACT + " & XN(ImporteCheque)
'        sql = sql & " WHERE"
'        sql = sql & " CTA_NROCTA=" & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
'        sql = sql & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
'        DBConn.Execute sql
'    End If
'        sql = "UPDATE CTA_BANCARIA"
'        sql = sql & " SET CTA_SALACT = CTA_SALACT - " & XN(TxtCheImport)
'        sql = sql & " WHERE"
'        sql = sql & " CTA_NROCTA=" & XS(cboCtaBancaria.List(cboCtaBancaria.ListIndex))
'        sql = sql & " AND BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
'        DBConn.Execute sql
    
    '************* PREGUNTAR POR SI DESEA IMPRIMIR ***************
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
 End If
 Exit Sub
      
CLAVOSE:
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    Me.TxtCheFecEnt.Text = ""
    Me.TxtCheNumero.Enabled = True
    Me.CboBanco.Enabled = True
    cboCtaBancaria.Clear
    Me.TxtCheNombre.Enabled = True
    MtrObjetos = Array(TxtCheNumero, CboBanco)
    Call CambiarColor(MtrObjetos, 2, &H80000005, "E")
    Me.TxtCheNumero.Text = ""
    Me.CboBanco.ListIndex = 0
    Me.TxtCheNombre.Text = ""
    Me.TxtCheMotivo.Text = ""
    Me.TxtCheFecEmi.Text = ""
    Me.TxtCheFecVto.Text = ""
    Me.TxtCheImport.Text = ""
    Me.TxtCheObserv.Text = ""
    ImporteCheque = ""
    Me.TxtCheFecEnt.SetFocus
    'TxtCheNombre.ForeColor = &H80000005
    TxtCheNombre.ForeColor = &H80000008
    lblEstado.Caption = ""
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmCargaChequesPropios = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    TxtCheFecEnt.Text = Date
    lblEstado.Caption = ""
    ImporteCheque = ""
    'CARGO LOS BANCON DONDE TIENEN CUENTAS
    CargoBanco
    cboCtaBancaria.Clear
End Sub

Private Sub CargoBanco()
    sql = "SELECT DISTINCT B.BAN_CODINT, B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            CboBanco.AddItem Trim(rec!BAN_DESCRI)
            CboBanco.ItemData(CboBanco.NewIndex) = Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        CboBanco.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub TxtCheImport_GotFocus()
    SelecTexto TxtCheImport
End Sub

Private Sub TxtCheImport_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtCheImport.Text, KeyAscii)
End Sub

Private Sub TxtCheImport_LostFocus()
   If Trim(TxtCheImport.Text) <> "" Then TxtCheImport.Text = Valido_Importe(TxtCheImport)
    
End Sub

Private Sub TxtCheMotivo_GotFocus()
    SelecTexto TxtCheMotivo
End Sub

Private Sub TxtCheNombre_GotFocus()
    SelecTexto TxtCheNombre
End Sub

Private Sub TxtCheNombre_LostFocus()
   If Me.TxtCheNombre.Text <> "" Then
      Me.TxtCheMotivo.SetFocus
   End If
End Sub

Private Sub TxtCheNumero_GotFocus()
    SelecTexto TxtCheNumero
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheFecEnt_LostFocus()
    If TxtCheFecEnt.Text = "" Then
        TxtCheFecEnt.Text = Format(Date, "dd/mm/yyyy")
        Exit Sub
    End If
    Select Case Len(TxtCheFecEnt.Text)
        Case Is < 2
            MsgBox "Fecha mal Ingresada"
            TxtCheFecEnt.Text = ""
            TxtCheFecEnt.SetFocus
            Exit Sub
        Case Is = 2
            TxtCheFecEnt.Text = Mid(TxtCheFecEnt.Text, 1, 2) & "/" & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
            If IsDate(TxtCheFecEnt.Text) Then
                Exit Sub
            Else
                MsgBox "Fecha Inválida"
                TxtCheFecEnt.Text = ""
                TxtCheFecEnt.SetFocus
                Exit Sub
            End If
        Case Is = 3
            MsgBox "Fecha mal Ingresada"
            TxtCheFecEnt.Text = ""
            TxtCheFecEnt.SetFocus
            Exit Sub
        Case Is = 4
            TxtCheFecEnt.Text = Mid(TxtCheFecEnt.Text, 1, 2) & "/" & Mid(TxtCheFecEnt.Text, 3, 2) & "/" & Format(Year(Date), "0000")
            If IsDate(TxtCheFecEnt.Text) Then
                Exit Sub
            Else
                MsgBox "Fecha Inválida"
                TxtCheFecEnt.Text = ""
                TxtCheFecEnt.SetFocus
                Exit Sub
            End If
        Case Is = 5
            MsgBox "Fecha mal Ingresada"
            TxtCheFecEnt.Text = ""
            TxtCheFecEnt.SetFocus
            Exit Sub
        Case Is = 6
            TxtCheFecEnt.Text = Mid(TxtCheFecEnt.Text, 1, 2) & "/" & Mid(TxtCheFecEnt.Text, 3, 2) & "/" & Mid(Format(Year(Date), "0000"), 1, 2) & Mid(TxtCheFecEnt.Text, 5, 2)
            If IsDate(TxtCheFecEnt.Text) Then
                Exit Sub
            Else
                MsgBox "Fecha Inválida"
                TxtCheFecEnt.Text = ""
                TxtCheFecEnt.SetFocus
                Exit Sub
            End If
        Case Is = 7
            MsgBox "Fecha mal Ingresada"
            TxtCheFecEnt.Text = ""
            TxtCheFecEnt.SetFocus
            Exit Sub
        Case Is = 8
            TxtCheFecEnt.Text = Mid(TxtCheFecEnt.Text, 1, 2) & "/" & Mid(TxtCheFecEnt.Text, 3, 2) & "/" & Mid(TxtCheFecEnt.Text, 5, 4)
            If IsDate(TxtCheFecEnt.Text) Then
                Exit Sub
            Else
                MsgBox "Fecha Inválida"
                TxtCheFecEnt.Text = ""
                TxtCheFecEnt.SetFocus
                Exit Sub
            End If
        Case Is > 10
            MsgBox "Fecha mal Ingresada"
            TxtCheFecEnt.Text = ""
            TxtCheFecEnt.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub TxtCheMotivo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
   If Len(TxtCheNumero.Text) < 10 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 10)
End Sub

Private Sub TxtCheObserv_GotFocus()
    SelecTexto TxtCheObserv
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

