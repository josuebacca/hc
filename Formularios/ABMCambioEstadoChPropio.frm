VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form ABMCambioEstadoChPropio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Estado de Cheques Propios"
   ClientHeight    =   5295
   ClientLeft      =   2280
   ClientTop       =   435
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMCambioEstadoChPropio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7245
   Begin FechaCtl.Fecha TxtCheFecEmi 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   855
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
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
      Height          =   120
      Left            =   135
      TabIndex        =   22
      Top             =   4635
      Width           =   6945
   End
   Begin VB.ComboBox cboBanco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   5040
   End
   Begin VB.TextBox TxtCheImport 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   1200
      Width           =   1125
   End
   Begin VB.TextBox TxtCheNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6150
      TabIndex        =   10
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4290
      TabIndex        =   8
      Top             =   4860
      Width           =   915
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   135
      TabIndex        =   11
      Top             =   1515
      Width           =   6945
   End
   Begin VB.TextBox TxtCheObserv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   135
      TabIndex        =   7
      Top             =   3945
      Width           =   6945
   End
   Begin VB.ComboBox CboEstado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3900
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3315
      Width           =   3180
   End
   Begin MSFlexGridLib.MSFlexGrid Grd1 
      Height          =   1500
      Left            =   105
      TabIndex        =   12
      Top             =   1695
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   2646
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483624
      AllowBigSelection=   -1  'True
      Enabled         =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
   End
   Begin FechaCtl.Fecha TxtCheFecVto 
      Height          =   315
      Left            =   5295
      TabIndex        =   3
      Top             =   855
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin FechaCtl.Fecha TxtCesFecha 
      Height          =   315
      Left            =   1995
      TabIndex        =   5
      Top             =   3315
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   165
      TabIndex        =   21
      Top             =   535
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Emisión:"
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   20
      Top             =   890
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Vencimiento:"
      Height          =   195
      Index           =   5
      Left            =   3810
      TabIndex        =   19
      Top             =   885
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   18
      Top             =   1260
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cheque:"
      Height          =   195
      Index           =   7
      Left            =   165
      TabIndex        =   17
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   16
      Top             =   3720
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Cambio de Estado:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3345
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      Height          =   195
      Index           =   0
      Left            =   3165
      TabIndex        =   14
      Top             =   3345
      Width           =   690
      WordWrap        =   -1  'True
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
      Left            =   180
      TabIndex        =   13
      Top             =   4875
      Width           =   660
   End
End
Attribute VB_Name = "ABMCambioEstadoChPropio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboBanco_LostFocus()
    Dim MtrObjetos As Variant

    If cboBanco.ListIndex <> -1 Then
        lblEstado.Caption = "Buscando..."
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE_PROPIO "
        sql = sql & " WHERE CHEP_NUMERO = " & XS(TxtCheNumero.Text)
        sql = sql & " AND BAN_CODINT = " & XN(cboBanco.ItemData(cboBanco.ListIndex))
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If rec.EOF = False Then 'EXITE
            Me.TxtCheNumero.Text = Trim(rec!CHEP_NUMERO)
            Me.TxtCheFecEmi.Text = rec!CHEP_FECEMI
            Me.TxtCheFecVto.Text = rec!CHEP_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!CHEP_IMPORT)
            TxtCheNumero.Enabled = False
            cboBanco.Enabled = False
            MtrObjetos = Array(TxtCheNumero, cboBanco)
            Call CambiarColor(MtrObjetos, 2, &H80000018, "D")
            'CARGO GRILLA
            sql = "SELECT CPES_FECHA,ECH_DESCRI,CPES_DESCRI"
            sql = sql & " FROM CHEQUE_PROPIO_ESTADO CE, ESTADO_CHEQUE EC"
            sql = sql & " WHERE CE.ECH_CODIGO=EC.ECH_CODIGO"
            sql = sql & " AND CE.CHEP_NUMERO=" & XS(TxtCheNumero.Text)
            sql = sql & " AND CE.BAN_CODINT=" & XN(cboBanco.ItemData(cboBanco.ListIndex))
            sql = sql & " ORDER BY CPES_FECHA"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.EOF = False Then
               Rec1.MoveFirst
               Do While Not Rec1.EOF
                 Grd1.AddItem Rec1!CPES_FECHA & Chr(9) & Trim(Rec1.Fields(1)) & Chr(9) & Trim(Rec1.Fields(2))
                 Rec1.MoveNext
               Loop
            End If
            Rec1.Close
            Me.TxtCesFecha.SetFocus
        Else
           lblEstado.Caption = ""
           MsgBox "El Cheque no Existe", vbExclamation, TIT_MSGBOX
           rec.Close
        End If
        lblEstado.Caption = ""
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub CmdGrabar_Click()
   
 If Me.ActiveControl.Name <> "CmdNuevo" And Me.ActiveControl.Name <> "CmdSalir" Then
    lblEstado.Caption = "Actualizando..."
    'Verifico que NO graben dos veces el mismo estado en el mismo día
    sql = "SELECT ECH_CODIGO, MAX(CPES_FECHA)as maximo"
    sql = sql & " FROM CHEQUE_PROPIO_ESTADO"
    sql = sql & " WHERE CHEP_NUMERO = " & XS(Me.TxtCheNumero.Text)
    sql = sql & " AND ECH_CODIGO = " & XN(CboEstado.ItemData(CboEstado.ListIndex))
    sql = sql & " GROUP BY ECH_CODIGO, CPES_FECHA"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
       If DMY(Rec1!Maximo) = DMY(TxtCesFecha.Text) Then
            lblEstado.Caption = ""
            MsgBox "NO se puede registrar el mismo carácter en la misma fecha.", 16, TIT_MSGBOX
            Rec1.Close
            Exit Sub
       End If
    End If
    Rec1.Close
            
    If Trim(Me.TxtCheNumero.Text) = "" Or _
       Trim(Me.cboBanco.ListIndex = -1) = "" Or _
       Trim(Me.TxtCesFecha.Text) = "" Then
       
        If Trim(Me.TxtCheNumero.Text) = "" Then
           MsgBox "Falta el Número de Cheque.", 16, TIT_MSGBOX
           TxtCheNumero.SetFocus
           lblEstado.Caption = ""
           Exit Sub
        End If
        If cboBanco.ListIndex = -1 Then
           MsgBox "Falta el BANCO", 16, TIT_MSGBOX
           cboBanco.SetFocus
           lblEstado.Caption = ""
           Exit Sub
        End If
        If Trim(Me.TxtCesFecha.Text) = "" Then
           MsgBox "Falta la Fecha.", 16, TIT_MSGBOX
           TxtCesFecha.SetFocus
           lblEstado.Caption = ""
           Exit Sub
        End If
 Else
        'Inserto en Cheque_Estados
         sql = "INSERT INTO CHEQUE_PROPIO_ESTADO (ECH_CODIGO, BAN_CODINT,"
         sql = sql & " CHEP_NUMERO, CPES_FECHA, CPES_DESCRI)"
         sql = sql & " VALUES ("
         sql = sql & XN(CboEstado.ItemData(CboEstado.ListIndex)) & ","
         sql = sql & XN(cboBanco.ItemData(cboBanco.ListIndex)) & ","
         sql = sql & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XDQ(TxtCesFecha) & ","
         sql = sql & XS(Me.TxtCheObserv.Text) & ")"
         DBConn.Execute sql
         
         CmdNuevo_Click
   End If
 End If
End Sub

Private Sub CmdNuevo_Click()
    Dim MtrObjetos As Variant

   lblEstado.Caption = ""
   Me.TxtCheNumero.Enabled = True
   Me.cboBanco.Enabled = True
   Me.TxtCheNumero.Text = ""
   Me.TxtCheFecEmi.Text = ""
   Me.TxtCheFecVto.Text = ""
   Me.TxtCheImport.Text = ""
   Me.Grd1.Rows = 1
   Me.TxtCesFecha.Text = ""
   Me.CboEstado.ListIndex = 0
   Me.cboBanco.ListIndex = 0
   Me.TxtCheObserv.Text = ""
   MtrObjetos = Array(TxtCheNumero, cboBanco)
   Call CambiarColor(MtrObjetos, 2, &H80000005, "E")
   Me.TxtCheNumero.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMCambioEstadoChPropio = Nothing
End Sub

Private Sub Form_Activate()
    Call Centrar_pantalla(Me)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'si presiono ESCAPE salgo del form
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
    If KeyAscii = vbKeyReturn Then 'avanza de campo
        MySendKeys Chr(9)
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
    Grd1.FormatString = "^Fecha|Estado|Observación"
    Grd1.ColWidth(0) = 1100
    Grd1.ColWidth(1) = 2500
    Grd1.ColWidth(2) = 4500
    Grd1.Rows = 1
    
    Set Rec1 = New ADODB.Recordset
    Set rec = New ADODB.Recordset
    'Cargo el Combo de Estados
    CargoEstados
    'CARGO COMBO CON BANCOS DONDE HAY CUENTAS
    CargoBanco
End Sub

Private Sub CargoEstados()
    sql = "SELECT ECH_CODIGO,ECH_DESCRI FROM ESTADO_CHEQUE"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            CboEstado.AddItem Trim(rec!ECH_DESCRI) '& Space(100 - Len(Trim(rec!ECH_DESCRI))) & Trim(rec!ech_codigo)
            CboEstado.ItemData(CboEstado.NewIndex) = rec!ECH_CODIGO
            rec.MoveNext
        Loop
        CboEstado.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub CargoBanco()
    sql = "SELECT DISTINCT B.BAN_CODINT, B.BAN_DESCRI"
    sql = sql & " FROM BANCO B, CTA_BANCARIA CB"
    sql = sql & " WHERE B.BAN_CODINT=CB.BAN_CODINT"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboBanco.AddItem Trim(rec!BAN_DESCRI)
            cboBanco.ItemData(cboBanco.NewIndex) = Trim(rec!BAN_CODINT)
            rec.MoveNext
        Loop
        cboBanco.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub TxtCheImport_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(TxtCheImport, KeyAscii)
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCheNumero_LostFocus()
    If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
       If Trim(Me.TxtCheNumero.Text) = "" Then
          Me.TxtCheNumero.SetFocus
       Else
           If Len(TxtCheNumero.Text) < 10 Then TxtCheNumero.Text = CompletarConCeros(TxtCheNumero.Text, 10)
       End If
    End If
End Sub

Private Sub TxtCheObserv_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub
