VERSION 5.00
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form FrmCargaCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Cheques de Terceros"
   ClientHeight    =   4065
   ClientLeft      =   2535
   ClientTop       =   1005
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCargaCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FechaCtl.Fecha TxtCheFecEmi 
      Height          =   360
      Left            =   1380
      TabIndex        =   8
      Top             =   2415
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
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
      Height          =   330
      Left            =   1380
      TabIndex        =   10
      Top             =   2745
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
      Height          =   330
      Left            =   1380
      MaxLength       =   60
      TabIndex        =   11
      Top             =   3120
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
      Height          =   330
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1665
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
      Height          =   330
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   7
      Top             =   2040
      Width           =   5040
   End
   Begin VB.Frame Frame2 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   495
      Width           =   6300
      Begin VB.TextBox TxtCodInt 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5370
         TabIndex        =   20
         Top             =   660
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TxtCODIGO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4755
         MaxLength       =   6
         TabIndex        =   5
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox TxtLOCALIDAD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2175
         MaxLength       =   3
         TabIndex        =   3
         Top             =   285
         Width           =   450
      End
      Begin VB.TextBox TxtBANCO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   780
         MaxLength       =   3
         TabIndex        =   2
         Top             =   285
         Width           =   450
      End
      Begin VB.TextBox TxtSUCURSAL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3525
         MaxLength       =   3
         TabIndex        =   4
         Top             =   285
         Width           =   450
      End
      Begin VB.CommandButton CmdBanco 
         DisabledPicture =   "FrmCargaCheques.frx":08CA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5655
         Picture         =   "FrmCargaCheques.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Agregar Banco"
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox TxtBanDescri 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   21
         Top             =   675
         Width           =   5820
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4125
         TabIndex        =   25
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2820
         TabIndex        =   24
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   225
         TabIndex        =   23
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   1395
         TabIndex        =   22
         Top             =   315
         Width           =   720
      End
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
      Left            =   5010
      MaxLength       =   10
      TabIndex        =   1
      Top             =   150
      Width           =   1380
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5490
      TabIndex        =   15
      Top             =   3645
      Width           =   900
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   4575
      TabIndex        =   14
      Top             =   3645
      Width           =   900
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   3660
      TabIndex        =   13
      Top             =   3645
      Width           =   900
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   2745
      TabIndex        =   12
      Top             =   3645
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
      Left            =   120
      TabIndex        =   16
      Top             =   3540
      Width           =   6345
   End
   Begin FechaCtl.Fecha TxtCheFecVto 
      Height          =   360
      Left            =   5295
      TabIndex        =   9
      Top             =   2415
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin FechaCtl.Fecha TxtCheFecEnt 
      Height          =   360
      Left            =   1335
      TabIndex        =   0
      Top             =   150
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   635
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Vto:"
      Height          =   195
      Index           =   3
      Left            =   4185
      TabIndex        =   33
      Top             =   2445
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Emisión:"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   32
      Top             =   2444
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   31
      Top             =   2775
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Index           =   0
      Left            =   795
      TabIndex        =   30
      Top             =   195
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   29
      Top             =   3150
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cheque:"
      Height          =   195
      Index           =   7
      Left            =   3960
      TabIndex        =   28
      Top             =   210
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Responsable:"
      Height          =   195
      Index           =   9
      Left            =   150
      TabIndex        =   27
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Concepto:"
      Height          =   195
      Index           =   10
      Left            =   150
      TabIndex        =   26
      Top             =   2077
      Width           =   750
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
      TabIndex        =   17
      Top             =   3690
      Width           =   660
   End
End
Attribute VB_Name = "FrmCargaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function Validar() As Boolean
   If Trim(TxtCheNumero.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Número de Cheque.", 16, TIT_MSGBOX
        TxtCheNumero.SetFocus
        Exit Function
        
   ElseIf Trim(TxtBanco.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Banco.", 16, TIT_MSGBOX
        TxtBanco.SetFocus
        Exit Function
        
   ElseIf Trim(TxtLOCALIDAD.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Localidad del Banco.", 16, TIT_MSGBOX
        TxtLOCALIDAD.SetFocus
        Exit Function
        
   ElseIf Trim(TxtSucursal.Text) = "" Then
        Validar = False
        MsgBox "Ingrese la Sucursal del Banco.", 16, TIT_MSGBOX
        TxtSucursal.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCODIGO.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Código del Banco.", 16, TIT_MSGBOX
        TxtCODIGO.SetFocus
        Exit Function
        
   ElseIf Trim(Me.TxtCodInt.Text) = "" Then
        Validar = False
        MsgBox "Verifique el Código de Banco.", 16, TIT_MSGBOX
        TxtCODIGO.SetFocus
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
    
   ElseIf Me.TxtCheNombre.Text = "" Then
        Validar = False
        MsgBox "Debe ingresar la Persona responsable.!", 16, TIT_MSGBOX
        TxtCheNombre.SetFocus
        Exit Function
        
   ElseIf Trim(TxtCheImport.Text) = "" Then
        Validar = False
        MsgBox "Ingrese el Importe del Cheque.", 16, TIT_MSGBOX
        TxtCheImport.SetFocus
        Exit Function
   End If
   
   Validar = True
End Function

Private Sub CmdBanco_Click()
    Viene_Cheque = True
    ABMBancos.Show vbModal
    Viene_Cheque = False
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo CLAVOSE
    
    If Trim(TxtCheNumero.Text) <> "" And Trim(Me.TxtCodInt.Text) <> "" Then
        If MsgBox("Seguro desea eliminar el Cheque Nº: " & Trim(Me.TxtCheNumero.Text) & "? ", 36, TIT_MSGBOX) = vbYes Then
        
            Screen.MousePointer = vbHourglass
            lblEstado.Caption = "Borrando..."
            
            sql = "SELECT BOL_NUMERO "
            sql = sql & " FROM ChequeEstadoVigente "
            sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
            sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                If Not IsNull(rec!BOL_NUMERO) Then
                   MsgBox "No se puede eliminar este Cheque porque fue depositado", vbExclamation, TIT_MSGBOX
                   rec.Close
                   Screen.MousePointer = vbNormal
                   lblEstado.Caption = ""
                   Exit Sub
                 End If
            End If
            rec.Close
    
            DBConn.BeginTrans
            DBConn.Execute "DELETE FROM CHEQUE_ESTADOS WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
                           
            DBConn.Execute "DELETE FROM CHEQUE WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            
            Screen.MousePointer = vbNormal
            lblEstado.Caption = ""
            DBConn.CommitTrans
            CmdNuevo_Click
        End If
    End If
    Exit Sub
    
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
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
    
    sql = "SELECT * FROM CHEQUE WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & " AND BAN_CODINT = " & XN(TxtCodInt.Text)
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount = 0 Then
         sql = "INSERT INTO CHEQUE(CHE_NUMERO,BAN_CODINT,CHE_NOMBRE,CHE_IMPORT,CHE_FECEMI,"
         sql = sql & "CHE_FECVTO,CHE_FECENT,CHE_MOTIVO,CHE_OBSERV)"
         sql = sql & " VALUES (" & XS(Me.TxtCheNumero.Text) & ","
         sql = sql & XN(Me.TxtCodInt.Text) & "," & XS(Me.TxtCheNombre.Text) & ","
         sql = sql & XN(Me.TxtCheImport.Text) & "," & XDQ(Me.TxtCheFecEmi.Text) & ","
         sql = sql & XDQ(Me.TxtCheFecVto.Text) & "," & XDQ(Me.TxtCheFecEnt.Text) & ","
         sql = sql & XS(Me.TxtCheMotivo.Text) & "," & XS(Me.TxtCheObserv.Text) & " )"
         DBConn.Execute sql
    Else
         sql = "UPDATE CHEQUE SET CHE_NOMBRE = " & XS(Me.TxtCheNombre.Text)
         sql = sql & ",CHE_IMPORT = " & XN(Me.TxtCheImport.Text)
         sql = sql & ",CHE_FECEMI =" & XDQ(Me.TxtCheFecEmi.Text)
         sql = sql & ",CHE_FECVTO =" & XDQ(Me.TxtCheFecVto.Text)
         sql = sql & ",CHE_FECENT = " & XDQ(Me.TxtCheFecEnt.Text)
         sql = sql & ",CHE_MOTIVO = " & XS(Me.TxtCheMotivo.Text)
         sql = sql & ",CHE_OBSERV = " & XS(Me.TxtCheObserv.Text)
         sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
         sql = sql & " AND BAN_CODINT = " & XN(Me.TxtCodInt.Text)
         DBConn.Execute sql
    End If
    rec.Close
     
    'Insert en la Tabla de Estados de Cheques
    sql = "INSERT INTO CHEQUE_ESTADOS (CHE_NUMERO,BAN_CODINT,ECH_CODIGO,CES_FECHA,CES_DESCRI)"
    sql = sql & " VALUES ("
    sql = sql & XS(Me.TxtCheNumero.Text) & ","
    sql = sql & XN(Me.TxtCodInt.Text) & "," & XN(1) & ","
    sql = sql & XDQ(Date) & ",'CHEQUE EN CARTERA')"
    DBConn.Execute sql
    
    '************* PREGUNTAR POR SI DESEA IMPRIMIR ***************
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.CommitTrans
    CmdNuevo_Click
 End If
 Exit Sub
      
CLAVOSE:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
End Sub

Private Sub CmdNuevo_Click()
    Me.TxtCheFecEnt.Text = ""
    Me.TxtCheNumero.Enabled = True
    Me.TxtBanco.Enabled = True
    Me.TxtLOCALIDAD.Enabled = True
    Me.TxtSucursal.Enabled = True
    Me.TxtCODIGO.Enabled = True
    Me.TxtCheNombre.Enabled = True
    MtrObjetos = Array(TxtCheNumero, TxtBanco, TxtLOCALIDAD, TxtSucursal, TxtCODIGO)
    Call CambiarColor(MtrObjetos, 5, &H80000005, "E")
    TxtCheNombre.ForeColor = &H80000008
    Me.TxtCheNumero.Text = ""
    Me.TxtBanco.Text = ""
    Me.TxtLOCALIDAD.Text = ""
    Me.TxtSucursal.Text = ""
    Me.TxtCODIGO.Text = ""
    Me.TxtCodInt.Text = ""
    Me.TxtBanDescri.Text = ""
    Me.TxtCheNombre.Text = ""
    Me.TxtCheMotivo.Text = ""
    Me.TxtCheFecEmi.Text = ""
    Me.TxtCheFecVto.Text = ""
    Me.TxtCheImport.Text = ""
    Me.TxtCheObserv.Text = ""
    Me.TxtCheFecEnt.SetFocus
    lblEstado.Caption = ""
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmCargaCheques = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    
    TxtCheFecEnt.Text = Date
    lblEstado.Caption = ""
End Sub

Private Sub TxtBANCO_GotFocus()
    SelecTexto TxtBanco
End Sub

Private Sub TxtBANCO_LostFocus()
    If Len(TxtBanco.Text) < 3 Then TxtBanco.Text = CompletarConCeros(TxtBanco.Text, 3)
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

Private Sub TxtCheNumero_GotFocus()
    SelecTexto TxtCheNumero
End Sub

Private Sub TxtCheNumero_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
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

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtCodigo_GotFocus()
    SelecTexto TxtCODIGO
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    Dim MtrObjetos As Variant
    
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
        
    'ChequeRegistrado = False
    
    If Len(TxtCODIGO.Text) < 6 Then TxtCODIGO.Text = CompletarConCeros(TxtCODIGO.Text, 6)
     
    If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBanco.Text) <> "" And _
       Trim(Me.TxtLOCALIDAD.Text) <> "" And _
       Trim(Me.TxtSucursal.Text) <> "" And _
       Trim(Me.TxtCODIGO.Text) <> "" Then
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT, BAN_DESCRI FROM BANCO "
       sql = sql & " WHERE BAN_BANCO = " & XS(TxtBanco.Text)
       sql = sql & " AND BAN_LOCALIDAD = " & XS(Me.TxtLOCALIDAD.Text)
       sql = sql & " AND BAN_SUCURSAL = " & XS(Me.TxtSucursal.Text)
       sql = sql & " AND BAN_CODIGO = " & XS(TxtCODIGO.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
          TxtCodInt.Text = rec!BAN_CODINT
          TxtBanDescri.Text = rec!BAN_DESCRI
          rec.Close
       Else
          If Me.ActiveControl.Name <> "CmdSalir" And Me.ActiveControl.Name <> "CmdNuevo" Then
            MsgBox "Banco NO Registrado.", 16, TIT_MSGBOX
            Me.CmdBanco.SetFocus
          End If
          rec.Close
          Exit Sub
       End If
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE "
        sql = sql & " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text)
        sql = sql & " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCheFecEnt.Text = rec!CHE_FECENT
            Me.TxtCheNumero.Text = Trim(rec!CHE_NUMERO)
            
            'BUSCO LOS ATRIBUTOS DE BANCO
            sql = "SELECT BAN_BANCO,BAN_LOCALIDAD,BAN_SUCURSAL,BAN_CODIGO FROM BANCO " & _
                   "WHERE BAN_CODINT = " & XN(Me.TxtCodInt.Text)
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount > 0 Then 'EXITE
                Me.TxtBanco.Text = Rec1!BAN_BANCO
                Me.TxtLOCALIDAD.Text = Rec1!BAN_LOCALIDAD
                Me.TxtSucursal.Text = Rec1!BAN_SUCURSAL
                Me.TxtCODIGO.Text = Rec1!BAN_CODIGO
            End If
            Rec1.Close
            Me.TxtCheNombre.Text = ChkNull(rec!CHE_NOMBRE)
            Me.TxtCheMotivo.Text = rec!CHE_MOTIVO
            Me.TxtCheFecEmi.Text = rec!CHE_FECEMI
            Me.TxtCheFecVto.Text = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Valido_Importe(rec!che_import)
            Me.TxtCheObserv.Text = ChkNull(rec!CHE_OBSERV)
            
            TxtCheNumero.Enabled = False
            TxtBanco.Enabled = False
            TxtLOCALIDAD.Enabled = False
            TxtSucursal.Enabled = False
            TxtCODIGO.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBanco, TxtLOCALIDAD, TxtSucursal, TxtCODIGO)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
        Else
           TxtCheNombre.ForeColor = &H80000008
           rec.Close
           Exit Sub
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub TxtLOCALIDAD_GotFocus()
    SelecTexto TxtLOCALIDAD
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If Len(TxtLOCALIDAD.Text) < 3 Then TxtLOCALIDAD.Text = CompletarConCeros(TxtLOCALIDAD.Text, 3)
End Sub

Private Sub txtSucursal_GotFocus()
    SelecTexto TxtSucursal
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
     KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If Len(TxtSucursal.Text) < 3 Then TxtSucursal.Text = CompletarConCeros(TxtSucursal.Text, 3)
End Sub
