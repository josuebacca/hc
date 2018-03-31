VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "FECHA32.OCX"
Begin VB.Form ABMCambioEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ABM de Estado de Cheques"
   ClientHeight    =   6030
   ClientLeft      =   2280
   ClientTop       =   435
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ABMCambioEstado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7170
   Begin FechaCtl.Fecha TxtCheFecVto 
      Height          =   300
      Left            =   5055
      TabIndex        =   6
      Top             =   1755
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin FechaCtl.Fecha TxtCheFecEmi 
      Height          =   375
      Left            =   1380
      TabIndex        =   5
      Top             =   1755
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
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
      Height          =   60
      Left            =   105
      TabIndex        =   31
      Top             =   5490
      Width           =   6975
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
      TabIndex        =   7
      Top             =   2100
      Width           =   1140
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
      Top             =   180
      Width           =   1230
   End
   Begin VB.Frame Frame3 
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
      Height          =   1125
      Left            =   165
      TabIndex        =   21
      Top             =   540
      Width           =   6135
      Begin VB.TextBox TxtBanDescri 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   165
         TabIndex        =   22
         Top             =   645
         Width           =   5820
      End
      Begin VB.TextBox TxtCODIGO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5145
         MaxLength       =   6
         TabIndex        =   4
         Top             =   285
         Width           =   795
      End
      Begin VB.TextBox TxtLOCALIDAD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2250
         MaxLength       =   3
         TabIndex        =   2
         Top             =   285
         Width           =   540
      End
      Begin VB.TextBox TxtBANCO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   1
         Top             =   285
         Width           =   540
      End
      Begin VB.TextBox TxtSUCURSAL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   3
         Top             =   285
         Width           =   540
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
         Left            =   4515
         TabIndex        =   26
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
         Left            =   3000
         TabIndex        =   25
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
         Left            =   210
         TabIndex        =   24
         Top             =   330
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
         Left            =   1470
         TabIndex        =   23
         Top             =   315
         Width           =   720
      End
   End
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
      Height          =   345
      Left            =   5730
      TabIndex        =   20
      Top             =   150
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5625
      Width           =   915
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   5190
      TabIndex        =   12
      Top             =   5625
      Width           =   915
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   4260
      TabIndex        =   11
      Top             =   5625
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
      Height          =   60
      Left            =   105
      TabIndex        =   14
      Top             =   2460
      Width           =   6975
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
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   6900
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
      Left            =   3915
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4170
      Width           =   3165
   End
   Begin MSFlexGridLib.MSFlexGrid Grd1 
      Height          =   1500
      Left            =   90
      TabIndex        =   15
      Top             =   2565
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
   Begin FechaCtl.Fecha TxtCesFecha 
      Height          =   300
      Left            =   2025
      TabIndex        =   8
      Top             =   4170
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      Separador       =   "/"
      Text            =   ""
      MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Emisión:"
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   30
      Top             =   1800
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Vencimiento:"
      Height          =   195
      Index           =   5
      Left            =   3570
      TabIndex        =   29
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   28
      Top             =   2145
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro Cheque:"
      Height          =   195
      Index           =   7
      Left            =   210
      TabIndex        =   27
      Top             =   225
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   19
      Top             =   4575
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Cambio de Estado:"
      Height          =   270
      Index           =   1
      Left            =   150
      TabIndex        =   18
      Top             =   4200
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado:"
      Height          =   195
      Index           =   0
      Left            =   3180
      TabIndex        =   17
      Top             =   4200
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
      Left            =   150
      TabIndex        =   16
      Top             =   5655
      Width           =   660
   End
End
Attribute VB_Name = "ABMCambioEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdGrabar_Click()
 If Me.ActiveControl.Name <> "CmdNuevo" And Me.ActiveControl.Name <> "CmdSalir" Then

    'Verifico que NO graben dos veces el mismo estado en el mismo día
    sql = "SELECT ECH_CODIGO,MAX(CES_FECHA)as maximo"
    sql = sql & " FROM CHEQUE_ESTADOS "
    sql = sql & " WHERE CHE_NUMERO = " & XS(Me.TxtCheNumero.Text)
    sql = sql & " AND ECH_CODIGO = " & CboEstado.ItemData(CboEstado.ListIndex)
    sql = sql & " GROUP BY ECH_CODIGO,CES_FECHA"
    
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.RecordCount > 0 Then
       If DMY(Rec1!Maximo) = DMY(TxtCesFecha.Text) Then
            MsgBox "NO se puede registrar el mismo carácter en la misma fecha.", 16, TIT_MSGBOX
            Rec1.Close
            Exit Sub
       End If
    End If
    Rec1.Close
            
    If Trim(Me.TxtCheNumero.Text) = "" Or _
       Trim(Me.TxtBANCO.Text) = "" Or _
       Trim(Me.TxtLOCALIDAD.Text) = "" Or _
       Trim(Me.TxtSUCURSAL.Text) = "" Or _
       Trim(Me.TxtCODIGO.Text) = "" Or _
       Trim(Me.TxtCesFecha.Text) = "" Then
       
        If Trim(Me.TxtCheNumero.Text) = "" Then
           MsgBox "Falta el Número de Cheque.", 16, TIT_MSGBOX
           TxtCheNumero.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtBANCO.Text) = "" Then
           MsgBox "Falta el BANCO.", 16, TIT_MSGBOX
           TxtBANCO.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtLOCALIDAD.Text) = "" Then
           MsgBox "Falta la LOCALIDAD.", 16, TIT_MSGBOX
           TxtLOCALIDAD.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtSUCURSAL.Text) = "" Then
           MsgBox "Falta la SUCURSAL.", 16, TIT_MSGBOX
           TxtSUCURSAL.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtCODIGO.Text) = "" Then
           MsgBox "Falta el CÓDIGO.", 16, TIT_MSGBOX
           TxtCODIGO.SetFocus
           Exit Sub
        End If
        
        If Trim(Me.TxtCesFecha.Text) = "" Then
           MsgBox "Falta la Fecha.", 16, TIT_MSGBOX
           TxtCesFecha.SetFocus
           Exit Sub
        End If
 Else
        
        'Inserto en Cheque_Estados
         sql = "INSERT INTO CHEQUE_ESTADOS(ECH_CODIGO,BAN_CODINT,CHE_NUMERO,CES_FECHA,"
         sql = sql & " CES_DESCRI)VALUES ( " & CboEstado.ItemData(CboEstado.ListIndex)
         sql = sql & "," & XN(Me.TxtCodInt.Text) & "," & XS(Me.TxtCheNumero.Text)
         sql = sql & "," & XDQ(TxtCesFecha) & "," & XS(Me.TxtCheObserv.Text) & " )"
         DBConn.Execute sql
         
         CmdNuevo_Click
   End If
 End If
End Sub

Private Sub CmdNuevo_Click()
   Dim MtrObjetos As Variant
   
   Me.TxtCheNumero.Enabled = True
   Me.TxtBANCO.Enabled = True
   Me.TxtLOCALIDAD.Enabled = True
   Me.TxtSUCURSAL.Enabled = True
   Me.TxtCODIGO.Enabled = True
   MtrObjetos = Array(TxtCheNumero, TxtBANCO, TxtLOCALIDAD, TxtSUCURSAL, TxtCODIGO)
   Call CambiarColor(MtrObjetos, 5, &H80000005, "E")
            
   Me.TxtCheNumero.Text = ""
   Me.TxtBANCO.Text = ""
   Me.TxtLOCALIDAD.Text = ""
   Me.TxtSUCURSAL.Text = ""
   Me.TxtCODIGO.Text = ""
   
   Me.TxtCodInt.Text = ""
   Me.TxtCheFecEmi.Text = ""
   Me.TxtCheFecVto.Text = ""
   Me.TxtCheImport.Text = ""
   Me.Grd1.Rows = 1
   Me.TxtCesFecha.Text = ""
   Me.CboEstado.ListIndex = 0
   Me.TxtCheObserv.Text = ""
   Me.TxtCheNumero.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set ABMCambioEstado = Nothing
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
    
    'Cargo el Combo de Estados
    Set rec = New ADODB.Recordset
    
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

Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtBANCO_LostFocus()
    If TxtBANCO.Text <> "" Then TxtBANCO.Text = Format(TxtBANCO.Text, "000")
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

Private Sub TxtCodigo_Change()
    If Trim(TxtCODIGO) = "" And CmdNuevo.Enabled Then
        CmdNuevo.Enabled = False
    ElseIf Trim(TxtCODIGO) <> "" Then
        CmdNuevo.Enabled = True
    End If
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtCodigo_LostFocus()
    
    If Trim(Me.TxtCheNumero.Text) <> "" And _
       Trim(Me.TxtBANCO.Text) <> "" And _
       Trim(Me.TxtLOCALIDAD.Text) <> "" And _
       Trim(Me.TxtSUCURSAL.Text) <> "" And _
       Trim(Me.TxtCODIGO.Text) <> "" Then
       
       If Len(Me.TxtCODIGO.Text) < 6 Then Me.TxtCODIGO.Text = CompletarConCeros(Me.TxtCODIGO.Text, 6)
           
       Dim MtrObjetos As Variant
    
       Set rec = New ADODB.Recordset
       Set Rec1 = New ADODB.Recordset
       
       'BUSCO EL CODIGO INTERNO
       sql = "SELECT BAN_CODINT,BAN_DESCRI FROM BANCO WHERE BAN_BANCO = " & _
       XS(TxtBANCO.Text) & " AND BAN_LOCALIDAD = " & _
       XS(Me.TxtLOCALIDAD.Text) & " AND BAN_SUCURSAL = " & _
       XS(Me.TxtSUCURSAL.Text) & " AND BAN_CODIGO = " & XS(Me.TxtCODIGO.Text)
       rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
       If rec.RecordCount > 0 Then 'EXITE
            Me.TxtCodInt.Text = rec!BAN_CODINT
            TxtBanDescri.Text = rec!BAN_DESCRI
       Else
          MsgBox "NO ESTA REGISTRADO EL BANCO.", 16, TIT_MSGBOX
          Me.TxtCODIGO.Text = ""
          Me.TxtCODIGO.SetFocus
          Exit Sub
       End If
       rec.Close
       
       'CONSULTO SI EXISTE EL CHEQUE
        sql = "SELECT * FROM CHEQUE " & _
              " WHERE CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                " AND BAN_CODINT = " & XN(TxtCodInt.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then 'EXITE
            
            TxtCheNumero.Enabled = False
            TxtBANCO.Enabled = False
            TxtLOCALIDAD.Enabled = False
            TxtSUCURSAL.Enabled = False
            TxtCODIGO.Enabled = False
            
            MtrObjetos = Array(TxtCheNumero, TxtBANCO, TxtLOCALIDAD, TxtSUCURSAL, TxtCODIGO)
            Call CambiarColor(MtrObjetos, 5, &H80000018, "D")
            
            Me.TxtCheFecEmi.Text = rec!CHE_FECEMI
            Me.TxtCheFecVto.Text = rec!CHE_FECVTO
            Me.TxtCheImport.Text = Format(rec!che_import, "$ #0.00")

            'Cargo la Grilla
            sql = "SELECT CES_FECHA,ECH_DESCRI,CES_DESCRI" & _
                  " FROM CHEQUE_ESTADOS CE, ESTADO_CHEQUE EC " & _
                  " WHERE CE.ECH_CODIGO = EC.ECH_CODIGO " & _
                    " AND CE.CHE_NUMERO = " & XS(TxtCheNumero.Text) & _
                    " AND CE.BAN_CODINT = " & XN(TxtCodInt.Text) & _
                    " ORDER BY CES_FECHA"
            Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If Rec1.RecordCount > 0 Then
               Rec1.MoveFirst
               Do While Not Rec1.EOF
                 Grd1.AddItem Rec1!CES_FECHA & Chr(9) & Trim(Rec1.Fields(1)) & Chr(9) & Trim(Rec1.Fields(2))
                 Rec1.MoveNext
               Loop
            End If
            Rec1.Close
            Me.TxtCesFecha.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)
KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub TxtLOCALIDAD_LostFocus()
    If TxtLOCALIDAD.Text <> "" Then TxtLOCALIDAD.Text = Format(TxtLOCALIDAD.Text, "000")
End Sub

Private Sub TxtSucursal_KeyPress(KeyAscii As Integer)
KeyAscii = CarNumeroTE(KeyAscii)
End Sub

Private Sub txtSucursal_LostFocus()
    If TxtSUCURSAL.Text <> "" Then TxtSUCURSAL.Text = Format(TxtSUCURSAL.Text, "000")
End Sub
