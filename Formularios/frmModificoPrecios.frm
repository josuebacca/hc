VERSION 5.00
Begin VB.Form frmModificoPrecios 
   Caption         =   "Modificar Precios"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "A..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   90
      TabIndex        =   13
      Top             =   960
      Width           =   5160
      Begin VB.TextBox txtTodos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   9
         Top             =   975
         Width           =   840
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   8
         Top             =   960
         Width           =   1170
      End
      Begin VB.OptionButton OptRubro 
         Caption         =   "Rubro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   5
         Top             =   660
         Width           =   885
      End
      Begin VB.OptionButton OptLinea 
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   345
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtLinea 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4185
         MaxLength       =   10
         TabIndex        =   4
         Top             =   285
         Width           =   855
      End
      Begin VB.ComboBox cboRubro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1725
         TabIndex        =   6
         Top             =   630
         Width           =   2160
      End
      Begin VB.ComboBox cboLinea 
         Height          =   315
         Left            =   1725
         TabIndex        =   3
         Top             =   285
         Width           =   2160
      End
      Begin VB.TextBox txtRubro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   4185
         MaxLength       =   10
         TabIndex        =   7
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lblTodos 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1455
         TabIndex        =   16
         Top             =   1020
         Width           =   195
      End
      Begin VB.Label lblRub 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3975
         TabIndex        =   15
         Top             =   675
         Width           =   195
      End
      Begin VB.Label lblLinea 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3975
         TabIndex        =   14
         Top             =   345
         Width           =   195
      End
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   525
      Left            =   3480
      TabIndex        =   18
      Top             =   2445
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Caption         =   "Por..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   12
      Top             =   30
      Width           =   5160
      Begin VB.OptionButton OptPorc 
         Caption         =   "Porcentaje (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   0
         Top             =   315
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptPesos 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   1
         Top             =   585
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   525
      Left            =   4335
      TabIndex        =   11
      Top             =   2445
      Width           =   900
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   525
      Left            =   2565
      TabIndex        =   10
      Top             =   2445
      Width           =   900
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
      Left            =   135
      TabIndex        =   17
      Top             =   2550
      Width           =   660
   End
End
Attribute VB_Name = "frmModificoPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim codlista As String

Private Sub cboLinea_LostFocus()
    If OptRubro.Value = True Then
        If cboLinea.ListCount > 0 Then
            cargocboRubro cboLinea.ItemData(cboLinea.ListIndex)
        Else
            cboRubro.Clear
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim Porc As Double
    Dim TOTAL As String
    Dim i As Integer
    
  On Error GoTo SeReclavose
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Actualizando..."
    DBConn.BeginTrans
    
    sql = " SELECT P.PTO_DESCRI, D.LIS_PRECIO, P.PTO_CODIGO"
    sql = sql & " FROM PRODUCTO P, LISTA_PRECIO LP, DETALLE_LISTA_PRECIO D"
    sql = sql & " WHERE"
    sql = sql & " D.PTO_CODIGO = P.PTO_CODIGO"
    sql = sql & " AND LP.LIS_CODIGO = D.LIS_CODIGO"
    sql = sql & " AND LP.LIS_CODIGO = " & XN(Trim(codlista))
    If OptLinea.Value = True And cboLinea.ListCount > 0 Then
        sql = sql & " AND P.LNA_CODIGO = " & XN(cboLinea.ItemData(cboLinea.ListIndex))
    End If
    If OptRubro.Value = True And cboRubro.ListCount > 0 Then
        sql = sql & " AND P.LNA_CODIGO = " & XN(cboLinea.ItemData(cboLinea.ListIndex))
        sql = sql & " AND P.RUB_CODIGO = " & XN(cboRubro.ItemData(cboRubro.ListIndex))
    End If
    sql = sql & " ORDER BY P.PTO_DESCRI "
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    
    If rec.EOF = False Then
        Do While rec.EOF = False
            'VERIFICO SI ES PORCENTAJE O IMPORTE Y LO CALCULO
                If OptPorc.Value = True Then
                    If OptLinea.Value = True Then
                        Porc = (CDbl(rec!LIS_PRECIO) * CDbl(txtLinea.Text)) / 100
                        TOTAL = CDbl(rec!LIS_PRECIO) + Porc
                    ElseIf OptRubro.Value = True Then
                        Porc = (CDbl(rec!LIS_PRECIO) * CDbl(txtRubro.Text)) / 100
                        TOTAL = CDbl(rec!LIS_PRECIO) + Porc
                    ElseIf OptTodos.Value = True Then
                        Porc = (CDbl(rec!LIS_PRECIO) * CDbl(txtTodos.Text)) / 100
                        TOTAL = CDbl(rec!LIS_PRECIO) + Porc
                    End If
                End If
                If OptPesos.Value = True Then
                   If OptLinea.Value = True Then
                       TOTAL = CDbl(rec!LIS_PRECIO) + CDbl((txtLinea.Text))
                   ElseIf OptRubro.Value = True Then
                       TOTAL = CDbl(rec!LIS_PRECIO) + CDbl(txtRubro.Text)
                   ElseIf OptTodos.Value = True Then
                       TOTAL = CDbl(rec!LIS_PRECIO) + CDbl(txtTodos.Text)
                   End If
                End If
                
            If Trim(codlista) <> "0" Then
                 'GUARDO LOS CAMBIOS EN LA GRILLA Y EN LA TABLA
'                 i = 1
'                 For i = 1 To FrmListadePrecios.GrdModulos.Rows - 1
'                    If FrmListadePrecios.GrdModulos.TextMatrix(i, 0) = rec!PTO_CODIGO Then
'                        If optPrecio.Value = True Then
'                            FrmListadePrecios.GrdModulos.TextMatrix(i, 3) = Valido_Importe(TOTAL)
'                        Else
'                            FrmListadePrecios.GrdModulos.TextMatrix(i, 4) = Valido_Importe(TOTAL)
'                        End If
'                        Exit For
'                    End If
'                 Next
                 'ACTUALIZO LA TABLA
                 sql = "UPDATE DETALLE_LISTA_PRECIO "
                 sql = sql & " SET LIS_PRECIO = " & XN(TOTAL)
                 sql = sql & " WHERE LIS_CODIGO = " & XN(Trim(codlista))
                 sql = sql & " AND PTO_CODIGO = " & XN(rec!PTO_CODIGO)
                 DBConn.Execute sql
            Else
                'GUARDO LOS CAMBIOS SOLO EN LA GRILLA
                i = 1
                 For i = 1 To FrmListadePrecios.GrdModulos.Rows - 1
                    If FrmListadePrecios.GrdModulos.TextMatrix(i, 0) = rec!PTO_CODIGO Then
                            FrmListadePrecios.GrdModulos.TextMatrix(i, 5) = Valido_Importe(TOTAL)
                        Exit For
                    End If
                 Next
            End If
            rec.MoveNext
        Loop
    Else
        MsgBox "No hay coincidencias en la busqueda.", vbOKOnly + vbCritical, TIT_MSGBOX
    End If
    DBConn.CommitTrans
    rec.Close
    Screen.MousePointer = vbNormal
    CmdNuevo_Click
    CmdSalir_Click
    Exit Sub
    
SeReclavose:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    lblEstado.Caption = ""
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub CmdNuevo_Click()
    cboLinea.ListIndex = 0
    cboRubro.Clear
    txtLinea.Text = "0,00"
    txtRubro.Text = "0,00"
    txtTodos.Text = "0,00"
    lblEstado.Caption = ""
    OptPorc.Value = True
    OptLinea.Value = True
    If Me.Visible = True Then OptPorc.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set frmModificoPrecios = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    Call Centrar_pantalla(Me)
    cargocboLinea
    If FrmListadePrecios.txtcodigo.Text = "" Then
        codlista = "0"
    Else
        codlista = Trim(FrmListadePrecios.txtcodigo.Text)
    End If
    CmdNuevo_Click
End Sub

Function cargocboLinea()
    cboLinea.Clear
    sql = "SELECT * FROM LINEAS  ORDER BY LNA_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboLinea.AddItem rec!LNA_DESCRI
            cboLinea.ItemData(cboLinea.NewIndex) = rec!LNA_CODIGO
            rec.MoveNext
        Loop
        If cboLinea.ListCount > 0 Then cboLinea.ListIndex = 0
    End If
    rec.Close
End Function

Function cargocboRubro(mLinea As String)
    cboRubro.Clear
    sql = "SELECT * FROM RUBROS "
    sql = sql & " WHERE LNA_CODIGO=" & XN(mLinea)
    sql = sql & " ORDER BY RUB_DESCRI"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            cboRubro.AddItem rec!RUB_DESCRI
            cboRubro.ItemData(cboRubro.NewIndex) = rec!RUB_CODIGO
            rec.MoveNext
        Loop
        If cboRubro.ListCount > 0 Then cboRubro.ListIndex = 0
    End If
    rec.Close
End Function

Function SeteosDos(Num As Integer)
    If Num = 1 Then
        cboLinea.Enabled = True
        txtLinea.Enabled = True
        cboRubro.Enabled = False
        txtRubro.Enabled = False
        txtTodos.Enabled = False
        txtLinea.SetFocus
        txtRubro.Text = "0,00"
        txtTodos.Text = "0,00"
        
        If cboLinea.ListCount > 0 Then cboLinea.ListIndex = 0
        cboRubro.Clear
    End If
    If Num = 2 Then
        cboLinea.Enabled = True
        txtLinea.Enabled = True
        cboRubro.Enabled = True
        txtRubro.Enabled = True
        txtTodos.Enabled = False
        txtRubro.SetFocus
        txtLinea.Text = "0,00"
        txtTodos.Text = "0,00"
        
        cboRubro.Clear
        If cboLinea.ListCount > 0 Then cboLinea.ListIndex = 0
        cboLinea_LostFocus
    End If
    If Num = 3 Then
        cboLinea.Enabled = False
        txtLinea.Enabled = False
        cboRubro.Enabled = False
        txtRubro.Enabled = False
        txtTodos.Enabled = True
        txtTodos.SetFocus
        txtRubro.Text = "0,00"
        txtLinea.Text = "0,00"
        
        If cboLinea.ListCount > 0 Then cboLinea.ListIndex = 0
        cboRubro.Clear
  End If
End Function

Private Sub OptLinea_Click()
    SeteosDos (1)
End Sub

Private Sub OptPesos_Click()
    lblLinea.Caption = "$"
    lblRub.Caption = "$"
    lblTodos.Caption = "$"
End Sub

Private Sub OptPorc_Click()
    lblLinea.Caption = "%"
    lblRub.Caption = "%"
    lblTodos.Caption = "%"
End Sub

Private Sub OptRep_Click()
    SeteosDos (3)
End Sub

Private Sub OptRubro_Click()
    SeteosDos (2)
End Sub

Private Sub OptTodos_Click()
    SeteosDos (3)
End Sub

Private Sub txtLinea_GotFocus()
    SelecTexto txtLinea
End Sub

Private Sub txtLinea_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtLinea, KeyAscii, True)
End Sub

Private Sub txtLinea_LostFocus()
    If txtLinea.Text <> "" Then
        If OptPesos.Value = True Then
            txtLinea.Text = Valido_Importe(txtLinea)
        Else
            If ValidarPorcentaje(txtLinea) = False Then
                txtLinea.SetFocus
            End If
        End If
    Else
        txtLinea.Text = "0,00"
    End If
End Sub

Private Sub txtRubro_GotFocus()
    SelecTexto txtRubro
End Sub

Private Sub txtRubro_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtRubro, KeyAscii, True)
End Sub

Private Sub txtRubro_LostFocus()
    If txtRubro.Text <> "" Then
        If OptPesos.Value = True Then
            txtRubro.Text = Valido_Importe(txtRubro)
        Else
            If ValidarPorcentaje(txtRubro) = False Then
                txtRubro.SetFocus
            End If
        End If
    Else
        txtRubro.Text = "0,00"
    End If
End Sub

Private Sub txtTodos_GotFocus()
    SelecTexto txtTodos
End Sub

Private Sub txtTodos_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroDecimal(txtTodos, KeyAscii, True)
End Sub

Private Sub txtTodos_LostFocus()
    If txtTodos.Text <> "" Then
        If OptPesos.Value = True Then
            txtTodos.Text = Valido_Importe(txtTodos)
        Else
            If ValidarPorcentaje(txtTodos) = False Then
                txtTodos.SetFocus
            End If
        End If
    Else
        txtTodos.Text = "0,00"
    End If
End Sub
