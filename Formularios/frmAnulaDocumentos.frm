VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{5F09B5DF-6F4D-11D2-8355-4854E82A9183}#15.0#0"; "Fecha32.ocx"
Begin VB.Form frmAnulaDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de ...."
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10005
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   480
      Left            =   9030
      Picture         =   "frmAnulaDocumentos.frx":0000
      TabIndex        =   9
      Top             =   5520
      Width           =   870
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Aceptar"
      Height          =   480
      Left            =   7260
      Picture         =   "frmAnulaDocumentos.frx":030A
      TabIndex        =   7
      Top             =   5520
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   480
      Left            =   8145
      Picture         =   "frmAnulaDocumentos.frx":0614
      TabIndex        =   8
      Top             =   5520
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid GrdModulos 
      Height          =   3750
      Left            =   45
      TabIndex        =   6
      Top             =   1455
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6615
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorSel    =   16761024
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameBuscar 
      Caption         =   "xxx..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   75
      TabIndex        =   10
      Top             =   30
      Width           =   9825
      Begin VB.TextBox txtDesCli 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3075
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Descripción"
         Top             =   270
         Width           =   4155
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   0
         Top             =   270
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscAprox 
         Caption         =   "Buscar"
         Height          =   420
         Left            =   7680
         MaskColor       =   &H000000FF&
         TabIndex        =   5
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   615
         Width           =   3630
      End
      Begin FechaCtl.Fecha FechaHasta 
         Height          =   285
         Left            =   4785
         TabIndex        =   4
         Top             =   975
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin FechaCtl.Fecha FechaDesde 
         Height          =   330
         Left            =   2280
         TabIndex        =   3
         Top             =   975
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         Separador       =   "/"
         Text            =   ""
         MensajeErrMin   =   "La fecha ingresada no alcanza el mínimo permitido"
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         Height          =   195
         Left            =   3735
         TabIndex        =   14
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label lblFechaDesde 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1185
         TabIndex        =   13
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1185
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   1185
         TabIndex        =   11
         Top             =   645
         Width           =   360
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   4395
      TabIndex        =   19
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Anulado"
      Height          =   195
      Left            =   5055
      TabIndex        =   18
      Top             =   5850
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pendiente"
      Height          =   195
      Left            =   5055
      TabIndex        =   17
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Definitivo"
      Height          =   195
      Left            =   5055
      TabIndex        =   16
      Top             =   5445
      Width           =   675
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   4380
      Top             =   5880
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   150
      Left            =   4380
      Top             =   5685
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   150
      Left            =   4380
      Top             =   5490
      Width           =   540
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
      Left            =   210
      TabIndex        =   15
      Top             =   5535
      Width           =   660
   End
End
Attribute VB_Name = "frmAnulaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TipodeAnulacion As Integer
Dim i As Integer

Private Sub CmdBuscAprox_Click()
    GrdModulos.Rows = 1
    BuscoFacturas
End Sub

Private Sub BuscoFacturas()
    lblEstado.Caption = "Buscando Facturas..."
    Screen.MousePointer = vbHourglass
    'poner sucursal
    SQL = "SELECT DISTINCT FC.FCL_NUMERO,FC.FCL_SUCURSAL,FC.FCL_FECHA, FC.EST_CODIGO, E.EST_DESCRI,"
    SQL = SQL & " C.CLI_CODIGO, C.CLI_RAZSOC, TC.TCO_ABREVIA, FC.TCO_CODIGO"
    SQL = SQL & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    SQL = SQL & " TIPO_COMPROBANTE TC, ESTADO_DOCUMENTO E"
    SQL = SQL & " WHERE"
    SQL = SQL & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    SQL = SQL & " AND FC.EST_CODIGO=E.EST_CODIGO"
    SQL = SQL & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    If txtCliente.Text <> "" Then
        SQL = SQL & " AND FC.CLI_CODIGO=" & XN(txtCliente.Text)
    End If
    If FechaDesde.Text <> "" Then
        SQL = SQL & " AND FC.FCL_FECHA>=" & XDQ(FechaDesde.Text)
    End If
    If FechaHasta.Text <> "" Then
        SQL = SQL & " AND FC.FCL_FECHA<=" & XDQ(FechaHasta.Text)
    End If
    If cboDocumento.List(cboDocumento.ListIndex) <> "(Todos)" Then
        SQL = SQL & " AND FC.TCO_CODIGO=" & XN(cboDocumento.ItemData(cboDocumento.ListIndex))
    End If
    SQL = SQL & " ORDER BY FC.FCL_FECHA,FC.FCL_NUMERO"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            GrdModulos.AddItem rec!TCO_ABREVIA & Chr(9) & Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000") & Chr(9) & rec!FCL_FECHA _
                            & Chr(9) & rec!CLI_RAZSOC & Chr(9) & rec!EST_DESCRI & Chr(9) & rec!EST_CODIGO _
                            & Chr(9) & rec!EST_CODIGO & Chr(9) & rec!TCO_CODIGO & Chr(9) & rec!CLI_CODIGO
                                                        
            If rec!EST_CODIGO = 2 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbRed)
            ElseIf rec!EST_CODIGO = 1 Then
                Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.Rows - 1, vbBlue)
            End If
            rec.MoveNext
        Loop
        GrdModulos.SetFocus
    Else
        lblEstado.Caption = ""
        Screen.MousePointer = vbNormal
        MsgBox "No se encontraron Facturas...", vbExclamation, TIT_MSGBOX
    End If
    rec.Close
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdGrabar_Click()
    If MsgBox("¿Confirma Anular?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    On Error GoTo SeClavo
    lblEstado.Caption = "Actualizando..."
    Screen.MousePointer = vbHourglass
    DBConn.BeginTrans
        
    ActualizoFactura
        
    DBConn.CommitTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    cmdNuevo_Click
    Exit Sub

SeClavo:
    DBConn.RollbackTrans
    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub ActualizoFactura()
    For i = 1 To GrdModulos.Rows - 1
        If GrdModulos.TextMatrix(i, 5) <> GrdModulos.TextMatrix(i, 6) Then 'PREGUNTA SI HUBO CAMBIO
            Set Rec2 = New ADODB.Recordset
            SQL = "SELECT FCL_TCO_CODIGO FROM FACTURAS_NOTA_CREDITO_CLIENTE"
            SQL = SQL & " WHERE"
            SQL = SQL & " FCL_TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
            SQL = SQL & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
            SQL = SQL & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
            Rec2.Open SQL, DBConn, adOpenStatic, adLockOptimistic
            
            If Rec2.EOF = True Then
                SQL = "UPDATE FACTURA_CLIENTE"
                SQL = SQL & " SET EST_CODIGO=" & XN(GrdModulos.TextMatrix(i, 6))
                SQL = SQL & " WHERE"
                SQL = SQL & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
                SQL = SQL & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
                SQL = SQL & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
                DBConn.Execute SQL
                
                'VUELVO ATRAS EL STOCK
                SQL = "SELECT PTO_CODIGO, DFC_CANTIDAD"
                SQL = SQL & " FROM DETALLE_FACTURA_CLIENTE"
                SQL = SQL & " WHERE"
                SQL = SQL & " TCO_CODIGO=" & XN(GrdModulos.TextMatrix(i, 7))
                SQL = SQL & " AND FCL_NUMERO=" & XN(Right(GrdModulos.TextMatrix(i, 1), 8))
                SQL = SQL & " AND FCL_SUCURSAL=" & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
                rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
                If rec.EOF = False Then
                    Do While rec.EOF = False
                        SQL = "UPDATE STOCK SET"
                        SQL = SQL & " DST_STKFIS = DST_STKFIS + " & XN(rec!DFC_CANTIDAD)
                        SQL = SQL & " WHERE STK_CODIGO = " & XN(Left(GrdModulos.TextMatrix(i, 1), 4))
                        SQL = SQL & " AND PTO_CODIGO = " & XN(rec!PTO_CODIGO)
                        DBConn.Execute SQL
                        rec.MoveNext
                    Loop
                End If
                rec.Close
            Else
                MsgBox "La Factura número: " & GrdModulos.TextMatrix(i, 1) & ", no puede ser ANULADA" & Chr(13) & _
                                           " por estar relacionada con una Nota de Crédito", vbCritical, TIT_MSGBOX
                GrdModulos_dblClick
            End If
            If Rec2.State = 1 Then Rec2.Close
        End If
    Next
End Sub

Private Sub CambiColoryEstado(Estado As Boolean)
    cboDocumento.Enabled = Estado
    If Estado = False Then
        cboDocumento.BackColor = &H8000000F
    Else
        cboDocumento.BackColor = &H80000005
    End If
End Sub

Private Sub cmdSalir_Click()
    Set frmAnulaDocumentos = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
End Sub

Private Sub Form_Load()
     Set rec = New ADODB.Recordset
     Set Rec2 = New ADODB.Recordset
     
    Me.Left = 0
    Me.Top = 0
    
    Select Case TipodeAnulacion
        Case 3 'FACTURAS
            frmAnulaDocumentos.Caption = "Anular Facturas"
            frameBuscar.Caption = "Buscar Facturas por..."
            'CARGO COMBO FACTURA
            LlenarComboFactura
            ConfiguroGrillaFactura
            Call CambiColoryEstado(True)
            
        Case 4 'RECIBOS
            frmAnulaDocumentos.Caption = "Anular Recibos"
            frameBuscar.Caption = "Buscar Recibos por..."
            'CARGO COMBO RECIBO
            LlenarComboRecibo
            ConfiguroGrillaRecibo
            Call CambiColoryEstado(True)
            
        Case 5 'NOTA DE CREDITO
            frmAnulaDocumentos.Caption = "Anular Nota de Crédito"
            frameBuscar.Caption = "Buscar Nota de Crédito por..."
            'CARGO COMBO NOTA DE CREDITO
            LlenarComboNotaCredito
            ConfiguroGrillaNotaDC
            Call CambiColoryEstado(True)
            
        Case 6 'NOTA DE DEBITO
            frmAnulaDocumentos.Caption = "Anular Nota de Débito"
            frameBuscar.Caption = "Buscar Nota de Débito por..."
            'CARGO COMBO NOTA DE DEBITO
            LlenarComboNotaDebito
            ConfiguroGrillaNotaDC
            Call CambiColoryEstado(True)
            
    End Select
    lblEstado.Caption = ""
End Sub

Private Sub LlenarComboNotaDebito()
    SQL = "SELECT * FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_DESCRI LIKE 'NOTA DE DEB%'"
    SQL = SQL & " ORDER BY TCO_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboNotaCredito()
    SQL = "SELECT * FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_DESCRI LIKE 'NOTA DE CRED%'"
    SQL = SQL & " ORDER BY TCO_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub ConfiguroGrillaRecibo()
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|^Estado|codigo estado|" _
                            & "codigo estado que cambio|TIPO RECIBO|COD CLIENTE"
                            
    GrdModulos.ColWidth(0) = 1000 'TIPO_NOTA
    GrdModulos.ColWidth(1) = 1300 'NRO RECIBO
    GrdModulos.ColWidth(2) = 1200 'FECHA_RECIBO
    GrdModulos.ColWidth(3) = 3900 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'TIPO RECIBO (TCO_CODIGO)
    GrdModulos.ColWidth(8) = 0    'CODIGO CLIENTE
    GrdModulos.Cols = 9
    GrdModulos.Rows = 2
    
End Sub

Private Sub ConfiguroGrillaNotaDC()
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|^Estado|codigo estado|" _
                            & "codigo estado QUE CAMBIO|TIPO Nota credito|COD CLIENTE|"
                            
    GrdModulos.ColWidth(0) = 1000 'TIPO_NOTA
    GrdModulos.ColWidth(1) = 1300 'NRO NOTA
    GrdModulos.ColWidth(2) = 1200 'FECHA
    GrdModulos.ColWidth(3) = 3900 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'TIPO nota credito (TCO_CODIGO)
    GrdModulos.ColWidth(8) = 0    'CODIGO CLIENTE
    GrdModulos.Cols = 9
    GrdModulos.Rows = 2
End Sub

Private Sub ConfiguroGrillaFactura()
    GrdModulos.FormatString = "^Tipo|^Número|^Fecha|Cliente|^Estado|codigo estado|" _
                            & "codigo estado QUE CAMBIO|TIPO FACTURA|COD CLIENTE"
                                                        
    GrdModulos.ColWidth(0) = 1000 'TIPO_NOTA
    GrdModulos.ColWidth(1) = 1300 'NRO FACTURA
    GrdModulos.ColWidth(2) = 1200 'FECHA_FACTURA
    GrdModulos.ColWidth(3) = 3900 'CLIENTE
    GrdModulos.ColWidth(4) = 2000 'ESTADO
    GrdModulos.ColWidth(5) = 0    'CODIGO ESTADO
    GrdModulos.ColWidth(6) = 0    'CODIGO ESTADO QUE CAMBIO
    GrdModulos.ColWidth(7) = 0    'TIPO FACTURA (TCO_CODIGO)
    GrdModulos.ColWidth(8) = 0    'CODIGO CLIENTE
    GrdModulos.Cols = 9
    GrdModulos.Rows = 2
End Sub

Private Sub LlenarComboFactura()
    SQL = "SELECT * FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_DESCRI LIKE 'FAC%'"
    SQL = SQL & " ORDER BY TCO_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub LlenarComboRecibo()
    SQL = "SELECT * FROM TIPO_COMPROBANTE"
    SQL = SQL & " WHERE TCO_DESCRI LIKE 'RECIB%'"
    SQL = SQL & " ORDER BY TCO_DESCRI"
    rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        cboDocumento.AddItem "(Todos)"
        Do While rec.EOF = False
            cboDocumento.AddItem rec!TCO_DESCRI
            cboDocumento.ItemData(cboDocumento.NewIndex) = rec!TCO_CODIGO
            rec.MoveNext
        Loop
        cboDocumento.ListIndex = 0
    End If
    rec.Close
End Sub

Private Sub GrdModulos_dblClick()
    If GrdModulos.Rows > 1 Then
        Select Case TipodeAnulacion
            Case 3 'FACTURAS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Factura" & Chr(13) & _
                           "el mimo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
                
            Case 4 'RECIBOS
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado al Recibo" & Chr(13) & _
                           "el mimo ya fue Anulado", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
            
            Case 5 'NOTA DE CREDITO
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Nota de Crédito" & Chr(13) & _
                           "la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
                
            Case 6 'NOTA DE DEBITO
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 5) = 2 Then
                    MsgBox "No se puede cambiar el estado a la Nota de Débito" & Chr(13) & _
                           "la misma ya fue Anulada", vbExclamation, TIT_MSGBOX
                    
                    Exit Sub
                End If
                If GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "ANULADO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbRed)
                     
                ElseIf GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 2 Then
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 6) = 3
                    GrdModulos.TextMatrix(GrdModulos.RowSel, 4) = "DEFINITIVO"
                    Call CambiaColorAFilaDeGrilla(GrdModulos, GrdModulos.RowSel, vbBlack)
                End If
        End Select
    End If
End Sub

Private Sub GrdModulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then GrdModulos_dblClick
End Sub

Private Sub cmdNuevo_Click()
    txtCliente.Text = ""
    txtDesCli.Text = ""
    FechaDesde.Text = ""
    FechaHasta.Text = ""
    cboDocumento.ListIndex = 0
    GrdModulos.Rows = 1
    GrdModulos.Rows = 2
    txtCliente.SetFocus
End Sub

Private Sub txtCliente_Change()
    If txtCliente.Text = "" Then
        txtDesCli.Text = ""
    End If
End Sub

Private Sub txtCliente_GotFocus()
    SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes txtCliente, "CODIGO"
    End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub

Private Sub txtCliente_LostFocus()
    If txtCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC"
        SQL = SQL & " FROM CLIENTE C"
        SQL = SQL & " WHERE"
        SQL = SQL & " CLI_CODIGO =" & XN(txtCliente.Text)
        If rec.State = 1 Then rec.Close
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            txtDesCli.Text = rec!CLI_RAZSOC
        Else
            MsgBox "El Cliente no existe", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Private Sub txtDesCli_Change()
    If txtDesCli.Text = "" Then
        txtCliente.Text = ""
    End If
End Sub

Private Sub txtDesCli_GotFocus()
    SelecTexto txtDesCli
End Sub

Private Sub txtDesCli_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtDesCli_LostFocus()
    If txtCliente.Text = "" And txtDesCli.Text <> "" Then
        Set rec = New ADODB.Recordset
        SQL = "SELECT C.CLI_CODIGO,C.CLI_RAZSOC"
        SQL = SQL & " FROM CLIENTE C"
        SQL = SQL & " WHERE"
        SQL = SQL & " CLI_RAZSOC LIKE '" & XN(Trim(txtDesCli.Text)) & "%'"
        rec.Open SQL, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes txtCliente, "CADENA", Trim(txtDesCli.Text)
                If rec.State = 1 Then rec.Close
                txtDesCli.SetFocus
            Else
                txtCliente.Text = rec!CLI_CODIGO
                txtDesCli.Text = rec!CLI_RAZSOC
            End If
        Else
            MsgBox "No se encontro el Cliente", vbExclamation, TIT_MSGBOX
            txtCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub

Public Sub BuscarClientes(Txt As Control, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código"
        .SQL = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            txtCliente.Text = .ResultFields(2)
            txtCliente_LostFocus
        End If
    End With
    
    Set B = Nothing
End Sub

