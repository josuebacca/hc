VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLibroIvaVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5625
   Begin VB.Frame FrameImpresora 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15
      TabIndex        =   12
      Top             =   1530
      Width           =   5595
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   195
         Width           =   1665
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "AgregarEstudioGinecologico.frx":0000
         Left            =   450
         List            =   "AgregarEstudioGinecologico.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   285
         Width           =   1635
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "AgregarEstudioGinecologico.frx":002F
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         Picture         =   "AgregarEstudioGinecologico.frx":0131
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         Picture         =   "AgregarEstudioGinecologico.frx":0233
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   3045
      TabIndex        =   2
      Top             =   2340
      Width           =   840
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   4755
      TabIndex        =   4
      Top             =   2340
      Width           =   840
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
      Height          =   1575
      Left            =   15
      TabIndex        =   5
      Top             =   -30
      Width           =   5595
      Begin VB.TextBox txtUltimo 
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   585
         Width           =   1125
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   75
         TabIndex        =   6
         Top             =   1095
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   556
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox FechaDesde 
         Height          =   300
         Left            =   1575
         ScaleHeight     =   240
         ScaleWidth      =   1095
         TabIndex        =   0
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   345
         TabIndex        =   10
         Top             =   285
         Width           =   600
      End
      Begin VB.Label lblPeriodo1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2730
         TabIndex        =   9
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   1155
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Último Nro Hoja:"
         Height          =   195
         Left            =   345
         TabIndex        =   7
         Top             =   645
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "AgregarEstudioGinecologico.frx":0335
      Height          =   405
      Left            =   3900
      TabIndex        =   3
      Top             =   2340
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1800
      Top             =   2325
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2145
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
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
      TabIndex        =   11
      Top             =   2400
      Width           =   660
   End
End
Attribute VB_Name = "frmLibroIvaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Registro As Long
Dim Tamanio As Long
Dim TotIva As Double
Dim ComiLiquidoProducto As Double

Private Sub CBImpresora_Click()
    CDImpresora.PrinterDefault = True
    CDImpresora.ShowPrinter
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
End Sub

Private Sub cmdAceptar_Click()
    Registro = 0
    Tamanio = 0
    TotIva = 0
    
    If FechaDesde.Text = "" Then
        MsgBox "Debe ingresar el periodo", vbExclamation, TIT_MSGBOX
        FechaDesde.SetFocus
        Exit Sub
    End If
    
    On Error GoTo CLAVO
    Screen.MousePointer = vbHourglass
    DBConn.BeginTrans
    lblEstado.Caption = "Buscando Datos..."
    
    'BORRO LA TABLA PRECIO_MEDICAMENTO
    sql = "DELETE FROM TMP_LIBRO_IVA_VENTAS"
    DBConn.Execute sql
    
    'FACTURAS
    BUSCO_FACTURAS
    'NOTAS DE CREDITO
    BUSCO_NOTA_CREDITO
    'NOTAS DE DEBITO
    BUSCO_NOTA_DEBITO
    'RETENCIONES
    BUSCO_RETENCIONES
    
    lblEstado.Caption = ""
    DBConn.CommitTrans
    'CARGO EL REPORTE
    ListarLibroIVA
    
    Screen.MousePointer = vbNormal
    Exit Sub

CLAVO:
 Screen.MousePointer = vbNormal
 lblEstado.Caption = ""
 DBConn.RollbackTrans
 If rec.State = 1 Then rec.Close
 MsgBox Err.Description, vbCritical, TIT_MSGBOX
End Sub

Private Sub BUSCO_FACTURAS()
    TotIva = 0
    'BUSCO FACTURAS POR REMITO ---------------------------------
    sql = "SELECT FC.FCL_NUMERO, FC.FCL_SUCURSAL, FC.FCL_FECHA, FC.FCL_IVA,"
    sql = sql & " FC.FCL_SUBTOTAL, FC.FCL_TOTAL,"
    sql = sql & " FC.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
    sql = sql & " C.CLI_RAZSOC, TC.TCO_ABREVIA"
    sql = sql & " FROM FACTURA_CLIENTE FC, CLIENTE C,"
    sql = sql & " TIPO_COMPROBANTE TC"
    sql = sql & " WHERE"
    sql = sql & " FC.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND FC.CLI_CODIGO=C.CLI_CODIGO"
    sql = sql & " AND FC.EST_CODIGO <> 1" 'ESTADO DEFINITIVO Y ANULADO
    If FechaDesde <> "" Then
        sql = sql & " AND YEAR (FC.FCL_FECHA)>=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND YEAR (FC.FCL_FECHA)<=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND MONTH (FC.FCL_FECHA)>=" & XN(Mid(FechaDesde, 4, 2))
        sql = sql & " AND MONTH (FC.FCL_FECHA)<=" & XN(Mid(FechaDesde, 4, 2))
    End If
    sql = sql & " ORDER BY FC.FCL_NUMERO, FC.FCL_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!FCL_FECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!FCL_SUCURSAL, "0000") & "-" & Format(rec!FCL_NUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & XN(rec!FCL_IVA) & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCIONES
            Else
                sql = sql & XN(Chk0(rec!FCL_SUBTOTAL)) & ","
                sql = sql & XN(rec!FCL_IVA) & ","
                TotIva = (CDbl(Chk0(rec!FCL_SUBTOTAL)) * CDbl(rec!FCL_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0(rec!FCL_TOTAL)) & ","
                sql = sql & "0" & ")" 'RETENCIONES
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCO_NOTA_CREDITO()
    TotIva = 0
    'BUSCO NOTA DE CREDITO------------------------------------
     sql = "SELECT NC.NCC_NUMERO, NC.NCC_SUCURSAL, NC.NCC_FECHA, NC.NCC_IVA,"
     sql = sql & " NC.NCC_SUBTOTAL, NC.NCC_TOTAL,"
     sql = sql & " NC.EST_CODIGO,C.CLI_CUIT,C.CLI_INGBRU,"
     sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
     sql = sql & " FROM NOTA_CREDITO_CLIENTE NC,"
     sql = sql & " TIPO_COMPROBANTE TC, CLIENTE C"
     sql = sql & " WHERE"
     sql = sql & " NC.TCO_CODIGO=TC.TCO_CODIGO"
     sql = sql & " AND NC.CLI_CODIGO=C.CLI_CODIGO"
     If FechaDesde <> "" Then
        sql = sql & " AND YEAR (NC.NCC_FECHA)>=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND YEAR (NC.NCC_FECHA)<=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND MONTH (NC.NCC_FECHA)>=" & XN(Mid(FechaDesde, 4, 2))
        sql = sql & " AND MONTH (NC.NCC_FECHA)<=" & XN(Mid(FechaDesde, 4, 2))
    End If
     sql = sql & " ORDER BY NC.NCC_FECHA"
     
     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!NCC_FECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!NCC_SUCURSAL, "0000") & "-" & Format(rec!NCC_NUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & XN(rec!NCC_IVA) & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCION
            Else
                sql = sql & XN(CStr((-1) * CDbl(rec!NCC_SUBTOTAL))) & ","
                sql = sql & XN(rec!NCC_IVA) & ","
                TotIva = (CDbl(rec!NCC_SUBTOTAL) * CDbl(rec!NCC_IVA)) / 100
                sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!NCC_TOTAL))) & ","
                sql = sql & "0" & ")" 'RETENCION
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCO_NOTA_DEBITO()
    TotIva = 0
    'BUSCO NOTA DE DEBITO SERVICIOS, CONCEPTO Y CHEQUES DEVUELTOS-----
    sql = "SELECT ND.NDC_NUMERO, ND.NDC_SUCURSAL, ND.NDC_FECHA, ND.NDC_IVA,"
    sql = sql & " ND.NDC_SUBTOTAL, ND.NDC_TOTAL,"
    sql = sql & " ND.EST_CODIGO, C.CLI_CUIT, C.CLI_INGBRU,"
    sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
    sql = sql & " FROM NOTA_DEBITO_CLIENTE ND,"
    sql = sql & " TIPO_COMPROBANTE TC , CLIENTE C"
    sql = sql & " WHERE ND.TCO_CODIGO=TC.TCO_CODIGO"
    sql = sql & " AND ND.CLI_CODIGO=C.CLI_CODIGO"
    If FechaDesde <> "" Then
        sql = sql & " AND YEAR (ND.NDC_FECHA)>=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND YEAR (ND.NDC_FECHA)<=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND MONTH (ND.NDC_FECHA)>=" & XN(Mid(FechaDesde, 4, 2))
        sql = sql & " AND MONTH (ND.NDC_FECHA)<=" & XN(Mid(FechaDesde, 4, 2))
    End If
    sql = sql & " ORDER BY ND.NDC_FECHA"
    
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!NDC_FECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!NDC_SUCURSAL, "0000") & "-" & Format(rec!NDC_NUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & XN(rec!NDC_IVA) & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCION
            Else
                sql = sql & XN(rec!NDC_SUBTOTAL) & ","
                sql = sql & XN(rec!NDC_IVA) & ","
                TotIva = (CDbl(rec!NDC_SUBTOTAL) * CDbl(rec!NDC_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(rec!NDC_TOTAL) & ","
                sql = sql & "0" & ")" 'RETENCION
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub BUSCO_RETENCIONES()
    TotIva = 0
    'BUSCO NOTA DE CREDITO------------------------------------
     sql = "SELECT DR.DRE_COMNUMERO, DR.DRE_COMSUCURSAL, DR.DRE_COMFECHA,"
     sql = sql & " DR.DRE_COMIMP, R.EST_CODIGO, C.CLI_CUIT,C.CLI_INGBRU,"
     sql = sql & " C.CLI_RAZSOC,TC.TCO_ABREVIA"
     sql = sql & " FROM RECIBO_CLIENTE R,DETALLE_RECIBO_CLIENTE DR"
     sql = sql & ",TIPO_COMPROBANTE TC , CLIENTE C"
     sql = sql & " WHERE"
     sql = sql & " R.TCO_CODIGO=DR.TCO_CODIGO"
     sql = sql & " AND R.REC_NUMERO=DR.REC_NUMERO"
     sql = sql & " AND R.REC_SUCURSAL=DR.REC_SUCURSAL"
     sql = sql & " AND DR.DRE_TCO_CODIGO=TC.TCO_CODIGO"
     sql = sql & " AND DR.DRE_TCO_CODIGO IN (14,15,16)" 'LAS TRES RETENCIONES
     sql = sql & " AND R.CLI_CODIGO=C.CLI_CODIGO"
     If FechaDesde <> "" Then
        sql = sql & " AND YEAR (R.REC_FECHA)>=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND YEAR (R.REC_FECHA)<=" & XN(Mid(FechaDesde, 7, 4))
        sql = sql & " AND MONTH (R.REC_FECHA)>=" & XN(Mid(FechaDesde, 4, 2))
        sql = sql & " AND MONTH (R.REC_FECHA)<=" & XN(Mid(FechaDesde, 4, 2))
    End If
     sql = sql & " ORDER BY R.REC_FECHA"
     
     rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Registro = 0
        Tamanio = rec.RecordCount
        Do While rec.EOF = False
            sql = "INSERT INTO TMP_LIBRO_IVA_VENTAS (FECHA,COMPROBANTE,NUMERO,"
            sql = sql & "CLIENTE,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,TOTAL,RETENCION)"
            sql = sql & "VALUES ("
            sql = sql & XDQ(rec!DRE_COMFECHA) & ","
            sql = sql & XS(rec!TCO_ABREVIA) & ","
            sql = sql & XS(Format(rec!DRE_COMSUCURSAL, "0000") & "-" & Format(rec!DRE_COMNUMERO, "00000000")) & ","
            sql = sql & XS(rec!CLI_RAZSOC) & ","
            sql = sql & XS(Format(ChkNull(rec!CLI_CUIT), "##-########-#")) & ","
            sql = sql & "NULL" & ","
            'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
            If rec!EST_CODIGO = 2 Then
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ")" 'RETENCION
            Else
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & "0" & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!DRE_COMIMP))) & "," 'TOTAL
                sql = sql & XN(CStr((-1) * CDbl(rec!DRE_COMIMP))) & ")" 'RETENCION
            End If
            DBConn.Execute sql
            rec.MoveNext
            
            Registro = Registro + 1
            ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
            lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
        Loop
    End If
    rec.Close
End Sub

Private Sub ListarLibroIVA()
    lblEstado.Caption = "Buscando Listado..."
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    Rep.Formulas(5) = ""
            
    sql = "SELECT CUIT,DIRECCION,RAZ_SOCIAL FROM PARAMETROS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Rep.Formulas(0) = "EMPRESA='     Empresa:  " & Trim(rec!RAZ_SOCIAL) & "'"
        Rep.Formulas(1) = "CUIT='       C.U.I.T.:  " & Format(rec!cuit, "##-########-#") & "'"
        Rep.Formulas(2) = "DIRECCION='    Dirección:  " & Trim(rec!DIRECCION) & "'"
    End If
    rec.Close
    
    If txtUltimo.Text = "" Then
        txtUltimo.Text = "0"
    End If
        Rep.Formulas(3) = "Numero='" & Trim(txtUltimo.Text) & "'"
        Rep.Formulas(4) = "PERIODO='" & Trim(lblPeriodo1.Caption) & "'"
    
    sql = "SELECT * FROM TMP_LIBRO_IVA_VENTAS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = True Then
        Rep.Formulas(5) = "NumeroHoja='" & CStr(CInt(txtUltimo.Text) + 1) & "'"
    End If
    rec.Close
    
    Rep.WindowTitle = "Libro I.V.A. Ventas"
    Rep.ReportFileName = DRIVE & DirReport & "libro_iva_ventas.rpt"
    
'    If optPantalla.Value = True Then
'        Rep.Destination = crptToWindow
'    ElseIf optImpresora.Value = True Then
'        Rep.Destination = crptToPrinter
'    End If
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2
            Rep.Destination = crptToFile
    End Select
    Rep.Action = 1
    
    lblEstado.Caption = ""
    Rep.Formulas(0) = ""
    Rep.Formulas(1) = ""
    Rep.Formulas(2) = ""
    Rep.Formulas(3) = ""
    Rep.Formulas(4) = ""
    Rep.Formulas(5) = ""
End Sub

Private Sub CmdNuevo_Click()
    FechaDesde.Text = ""
    lblPeriodo1.Caption = ""
    txtUltimo.Text = ""
    FechaDesde.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroIvaVentas = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then MySendKeys Chr(9)
  If KeyAscii = vbKeyEscape Then CmdSalir_Click
End Sub

Private Sub Form_Load()
    lblEstado.Caption = ""
    lblPor.Caption = "100 %"
    Me.Left = 0
    Me.Top = 0
    Set rec = New ADODB.Recordset
    FrameImpresora.Caption = "Impresora Actual: " & Printer.DeviceName
    cboDestino.ListIndex = 0
End Sub

Private Sub FechaDesde_LostFocus()
    If Trim(FechaDesde.Text) <> "" Then
        lblPeriodo1.Caption = UCase(Format(FechaDesde.Text, "mmmm/yyyy"))
    Else
        lblPeriodo1.Caption = ""
    End If
End Sub

Private Sub txtUltimo_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
