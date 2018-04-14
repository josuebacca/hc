VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLibroIvaCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro IVA Compras"
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
      Left            =   30
      TabIndex        =   12
      Top             =   1545
      Width           =   5595
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
         Picture         =   "frmAgregarEstudioGinecologico.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   17
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
         Picture         =   "frmAgregarEstudioGinecologico.frx":0102
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
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
         Index           =   0
         Left            =   135
         Picture         =   "frmAgregarEstudioGinecologico.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmAgregarEstudioGinecologico.frx":0306
         Left            =   450
         List            =   "frmAgregarEstudioGinecologico.frx":0313
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   285
         Width           =   1635
      End
      Begin VB.CommandButton CBImpresora 
         Caption         =   "Configurar Impresora"
         Height          =   435
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   195
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "frmAgregarEstudioGinecologico.frx":0335
      Height          =   405
      Left            =   3915
      TabIndex        =   3
      Top             =   2325
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
      Left            =   30
      TabIndex        =   6
      Top             =   -15
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
         TabIndex        =   9
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Último Nro Hoja:"
         Height          =   195
         Left            =   345
         TabIndex        =   11
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         Height          =   195
         Left            =   4920
         TabIndex        =   10
         Top             =   1155
         Width           =   480
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
         TabIndex        =   8
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   345
         TabIndex        =   7
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   4770
      TabIndex        =   4
      Top             =   2325
      Width           =   840
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   3060
      TabIndex        =   2
      Top             =   2325
      Width           =   840
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   1875
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   2415
      Top             =   2130
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
      Left            =   105
      TabIndex        =   5
      Top             =   2385
      Width           =   660
   End
End
Attribute VB_Name = "frmLibroIvaCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Registro As Long
Dim Tamanio As Long
Dim TotIva As Double

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
     
        'BORRO LA TABLA TMP_LIBRO_IVA_COMPRAS
        sql = "DELETE FROM TMP_LIBRO_IVA_COMPRAS"
        DBConn.Execute sql
        
        'BUSCO FACTURAS
        sql = "SELECT FP.FPR_NROSUCTXT,FP.FPR_NUMEROTXT,"
        sql = sql & " FP.FPR_FECHA,FP.FPR_IVA,FP.FPR_NETO,FP.FPR_TOTAL,"
        sql = sql & " FP.FPR_IVA1,FP.FPR_NETO1,FP.FPR_IMPUESTOS,"
        sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
        sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA,FP.TCO_CODIGO,"
        sql = sql & " FP.FPR_PEAJETOT,FP.FPR_NETOGRAV,FP.FPR_IVACF,FP.FPR_PERCEPCION"
        sql = sql & " FROM FACTURA_PROVEEDOR FP, PROVEEDOR P"
        sql = sql & " ,TIPO_COMPROBANTE TC"
        sql = sql & " WHERE"
        sql = sql & " FP.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND FP.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND FP.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND FP.FPR_LIBROIVA='S'"
        If FechaDesde <> "" Then
            sql = sql & " AND YEAR (FP.FPR_PERIODO)>=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND YEAR (FP.FPR_PERIODO)<=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND MONTH (FP.FPR_PERIODO)>=" & XN(Mid(FechaDesde, 4, 2))
            sql = sql & " AND MONTH (FP.FPR_PERIODO)<=" & XN(Mid(FechaDesde, 4, 2))
        End If
        sql = sql & " ORDER BY FP.FPR_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,RS_RNI,PERCEPCION,IMPUESTOS,TOTAL)"
                sql = sql & " VALUES ("
                sql = sql & XDQ(rec!FPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!FPR_NROSUCTXT & "-" & rec!FPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!FPR_PEAJETOT = 0 Then
                    If rec!TCO_CODIGO = 3 Then 'FACTURAS C
                        sql = sql & XN("0") & ","
                        sql = sql & XN(rec!FPR_IVA) & ","
                            TotIva = (CDbl("0") * CDbl(rec!FPR_IVA)) / 100
                        sql = sql & XN(CStr(TotIva)) & "," 'IVA BUENO
                        sql = sql & XN(Chk0(rec!FPR_NETO1)) & ","
                            TotIva = (CDbl(Chk0(rec!FPR_NETO1)) * CDbl(Chk0(rec!FPR_IVA1))) / 100
                        sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                        sql = sql & XN(Chk0(rec!FPR_NETO)) & "," 'FACTURAS C
                        sql = sql & XN(Chk0(rec!FPR_PERCEPCION)) & "," 'PERCEPCIONES
                        sql = sql & XN("0") & ","
                        sql = sql & XN(rec!FPR_TOTAL) & ")"
                    Else 'OTROS COMPROBANTES
                        sql = sql & XN(rec!FPR_NETO) & ","
                        sql = sql & XN(rec!FPR_IVA) & ","
                            TotIva = (CDbl(rec!FPR_NETO) * CDbl(rec!FPR_IVA)) / 100
                        sql = sql & XN(CStr(TotIva)) & "," 'IVA BUENO
                        sql = sql & XN(Chk0(rec!FPR_NETO1)) & ","
                            TotIva = (CDbl(Chk0(rec!FPR_NETO1)) * CDbl(Chk0(rec!FPR_IVA1))) / 100
                        sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                        sql = sql & XN("0") & "," 'PARA FACTURAS C
                        sql = sql & XN(Chk0(rec!FPR_PERCEPCION)) & "," 'PERCEPCIONES
                        sql = sql & XN(Chk0(rec!FPR_IMPUESTOS)) & ","
                        sql = sql & XN(rec!FPR_TOTAL) & ")"
                    End If
                    
                Else 'PEAJE
                    sql = sql & XN(rec!FPR_NETOGRAV) & ","
                    sql = sql & XN(rec!FPR_IVA) & ","
                    sql = sql & XN(rec!FPR_IVACF) & ","
                    sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                    sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                    sql = sql & XN("0") & "," 'PARA FACTURAS C
                    sql = sql & XN(Chk0(rec!FPR_PERCEPCION)) & "," 'PERCEPCIONES
                    sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                    sql = sql & XN(rec!FPR_PEAJETOT) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
                lblPor.Refresh
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE CREDITO------------------------------------
         sql = "SELECT NP.CPR_NROSUCTXT,NP.CPR_NUMEROTXT,NP.CPR_FECHA,NP.CPR_IVA,NP.CPR_NETO,NP.CPR_TOTAL,"
         sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
         sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA"
         sql = sql & " FROM NOTA_CREDITO_PROVEEDOR NP"
         sql = sql & ",TIPO_COMPROBANTE TC , PROVEEDOR P"
         sql = sql & " WHERE"
         sql = sql & " NP.TCO_CODIGO=TC.TCO_CODIGO"
         sql = sql & " AND NP.TPR_CODIGO=P.TPR_CODIGO"
         sql = sql & " AND NP.PROV_CODIGO=P.PROV_CODIGO"
         If FechaDesde <> "" Then
            sql = sql & " AND YEAR (NP.CPR_PERIODO)>=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND YEAR (NP.CPR_PERIODO)<=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND MONTH (NP.CPR_PERIODO)>=" & XN(Mid(FechaDesde, 4, 2))
            sql = sql & " AND MONTH (NP.CPR_PERIODO)<=" & XN(Mid(FechaDesde, 4, 2))
        End If
         sql = sql & " ORDER BY NP.CPR_FECHA"
         
         rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,RS_RNI,PERCEPCION,IMPUESTOS,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!CPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!CPR_NROSUCTXT & "-" & rec!CPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_NETO))) & ","
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_IVA))) & ","
                TotIva = (CDbl(rec!CPR_NETO) * CDbl(rec!CPR_IVA)) / 100
                sql = sql & XN(CStr((-1) * CDbl(TotIva))) & ","
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0("")) & "," 'PARA FCATURAS C
                sql = sql & XN(Chk0("")) & "," 'PERCEPCIONES
                sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                sql = sql & XN(CStr((-1) * CDbl(rec!CPR_TOTAL))) & ")"

                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
                lblPor.Refresh
            Loop
        End If
        rec.Close
        
        'BUSCO NOTA DE DEBITO SERVICIOS Y CHEQUES DEVUELTOS-----
        sql = "SELECT NP.DPR_NROSUCTXT,NP.DPR_NUMEROTXT,NP.DPR_FECHA,NP.DPR_IVA,NP.DPR_NETO,NP.DPR_TOTAL,"
        sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
        sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA"
        sql = sql & " FROM NOTA_DEBITO_PROVEEDOR NP,"
        sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P"
        sql = sql & " WHERE NP.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND NP.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND NP.PROV_CODIGO=P.PROV_CODIGO"
        If FechaDesde <> "" Then
            sql = sql & " AND YEAR (NP.DPR_PERIODO)>=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND YEAR (NP.DPR_PERIODO)<=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND MONTH (NP.DPR_PERIODO)>=" & XN(Mid(FechaDesde, 4, 2))
            sql = sql & " AND MONTH (NP.DPR_PERIODO)<=" & XN(Mid(FechaDesde, 4, 2))
        End If
        sql = sql & " ORDER BY NP.DPR_FECHA"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,RS_RNI,PERCEPCION,IMPUESTOS,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!DPR_FECHA) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!DPR_NROSUCTXT & "-" & rec!DPR_NUMEROTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                sql = sql & XN(rec!DPR_NETO) & ","
                sql = sql & XN(rec!DPR_IVA) & ","
                TotIva = (CDbl(rec!DPR_NETO) * CDbl(rec!DPR_IVA)) / 100
                sql = sql & XN(CStr(TotIva)) & ","
                sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                sql = sql & XN(Chk0("")) & "," 'PARA FCATURAS C
                sql = sql & XN(Chk0("")) & "," 'PERCEPCIONES
                sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                sql = sql & XN(rec!DPR_TOTAL) & ")"
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
                lblPor.Refresh
            Loop
        End If
        rec.Close
        
        'BUSCO COMPROBANTES DENTRO DE LOS GASTOS GENERALES -----
        sql = "SELECT GG.GGR_NROSUCTXT,GG.GGR_NROCOMPTXT,GG.GGR_FECHACOMP,GG.GGR_IVA,GG.GGR_NETO,GG.GGR_TOTAL,"
        sql = sql & " GG.GGR_IVA1,GG.GGR_NETO1,GG.GGR_IMPUESTOS,"
        sql = sql & " P.PROV_CUIT,P.PROV_INGBRU,"
        sql = sql & " P.PROV_RAZSOC,TC.TCO_ABREVIA,GG.TCO_CODIGO,"
        sql = sql & " GG.GGR_PEAJETOT,GG.GGR_NETOGRAV,GG.GGR_IVACF,GG.GGR_PERCEPCION"
        sql = sql & " FROM GASTOS_GENERALES GG,"
        sql = sql & " TIPO_COMPROBANTE TC, PROVEEDOR P"
        sql = sql & " WHERE GG.TCO_CODIGO=TC.TCO_CODIGO"
        sql = sql & " AND GG.TPR_CODIGO=P.TPR_CODIGO"
        sql = sql & " AND GG.PROV_CODIGO=P.PROV_CODIGO"
        sql = sql & " AND GG.GGR_LIBROIVA='S'"
        If FechaDesde <> "" Then
            sql = sql & " AND YEAR (GG.GGR_PERIODO)>=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND YEAR (GG.GGR_PERIODO)<=" & XN(Mid(FechaDesde, 7, 4))
            sql = sql & " AND MONTH (GG.GGR_PERIODO)>=" & XN(Mid(FechaDesde, 4, 2))
            sql = sql & " AND MONTH (GG.GGR_PERIODO)<=" & XN(Mid(FechaDesde, 4, 2))
        End If
        sql = sql & " ORDER BY GG.GGR_FECHACOMP"
        
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            Registro = 0
            Tamanio = rec.RecordCount
            Do While rec.EOF = False
                sql = "INSERT INTO TMP_LIBRO_IVA_COMPRAS (FECHA,COMPROBANTE,NUMERO,"
                sql = sql & "PROVEEDOR,CUIT,INGBRUTOS,SUBTOTAL,IVA,TOTIVA,"
                sql = sql & "SUBTOTAL1,TOTOTROIVA,RS_RNI,PERCEPCION,IMPUESTOS,TOTAL)"
                sql = sql & "VALUES ("
                sql = sql & XDQ(rec!GGR_FECHACOMP) & ","
                sql = sql & XS(rec!TCO_ABREVIA) & ","
                sql = sql & XS(rec!GGR_NROSUCTXT & "-" & rec!GGR_NROCOMPTXT) & ","
                sql = sql & XS(rec!PROV_RAZSOC) & ","
                sql = sql & XS(Format(rec!PROV_CUIT, "##-########-#")) & ","
                sql = sql & "NULL" & ","
                'sql = sql & XS(Format(rec!CLI_INGBRU, "###-#####-##")) & ","
                If rec!GGR_PEAJETOT = 0 Then
                    If rec!TCO_CODIGO = 3 Then 'FACTURAS C
                        sql = sql & XN("0") & ","
                        sql = sql & XN(rec!GGR_IVA) & ","
                            TotIva = (CDbl("0") * CDbl(rec!GGR_IVA)) / 100
                        sql = sql & XN(CStr(TotIva)) & ","
                        sql = sql & XN(Chk0(rec!GGR_NETO1)) & "," 'OTRO NETO
                            TotIva = (CDbl(Chk0(rec!GGR_NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                        sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                        sql = sql & XN(rec!GGR_NETO) & "," 'FACTURA C
                        sql = sql & XN(Chk0(rec!GGR_PERCEPCION)) & "," 'PERCEPCIONES
                        sql = sql & XN("0") & "," 'IMPUESTOS
                        sql = sql & XN(rec!GGR_TOTAL) & ")"
                                                
                    Else 'OTROS COMPROBANTES
                        sql = sql & XN(rec!GGR_NETO) & ","
                        sql = sql & XN(rec!GGR_IVA) & ","
                            TotIva = (CDbl(rec!GGR_NETO) * CDbl(rec!GGR_IVA)) / 100
                        sql = sql & XN(CStr(TotIva)) & ","
                        sql = sql & XN(Chk0(rec!GGR_NETO1)) & "," 'OTRO NETO
                            TotIva = (CDbl(Chk0(rec!GGR_NETO1)) * CDbl(Chk0(rec!GGR_IVA1))) / 100
                        sql = sql & XN(CStr(TotIva)) & "," 'OTRO IVA
                        sql = sql & XN("0") & "," 'PARA FACTURAS C
                        sql = sql & XN(Chk0(rec!GGR_PERCEPCION)) & "," 'PERCEPCIONES
                        sql = sql & XN(Chk0(rec!GGR_IMPUESTOS)) & "," 'IMPUESTOS
                        sql = sql & XN(rec!GGR_TOTAL) & ")"
                    End If
                    
                Else 'PEAJES
                    sql = sql & XN(rec!GGR_NETOGRAV) & ","
                    sql = sql & XN(rec!GGR_IVA) & ","
                    sql = sql & XN(rec!GGR_IVACF) & ","
                    sql = sql & XN(Chk0("")) & "," 'OTRO NETO
                    sql = sql & XN(Chk0("")) & "," 'OTRO IVA
                    sql = sql & XN("0") & "," 'PARA FACTURAS C
                    sql = sql & XN(Chk0(rec!GGR_PERCEPCION)) & "," 'PERCEPCIONES
                    sql = sql & XN(Chk0("")) & "," 'IMPUESTOS
                    sql = sql & XN(rec!GGR_PEAJETOT) & ")"
                End If
                DBConn.Execute sql
                rec.MoveNext
                
                Registro = Registro + 1
                ProgressBar1.Value = Format((Registro * 100) / Tamanio, "0.0")
                lblPor.Caption = Format((Registro * 100) / Tamanio, "0.0") & " %"
                lblPor.Refresh
            Loop
        End If
        rec.Close
        
    lblEstado.Caption = ""
    DBConn.CommitTrans
    'cargo el reporte
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
    
    sql = "SELECT * FROM TMP_LIBRO_IVA_COMPRAS"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = True Then
        Rep.Formulas(5) = "NumeroHoja='" & CStr(CInt(txtUltimo.Text) + 1) & "'"
    End If
    rec.Close
    
    Rep.WindowTitle = "Libro I.V.A. Compras"
    Rep.ReportFileName = DRIVE & DirReport & "libro_iva_compras.rpt"
    
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
    ProgressBar1.Value = 0
    FechaDesde.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Set frmLibroIvaCompras = Nothing
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
