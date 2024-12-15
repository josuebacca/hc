VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmImprime 
   Caption         =   "Reporte"
   ClientHeight    =   2925
   ClientLeft      =   1530
   ClientTop       =   1755
   ClientWidth     =   6150
   ClipControls    =   0   'False
   Icon            =   "Imprime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   6150
   Begin VB.TextBox txtEmpresaDire 
      Height          =   345
      Left            =   5685
      TabIndex        =   38
      Top             =   1860
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   315
      Left            =   5520
      Picture         =   "Imprime.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   360
      Width           =   330
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   60
      TabIndex        =   34
      Top             =   60
      Width           =   5970
      Begin VB.CheckBox chkApertura 
         Caption         =   "Imprime Apertura Impositiva?"
         Height          =   240
         Left            =   2955
         TabIndex        =   37
         Top             =   555
         Width           =   2745
      End
      Begin VB.CheckBox chkMembrete 
         Caption         =   "Usa Hoja Membretada?"
         Height          =   240
         Left            =   2955
         TabIndex        =   1
         Top             =   285
         Width           =   2745
      End
      Begin VB.TextBox txtInicial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1065
         TabIndex        =   0
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Nro. Inicial:"
         Height          =   255
         Left            =   210
         TabIndex        =   35
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2250
      Left            =   6735
      TabIndex        =   24
      Top             =   210
      Visible         =   0   'False
      Width           =   6915
      Begin VB.TextBox txtEmpresaCuit 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   32
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox txtEmp_Id 
         Height          =   315
         Left            =   4365
         TabIndex        =   31
         Top             =   1065
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame3 
         Height          =   75
         Left            =   90
         TabIndex        =   30
         Top             =   1560
         Width           =   6795
      End
      Begin VB.TextBox txtTipoLibro 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4365
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtEmpresa 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   28
         Top             =   375
         Width           =   3075
      End
      Begin VB.TextBox txtMes_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   2
         TabIndex        =   8
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtAnio_LibroI 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1725
         MaxLength       =   4
         TabIndex        =   9
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txtLibro_IdI 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4380
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "C.U.I.T."
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Empresa"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1005
         Width           =   540
      End
   End
   Begin VB.Frame fraNumCopias 
      Caption         =   "Número de Copias"
      Height          =   735
      Left            =   2250
      TabIndex        =   19
      Top             =   1080
      Width           =   1590
      Begin VB.TextBox tNumCopias 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   810
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "Imprime.frx":058C
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copias"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   20
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Páginas"
      Height          =   735
      Left            =   3840
      TabIndex        =   12
      Top             =   1080
      Width           =   2220
      Begin VB.TextBox tPagDesde 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   435
      End
      Begin VB.TextBox tPagHasta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1665
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   225
         Left            =   1170
         TabIndex        =   13
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.Frame fraImpresion 
      Caption         =   "Destino"
      Height          =   735
      Left            =   75
      TabIndex        =   11
      Top             =   1080
      Width           =   2175
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   135
         Picture         =   "Imprime.frx":058E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   135
         Picture         =   "Imprime.frx":0690
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSalida 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "Imprime.frx":0792
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         Top             =   315
         Width           =   240
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "Imprime.frx":0894
         Left            =   450
         List            =   "Imprime.frx":089E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CommandButton cCancelar 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   4965
      Picture         =   "Imprime.frx":08B7
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2115
      Width           =   1020
   End
   Begin VB.CommandButton cAceptar 
      Caption         =   "&Aceptar"
      Height          =   750
      Left            =   3915
      Picture         =   "Imprime.frx":0BC1
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2115
      Width           =   1020
   End
   Begin VB.CommandButton CBImpresora 
      Caption         =   "Cambiar Im&presora"
      Height          =   750
      Left            =   105
      Picture         =   "Imprime.frx":0ECB
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2115
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog CDImpresora 
      Left            =   1965
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   2565
      Top             =   2445
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label LBImpActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   18
      Top             =   1860
      Width           =   5940
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00-00/00/0000"
      Height          =   285
      Left            =   30
      TabIndex        =   17
      Top             =   2985
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Modo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6690
      TabIndex        =   16
      Top             =   2985
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label BE 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Top             =   2985
      Visible         =   0   'False
      Width           =   5100
   End
End
Attribute VB_Name = "frmImprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rptNombre As String
Private snp As ADODB.Recordset
Private rptEst(4, 30) As String * 50
Private Frm As Form
Private sSQL(3) As String
Private rptTitulo As String
Private Vista As String
Dim mTComp As String


Private Sub Apertura_Impositiva()
    
    Dim mNetoA As String
    Dim mNetoB As String
    Dim mIIA As String
    Dim mIIB As String
    Dim mInIA As String
    Dim mInIB As String
    Dim mExeA As String
    Dim mExeB As String
    
    Dim mTotalA As String
    Dim mTotalB As String
    
    Dim recApertura As ADODB.Recordset
    Set recApertura = New ADODB.Recordset
    
    'saco los totales de Neto
    cSQL = "SELECT sum (netograv) as netoA FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'A' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!netoa) = "" Then
        mNetoA = Format("0", "0.00")
    Else
        mNetoA = Format(recApertura!netoa, "0.00")
    End If
    recApertura.Close
    
    cSQL = "SELECT sum (netograv) as netoB FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'B' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!netob) = "" Then
        mNetoB = Format("0", "0.00")
    Else
        mNetoB = Format(recApertura!netob, "0.00")
    End If
    recApertura.Close
    
    'saco los totales de Iva Inscripto
    cSQL = "SELECT sum (ivai) as ivaiA FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'A' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!ivaia) = "" Then
        mIIA = Format("0", "0.00")
    Else
        mIIA = Format(recApertura!ivaia, "0.00")
    End If
    recApertura.Close
    
    cSQL = "SELECT sum (ivai) as ivaiB FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'B' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!ivaib) = "" Then
        mIIB = Format("0", "0.00")
    Else
        mIIB = Format(recApertura!ivaib, "0.00")
    End If
    recApertura.Close
    
    
    'saco los totales de Iva NI
    cSQL = "SELECT sum (ivani) as ivaniA FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'A' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!ivania) = "" Then
        mInIA = Format("0", "0.00")
    Else
        mInIA = Format(recApertura!ivania, "0.00")
    End If
    recApertura.Close
    
    cSQL = "SELECT sum (ivani) as ivaniB FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'B' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!ivanib) = "" Then
        mInIB = Format("0", "0.00")
    Else
        mInIB = Format(recApertura!ivanib, "0.00")
    End If
    recApertura.Close
    
    'saco exentos
    cSQL = "SELECT sum (exento) as exentoA FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'A' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!exentoa) = "" Then
        mExeA = Format("0", "0.00")
    Else
        mExeA = Format(recApertura!exentoa, "0.00")
    End If
    recApertura.Close
    
    cSQL = "SELECT sum (exento) as exentoB FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'B' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!exentob) = "" Then
        mExeB = Format("0", "0.00")
    Else
        mExeB = Format(recApertura!exentob, "0.00")
    End If
    recApertura.Close


    'saco los totales generales
    cSQL = "SELECT sum (total) as totalA FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'A' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!totala) = "" Then
        mTotalA = Format("0", "0.00")
    Else
        mTotalA = Format(recApertura!totala, "0.00")
    End If
    recApertura.Close
    
    cSQL = "SELECT sum (total) as totalB FROM libro_deta WHERE libro_id = " & XN(txtLibro_IdI.Text)
    cSQL = cSQL & " and tipo_doc = 'B' and tc_codigo = " & XN(mTComp)
    recApertura.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If ChkNull(recApertura!totalb) = "" Then
        mTotalB = Format("0", "0.00")
    Else
        mTotalB = Format(recApertura!totalb, "0.00")
    End If
    recApertura.Close

    rpt.Formulas(0) = ""
    rpt.Formulas(1) = ""
    rpt.Formulas(2) = ""
    rpt.Formulas(3) = ""
    rpt.Formulas(4) = ""
    rpt.Formulas(5) = ""
    rpt.Formulas(6) = ""
    rpt.Formulas(7) = ""
    
    If txtTipoLibro.Text = "V" Then
        rpt.Formulas(0) = "NETOINS='" & mNetoA & "'"
        rpt.Formulas(1) = "NETOCF='" & mNetoB & "'"
        
        rpt.Formulas(2) = "IIINS='" & mIIA & "'"
        rpt.Formulas(3) = "IICF='" & mIIB & "'"
        
        rpt.Formulas(4) = "INIINS='" & mInIA & "'"
        rpt.Formulas(5) = "INICF='" & mInIB & "'"
        
        rpt.Formulas(6) = "EXEINS='" & mInIA & "'"
        rpt.Formulas(7) = "EXECF='" & mInIB & "'"
        
        rpt.Formulas(8) = "TOTALINS='" & mTotalA & "'"
        rpt.Formulas(9) = "TOTALCF='" & mTotalB & "'"
    End If

End Sub



Private Sub cAceptar_Click()
    Dim j As Integer
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    'Dim mTComp As String
    
    'me fijo que tipo de comprobantes son las facturas
    cSQL = "select tc_codigo from tipo_comprobante where tc_descri like '%fact%'"
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
        mTComp = rec!tc_codigo
    Else
        mTComp = "1"
    End If
    rec.Close
    
    Dim mHayDatos As Boolean
    'me fijo que tipo de comprobantes son las facturas
    cSQL = "select * from libro_iva where tipo_libro = " & XS(txtTipoLibro.Text)
    cSQL = cSQL & " and LIBRO_ID=" & txtLibro_IdI.Text
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        mHayDatos = True
    Else
        mHayDatos = False
    End If
    rec.Close
    
    rpt.SelectionFormula = " {LIBRO_IVA.TIPO_LIBRO}=" & XS(txtTipoLibro.Text) & " AND {LIBRO_IVA.LIBRO_ID}=" & txtLibro_IdI.Text
    rpt.Formulas(0) = ""
    rpt.Formulas(1) = ""
    rpt.Formulas(2) = ""
    rpt.Formulas(3) = ""
    rpt.Formulas(4) = ""
    rpt.Formulas(5) = ""
    rpt.Formulas(6) = ""
    rpt.Formulas(7) = ""
    
    rpt.Formulas(2) = "XANIO = '" & txtAnio_LibroI.Text & "'"
    rpt.Formulas(3) = "XMES = '" & txtMes_LibroI.Text & "'"
    rpt.Formulas(4) = "XEMPDESCRI = '" & txtEmpresa.Text & "'"
    rpt.Formulas(5) = "XEMPDIRE = '" & txtEmpresaDire.Text & "'"
    rpt.Formulas(6) = "XEMPCUIT = '" & txtEmpresaCuit.Text & "'"
    
    If mHayDatos = True Then
        rpt.Formulas(1) = "nrohoja = " & XN(Val(txtInicial.Text) - 1)
        rpt.Formulas(7) = ""
    Else
        rpt.Formulas(1) = ""
        rpt.Formulas(7) = "nrohojasindatos = " & XN(txtInicial.Text)
    End If
    If chkMembrete.Value = 0 Then
        rpt.Formulas(0) = "HOJA = 'PAGINA'"
    End If
   
   Select Case cboDestino.ListIndex
        Case 0 'Pantalla
            rpt.Destination = 0
            rpt.WindowMinButton = 0
            rpt.WindowTitle = "Libro de Iva Ventas"
            rpt.WindowBorderStyle = 2
        Case 1 'Impresora
            rpt.Destination = 1
            rpt.CopiesToPrinter = tNumCopias.Text
            rpt.PrinterStartPage = Val(tPagDesde.Text)
            rpt.PrinterStopPage = Val(tPagHasta.Text)
        Case 2 'Archivo
            CDI.DialogTitle = "Ingrese Nombre de Archivo"
            'CDI.DefaultExt = "DOC"
            'CDI.Filter = "*.DOC"
            'CDI.DefaultExt = "XLS"
            'CDI.Filter = "*.XLS"
            CDI.InitDir = "C:\"
            CDI.Flags = &H800 + &H4 + &H8 + &H2
            CDI.ShowSave
            If Trim(CDI.FileName) = "" Then
                Screen.MousePointer = vbNormal
                BE.Caption = ""
                Exit Sub
            Else
                rpt.Destination = 2
                rpt.PrintFileType = 17
                'rpt.PrintFileType = 10
                rpt.Connect = StrCon
                rpt.PrintFileName = App.Path + CDI.FileName
            End If
    End Select
    rpt.Connect = StrCon
    If txtTipoLibro.Text = "V" Then
        rpt.ReportFileName = RptPath + "librov.rpt"
    Else
        rpt.ReportFileName = RptPath + "libroc.rpt"
    End If
    rpt.Action = 1
    
    If chkApertura.Value = 1 Then
        Apertura_Impositiva
        Select Case cboDestino.ListIndex
             Case 0 'Pantalla
                 rpt.Destination = 0
                 rpt.WindowMinButton = 0
                 rpt.WindowTitle = "Libro de Iva Ventas"
                 rpt.WindowBorderStyle = 2
             Case 1 'Impresora
                 rpt.Destination = 1
                 rpt.CopiesToPrinter = tNumCopias.Text
                 rpt.PrinterStartPage = Val(tPagDesde.Text)
                 rpt.PrinterStopPage = Val(tPagHasta.Text)
             Case 2 'Archivo
                 CDI.DialogTitle = "Ingrese Nombre de Archivo"
                 'CDI.DefaultExt = "DOC"
                 'CDI.Filter = "*.DOC"
                 'CDI.DefaultExt = "XLS"
                 'CDI.Filter = "*.XLS"
                 CDI.InitDir = "C:\"
                 CDI.Flags = &H800 + &H4 + &H8 + &H2
                 CDI.ShowSave
                 If Trim(CDI.FileName) = "" Then
                     Screen.MousePointer = vbNormal
                     BE.Caption = ""
                     Exit Sub
                 Else
                     rpt.Destination = 2
                     rpt.PrintFileType = 17
                     rpt.Connect = StrCon
                     rpt.PrintFileName = App.Path + CDI.FileName
                 End If
        End Select
        rpt.Formulas(0) = ""
        rpt.Formulas(1) = ""
        rpt.Formulas(2) = ""
        rpt.Formulas(3) = ""
        rpt.Formulas(4) = ""
        rpt.Formulas(5) = ""
        rpt.Formulas(6) = ""
        rpt.Formulas(7) = ""
        
        rpt.Connect = StrCon
        'rpt.SelectionFormula = " {LIBRO_IVA.TIPO_LIBRO}=" & XS(txtTipoLibro.Text) & " AND {LIBRO_IVA.LIBRO_ID}=" & txtLibro_IdI.Text
        rpt.ReportFileName = RptPath + "aperturas.rpt"
        rpt.Action = 1
    End If
    
    Screen.MousePointer = vbNormal
    'BE.Caption = MsgBE(0)
    Exit Sub
ErrorTran:
    ErrManager Err.Number
    Screen.MousePointer = vbNormal
End Sub

Private Sub CBImpresora_Click()
  CDImpresora.PrinterDefault = True
  CDImpresora.ShowPrinter
  LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
End Sub

Private Sub cboDestino_Click()
    picSalida(0).Visible = False
    picSalida(1).Visible = False
    picSalida(2).Visible = False
    picSalida(cboDestino.ListIndex).Visible = True
    If cboDestino.ListIndex = 1 Then 'impresora
        Impresora
    Else
        Pantalla
    End If
End Sub


Private Sub cCancelar_Click()
    Unload Frm
    Screen.MousePointer = vbNormal
End Sub



Private Sub cmdAyuda_Click()
    Call WinHelp(Me.hWnd, App.Path & "\help\AYUDA.HLP", cdlHelpContext, 11)
End Sub

Private Sub Form_Activate()
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    
    If txtTipoLibro.Text = "V" Then
        frmImprime.Caption = "Libro Iva Ventas ... Periodo: " & txtMes_LibroI.Text & "-" & txtAnio_LibroI.Text & " ... "
    Else
        frmImprime.Caption = "Libro Iva Compras ... Periodo: " & txtMes_LibroI.Text & "-" & txtAnio_LibroI.Text & " ... "
    End If
    
    cSQL = "SELECT * FROM empresa WHERE emp_id = " & XN(txtEmp_Id.Text)
    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If (rec.BOF And rec.EOF) = 0 Then
        'si encontró el registro muestro los datos
        txtEmpresa.Text = ChkNull(rec!emp_descri)
        txtEmpresaCuit.Text = ChkNull(rec!emp_cuit)
        txtEmpresaDire.Text = ChkNull(rec!emp_direccion)
        frmImprime.Caption = frmImprime.Caption & txtEmpresa.Text
    Else
        txtEmpresa.Text = ""
        txtEmpresaCuit.Text = ""
        txtEmpresaDire.Text = ""
    End If
    rec.Close
End Sub

Private Sub Impresora()
    'fraNumCopias.Visible = True
    'CBImpresora.Visible = True
    'Frame2.Visible = True
    fraNumCopias.Enabled = True
    Frame2.Enabled = True
    tNumCopias.Text = 1
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
    LBImpActual.Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

    If KeyAscii = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j, K As Integer
    'Dim QDF As QueryDef
    Dim FindIt As Integer
    Set Frm = frmImprime
    Tablas = ""
    Campos = ""
    Erase rptEst()
    cboDestino.ListIndex = 0
    sSQL(0) = "SELECT sap_vista, sap_titulo, sap_archivo from salida WHERE sap_nombre='ent'"
    If QRY(sSQL(0), snp) > 0 Then
        rptNombre = Trim(snp!sap_archivo & "")
        rptTitulo = Trim(snp!sap_titulo & "")
        Me.Caption = "Listado de Libro de Iva"
        Vista = Trim(snp!sap_vista & "")
    End If
    snp.Close
    FindIt = 0
'    For K = 0 To Base.QueryDefs.Count - 1
'        Debug.Print UCase(Base.QueryDefs(K).Name)
'        If UCase(Base.QueryDefs(K).Name) = UCase(Vista) Then
'            FindIt = -1
'            Set QDF = Base.QueryDefs(K)
'            For I = 0 To QDF.Fields.Count - 1
'                rptEst(0, I) = I
'                rptEst(1, I) = QDF.Fields(I).Name
'                rptEst(2, I) = "{" & Vista & "." & QDF.Fields(I).Name & "}"
'                rptEst(3, I) = Tipo(QDF.Fields(I).Type)
'                Clave(0).AddItem Trim(rptEst(1, I)), I
'                Clave(1).AddItem Trim(rptEst(1, I)), I
'                Clave(2).AddItem Trim(rptEst(1, I)), I
'            Next
'            For J = 0 To QDF.Fields.Count - 1
'                If J < 3 Then
'                  Clave(J).ListIndex = J
'                  Oper(J).ListIndex = 0
'                Else
'                  Exit For
'                End If
'            Next
'            For I = 2 To QDF.Fields.Count Step -1
'                Clave(I).Enabled = False
'                cbSent(I).Enabled = False
'                TXT(I).Enabled = False
'            Next
'            Exit For
'        End If
'    Next K
'    If FindIt = 0 Then
'        MsgBox "Se produjo un error al querer crear la estructura de la ayuda.", 16, AppName
'        Screen.MousePointer = vbNormal
'        Unload frm
'    End If
    Inicio = False
    LBImpActual.Caption = "Impresora Actual: " & Printer.DeviceName
    Impresora
    cboDestino.ListIndex = 0
    Pantalla
    'QDF.Close
    'Set QDF = Nothing

End Sub


Private Sub Archivo()
    fraNumCopias.Visible = False
    CBImpresora.Visible = False
    tNumCopias.Text = 1
    LBImpActual.Visible = False
    CDI.DialogTitle = "Ingrese Nombre de Archivo"
    CDI.DefaultExt = "DOC"
    CDI.Filter = "*.DOC"
    CDI.InitDir = "C:\TRABAJO"
    CDI.Flags = &H800 + &H4 + &H8 + &H2
    CDI.ShowSave
    If Trim(CDI.FileName) = "" Then
        oPantalla.Value = True
    End If
End Sub

Private Sub Pantalla()
    'fraNumCopias.Visible = False
    fraNumCopias.Enabled = False
    Frame2.Enabled = False
    'CBImpresora.Visible = False
    tNumCopias.Text = 1
    'LBImpActual.Visible = False
    'Frame2.Visible = False
End Sub

Private Sub Form_Resize()
    'If frm.WindowState = 0 Then
    '    frm.Height = 4740
    '    frm.Width = 7260
    'End If
End Sub





Private Sub txtInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
