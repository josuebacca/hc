VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CORE - Turnos de Pacientes"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   ForeColor       =   &H00000000&
   Icon            =   "frmTurnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCortar 
      Height          =   375
      Left            =   12120
      Picture         =   "frmTurnos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cortar Turnos"
      Top             =   50
      Width           =   495
   End
   Begin VB.CommandButton cmdCopiar 
      Height          =   375
      Left            =   11640
      Picture         =   "frmTurnos.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Copiar Turnos"
      Top             =   50
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   5520
      ScaleHeight     =   105
      ScaleWidth      =   345
      TabIndex        =   24
      Top             =   450
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3600
      ScaleHeight     =   105
      ScaleWidth      =   345
      TabIndex        =   22
      Top             =   450
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2040
      Picture         =   "frmTurnos.frx":0A1E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Reporte"
      Height          =   735
      Left            =   1080
      Picture         =   "frmTurnos.frx":1A60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   735
      Left            =   120
      Picture         =   "frmTurnos.frx":272A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar Turnos"
      Height          =   735
      Left            =   2040
      Picture         =   "frmTurnos.frx":376C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   735
      Left            =   1080
      Picture         =   "frmTurnos.frx":3AF6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   2895
      Begin VB.ComboBox cboDoctor 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   250
         Width           =   2700
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos del Turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   120
      TabIndex        =   12
      Top             =   3400
      Width           =   2895
      Begin VB.TextBox txtOSocial 
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
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "Descripción"
         Top             =   1680
         Width           =   2715
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         Tag             =   "Descripción"
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox txtBuscarCliDescri 
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Descripción"
         Top             =   620
         Width           =   2715
      End
      Begin VB.TextBox txtBuscaCliente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   2
         Top             =   250
         Width           =   1395
      End
      Begin VB.ComboBox cbohasta 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3420
         Width           =   1260
      End
      Begin VB.TextBox txtMotivo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   4
         Tag             =   "Descripción"
         Top             =   2400
         Width           =   2650
      End
      Begin VB.ComboBox cboDesde 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3420
         Width           =   1260
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Obra Social:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   17
         Top             =   3000
         Width           =   1320
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Motivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paciente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   250
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   2895
      Begin MSComCtl2.MonthView MViewFecha 
         Height          =   2370
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   54460418
         CurrentDate     =   40049
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdGrilla 
      Height          =   8325
      Left            =   3120
      TabIndex        =   8
      ToolTipText     =   "Doble Click para ver la Historia Clinica del Paciente"
      Top             =   525
      Width           =   10000
      _ExtentX        =   17648
      _ExtentY        =   14684
      _Version        =   393216
      Rows            =   25
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   290
      BackColor       =   12648384
      ForeColor       =   49152
      ForeColorFixed  =   -2147483635
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      GridColor       =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
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
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   735
      Left            =   120
      Picture         =   "frmTurnos.frx":4B38
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   975
   End
   Begin Crystal.CrystalReport Rep 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPegar 
      Height          =   375
      Left            =   12600
      Picture         =   "frmTurnos.frx":4EC2
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Pegar Turnos"
      Top             =   50
      Width           =   495
   End
   Begin VB.Label lblAux 
      Caption         =   "Label7"
      Height          =   255
      Left            =   8880
      TabIndex        =   33
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No asistio al Turno"
      Height          =   195
      Left            =   6000
      TabIndex        =   25
      Top             =   450
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Asistio al Turno"
      Height          =   195
      Left            =   4080
      TabIndex        =   23
      Top             =   450
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbldiaTurno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   3600
      TabIndex        =   13
      Top             =   60
      Width           =   945
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Left            =   3120
      Top             =   45
      Width           =   9885
   End
End
Attribute VB_Name = "frmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer
Dim J As Integer
Dim hDesde As Integer
Dim hHasta As Integer
Dim ActivoGrid As Integer ' 1 actio 0 desactivo
Dim sAction As String
Dim dFechaCopy As String
Dim nDoctorCopy As String
Dim sNameDoctorCopy As String



Private Sub cboDesde_LostFocus()
    If cboDesde.ListIndex < cboDesde.ListCount - 1 Then
        cbohasta.ListIndex = cboDesde.ListIndex + 1
    Else
        cbohasta.ListIndex = cboDesde.ListIndex
    End If
End Sub

Private Sub cboDoctor_LostFocus()
    LimpiarGrilla
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
End Sub

Private Sub cbohasta_LostFocus()
    If cboDesde.ListIndex = -1 Then
        If cbohasta.ListIndex > 0 Then
            cboDesde.ListIndex = cbohasta.ListIndex - 1
        Else
            cboDesde.ListIndex = cbohasta.ListIndex
        End If
    End If
End Sub
Private Function ValidarTurno() As Boolean
    If MViewFecha.Value < Date Then
        MsgBox "No puede agregar un turno para ese dia", vbCritical, TIT_MSGBOX
        MViewFecha.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If txtBuscaCliente.Text = "" Then
        MsgBox "No ha ingresado el paciente", vbCritical, TIT_MSGBOX
        txtBuscaCliente.SetFocus
        ValidarTurno = False
        Exit Function
    End If
'    If txtMotivo.Text = "" Then
'        MsgBox "No ha ingresado el Motivo del Turno", vbCritical, TIT_MSGBOX
'        txtMotivo.SetFocus
'        ValidarTurno = False
'        Exit Function
'    End If
    If cboDesde.ListIndex = -1 Then
        MsgBox "No ha ingresado la hora de comienzo del Turno", vbCritical, TIT_MSGBOX
        cboDesde.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    If cbohasta.ListIndex = -1 Then
        MsgBox "No ha ingresado la hora de finalización del Turno", vbCritical, TIT_MSGBOX
        cbohasta.SetFocus
        ValidarTurno = False
        Exit Function
    End If
    
    ValidarTurno = True
End Function
Private Function ImprimirTurno()
    Rep.SelectionFormula = ""
    Rep.Formulas(0) = ""
            
    Rep.SelectionFormula = " {TURNOS.TUR_FECHA}= " & XDQ(MViewFecha.Value)
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.VEN_CODIGO}= " & cboDoctor.ItemData(cboDoctor.ListIndex)
    Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.CLI_CODIGO}= " & XN(txtCodigo.Text)
    'Rep.SelectionFormula = Rep.SelectionFormula & " AND {TURNOS.TUR_HORAD}= #" & TRIM(cboDesde.Text) & "#"
    
    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    'Rep.Connect = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & SERVIDOR & ";"
    Rep.WindowTitle = "Impresion del Turno"
    Rep.ReportFileName = DirReport & "rptTurno.rpt"
    Rep.Action = 1
End Function
Private Sub cmdAgregar_Click()
    Dim nFilaD As Integer
    Dim nFilaH As Integer
    Dim sHoraD As String
    Dim sHoraDAux As String
    'Validar los campos requeridos
    If ValidarTurno = False Then Exit Sub
    If MsgBox("¿Confirma el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'agregar teniendo en cuentas loc combos de horas
    On Error GoTo HayErrorTurno
    
    grdGrilla.HighLight = flexHighlightAlways
    
    nFilaD = cboDesde.ListIndex
    nFilaH = cbohasta.ListIndex
    i = 0
    
    sHoraDAux = cboDesde.Text
    For i = 1 To nFilaH - nFilaD
        DBConn.BeginTrans
        
        sHoraD = cboDesde.Text
        sHoraD = Mid(sHoraD, 1, 1)
        
        If sHoraD = "0" Then
            sHoraD = Mid(cboDesde.Text, 2, 4)
        Else
            sHoraD = Trim(cboDesde.Text)
        End If
        
        'ACA TENGO QUE HACER UN CONTROL POR CLAVES PRIMARIAS
        sql = "SELECT * FROM TURNOS"
        sql = sql & " WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = #" & sHoraD & "#"
        sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
        'sql = sql & " AND CLI_CODIGO = " & XN(txtcodigo.Text)
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        
        If Not rec.EOF = False Then
            sql = "INSERT INTO TURNOS"
            sql = sql & " (TUR_FECHA, TUR_HORAD,TUR_HORAH,"
            sql = sql & " VEN_CODIGO,CLI_CODIGO,TUR_MOTIVO,TUR_ASISTIO,TUR_OSOCIAL,"
            'If User <> 99 Then
                sql = sql & " TUR_USER, "
            'End If
            sql = sql & " TUR_FECALTA)"
            sql = sql & " VALUES ("
            sql = sql & XDQ(MViewFecha.Value) & ",#"
            sql = sql & Left(Trim(grdGrilla.TextMatrix(i + nFilaD, 0)), 5) & "#,#"
            sql = sql & Right(Trim(grdGrilla.TextMatrix(i + nFilaD, 0)), 5) & "#,"
            sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
            sql = sql & XN(txtCodigo) & ","
            sql = sql & XS(txtMotivo) & ","
            sql = sql & 0 & ","
            sql = sql & XS(txtOSocial.Text) & ","
            'If User <> 99 Then
                sql = sql & User & ","
            'End If
            sql = sql & XDQ(Date) & ")"
            
            'ACTUALIZO LA GRILLA
            grdGrilla.row = nFilaD + i
            For J = 1 To grdGrilla.Cols - 1
                grdGrilla.Col = J
                grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR NEGRO
                grdGrilla.CellBackColor = &HC0C0FF      'ROSA
                grdGrilla.CellFontBold = True
            Next
            grdGrilla.TextMatrix(i + nFilaD, 1) = txtBuscarCliDescri.Text
            grdGrilla.TextMatrix(i + nFilaD, 2) = txtTelefono.Text
            grdGrilla.TextMatrix(i + nFilaD, 3) = txtOSocial.Text
            grdGrilla.TextMatrix(i + nFilaD, 4) = txtMotivo.Text
            grdGrilla.TextMatrix(i + nFilaD, 5) = cboDoctor.ItemData(cboDoctor.ListIndex)
            grdGrilla.TextMatrix(i + nFilaD, 6) = txtCodigo.Text
            grdGrilla.TextMatrix(i + nFilaD, 7) = 0
            grdGrilla.TextMatrix(i + nFilaD, 8) = txtBuscaCliente.Text
            
        Else
            'MsgBox "Ya hay un turno para ese horario", vbExclamation, TIT_MSGBOX
            'cboDesde.Text = sHoraDAux
'            rec.Close
'            DBConn.Execute sql
'            DBConn.CommitTrans
'            Exit Sub
            If MsgBox("Ya hay un turno para ese horario ¿Confirma la Modificación del Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then
                rec.Close
                Exit Sub
            End If
            ' aca hago el update
            sql = "UPDATE TURNOS SET "
            sql = sql & " CLI_CODIGO =" & XN(txtCodigo.Text) 'CAMBIAR CUANDO CARGUEMOS DNI
            sql = sql & " ,TUR_MOTIVO =" & XS(txtMotivo.Text)
            sql = sql & " ,TUR_OSOCIAL =" & XS(txtOSocial.Text)
            sql = sql & " ,TUR_FECALTA =" & XDQ(Date)
            If User <> 99 Then
                sql = sql & " ,TUR_USER =" & User
            End If
            sql = sql & " WHERE "
            sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
            sql = sql & " AND TUR_HORAD = #" & cboDesde.Text & "#"
            sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
            
            grdGrilla.TextMatrix(i + nFilaD, 1) = txtBuscarCliDescri.Text
            grdGrilla.TextMatrix(i + nFilaD, 2) = txtTelefono.Text
            grdGrilla.TextMatrix(i + nFilaD, 3) = txtOSocial.Text
            grdGrilla.TextMatrix(i + nFilaD, 4) = txtMotivo.Text
            grdGrilla.TextMatrix(i + nFilaD, 5) = cboDoctor.ItemData(cboDoctor.ListIndex)
            grdGrilla.TextMatrix(i + nFilaD, 6) = txtCodigo.Text
            grdGrilla.TextMatrix(i + nFilaD, 7) = 0
            grdGrilla.TextMatrix(i + nFilaD, 8) = txtBuscaCliente.Text
            'BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
        End If

        
        rec.Close
        DBConn.Execute sql
        DBConn.CommitTrans
        
        cboDesde.ListIndex = cboDesde.ListIndex + 1
    Next
    cboDesde.Text = sHoraDAux
    'If MsgBox("¿Imprime el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    'ImprimirTurno
    
    LimpiarTurno
            
    Exit Sub
    
HayErrorTurno:
    Screen.MousePointer = vbNormal
    If rec.State = 1 Then rec.Close
    If Rec1.State = 1 Then Rec1.Close
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, TIT_MSGBOX
    
    'agregar columnas en la grilla, para guardar el codigo de doctor, paciente
    
End Sub

Private Sub CmdBuscar_Click()
    frmBuscarTurnos.Show vbModal
End Sub
Private Sub LimpiarTurno()
    
    txtBuscaCliente.Text = ""
    txtBuscaCliente.ToolTipText = ""
    txtCodigo.Text = ""
    txtTelefono.Text = ""
    txtOSocial.Text = ""
    txtBuscarCliDescri.Text = ""
    txtMotivo.Text = ""
    cboDesde.ListIndex = -1
    cbohasta.ListIndex = -1
    
    txtBuscaCliente.SetFocus
End Sub

Private Sub cmdCopiar_Click()
    If MsgBox("Esta a punto de  Copiar los " & lbldiaTurno.Caption & " " & Chr(13) & " del Doctor: " & cboDoctor.Text & _
    " ¿Confirma Copiar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    sAction = "COPIAR"
    dFechaCopy = MViewFecha.Value
    nDoctorCopy = cboDoctor.ItemData(cboDoctor.ListIndex)
    sNameDoctorCopy = cboDoctor.Text
End Sub

Private Sub cmdCortar_Click()
    If MsgBox("Esta a punto de Cortar los " & lbldiaTurno.Caption & " " & Chr(13) & " del Doctor: " & cboDoctor.Text & _
    " ¿Confirma Cortar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
    
    sAction = "CORTAR"
    dFechaCopy = MViewFecha.Value
    nDoctorCopy = cboDoctor.ItemData(cboDoctor.ListIndex)
    sNameDoctorCopy = cboDoctor.Text
End Sub

Private Sub CmdNuevo_Click()
    LimpiarTurno
    MViewFecha.Value = Date
    'If User <> 99 Then
    '    Call BuscaCodigoProxItemData(XN(User), cboDoctor)
    'Else
    '    cboDoctor.ListIndex = 0
    'End If
End Sub

Private Sub cmdPegar_Click()
    Dim DIA As Integer
    Dim sDiaTurno As String
    DIA = Weekday(dFechaCopy, vbMonday)
    sDiaTurno = "Turnos del dia " & WeekdayName(DIA, False) & " " & day(dFechaCopy) & " de " & MonthName(Month(dFechaCopy), False) & " de " & Year(dFechaCopy)

    If sAction = "CORTAR" Then
        For i = 1 To grdGrilla.Rows - 1
            If grdGrilla.TextMatrix(i, 1) <> "" Then
                Exit For
            End If
        Next
        If i < grdGrilla.Rows - 1 Then
            If MsgBox("Hay Turnos previamente cargados en este dia que se eliminaran si realiza esta acción." & Chr(13) & _
            " ¿Confirma eliminar estos Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
            
            sql = "DELETE FROM TURNOS WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
            sql = sql & " AND VEN_CODIGO =" & cboDoctor.ItemData(cboDoctor.ListIndex)
            DBConn.Execute sql
            LimpiarGrilla
        End If
        
         If MsgBox("Esta a punto de Pegar los " & sDiaTurno & " " & Chr(13) & "previamente cortados del Doctor: " & sNameDoctorCopy & _
        " " & Chr(13) & "¿Confirma Pegar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
        
        sql = "UPDATE TURNOS SET"
        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & ", VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
        sql = sql & " WHERE TUR_FECHA = " & XDQ(dFechaCopy)
        sql = sql & " AND VEN_CODIGO = " & XN(nDoctorCopy)
        DBConn.Execute sql
    
    Else
        
        If sAction = "COPIAR" Then
            For i = 1 To grdGrilla.Rows - 1
                If grdGrilla.TextMatrix(i, 1) <> "" Then
                    Exit For
                End If
            Next
            If i < grdGrilla.Rows - 1 Then
                If MsgBox("Hay Turnos previamente cargados en este dia que se eliminaran si realiza esta acción." & Chr(13) & _
                " ¿Confirma eliminar estos Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
                
                sql = "DELETE FROM TURNOS WHERE TUR_FECHA = " & XDQ(MViewFecha.Value)
                sql = sql & " AND VEN_CODIGO =" & cboDoctor.ItemData(cboDoctor.ListIndex)
                DBConn.Execute sql
                LimpiarGrilla
            End If
            
            
        
             If MsgBox("Esta a punto de Pegar los " & sDiaTurno & " " & Chr(13) & "previamente copiados del Doctor: " & sNameDoctorCopy & _
            " " & Chr(13) & "¿Confirma Pegar los Turnos?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
                        
            sql = "SELECT * FROM TURNOS WHERE TUR_FECHA = " & XDQ(dFechaCopy)
            sql = sql & "AND VEN_CODIGO = " & XN(nDoctorCopy)
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.EOF = False Then
                Do While rec.EOF = False
                    sql = "INSERT INTO TURNOS"
                    sql = sql & " (TUR_FECHA, TUR_HORAD,TUR_HORAH,"
                    sql = sql & " VEN_CODIGO,CLI_CODIGO,"
                    If Not IsNull(rec!TUR_MOTIVO) Then
                        sql = sql & " TUR_MOTIVO,"
                    End If
                    If Not IsNull(rec!TUR_OSOCIAL) Then
                        sql = sql & " TUR_OSOCIAL,"
                    End If
                    sql = sql & "TUR_ASISTIO)"
                    sql = sql & " VALUES ("
                    sql = sql & XDQ(MViewFecha.Value) & ",#"
                    sql = sql & rec!TUR_HORAD & "#,#"
                    sql = sql & rec!TUR_HORAH & "#,"
                    sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
                    sql = sql & XN(rec!CLI_CODIGO) & ","
                    If Not IsNull(rec!TUR_MOTIVO) Then
                        sql = sql & XS(rec!TUR_MOTIVO) & ","
                    End If
                    If Not IsNull(rec!TUR_OSOCIAL) Then
                        sql = sql & XS(rec!TUR_OSOCIAL) & ","
                    End If
                    sql = sql & 0 & ")"
                    
                    DBConn.Execute sql
                    
                    rec.MoveNext
                Loop
            End If
            rec.Close
            
        End If
    End If
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    sAction = ""
    dFechaCopy = ""
    nDoctorCopy = ""
    sNameDoctorCopy = ""
End Sub

Private Sub cmdQuitar_Click()
    'Controlar que se pueda eliminar el turno
    'Borrar de la Grilla
    'Borrar de la BD
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
        If MsgBox("¿Confirma Eiminar el Turno?", vbQuestion + vbYesNo, TIT_MSGBOX) = vbNo Then Exit Sub
            
        sql = "DELETE FROM TURNOS WHERE"
        sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
        sql = sql & " AND TUR_HORAD = #" & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#"
        sql = sql & " AND VEN_CODIGO = " & cboDoctor.ItemData(cboDoctor.ListIndex)
        sql = sql & " AND CLI_CODIGO = " & grdGrilla.TextMatrix(grdGrilla.RowSel, 6)
        DBConn.Execute sql
        
        'ESTO LO HAGO PARA AUDITAR LO TURNOS BORRADOS
        sql = "INSERT INTO DEL_TURNOS"
        sql = sql & " (TUR_FECHA, TUR_HORAD,"
        sql = sql & " VEN_CODIGO,CLI_CODIGO,"
        sql = sql & " TUR_USER,TUR_FECBAJA)"
        sql = sql & " VALUES ("
        sql = sql & XDQ(MViewFecha.Value) & ",#"
        sql = sql & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5) & "#,"
        sql = sql & cboDoctor.ItemData(cboDoctor.ListIndex) & ","
        sql = sql & grdGrilla.TextMatrix(grdGrilla.RowSel, 6) & ","
        sql = sql & User & ","
        sql = sql & XDQ(Date) & ")"
        
        DBConn.Execute sql
    
        'LIMPIO LA GRILLA
        grdGrilla.TextMatrix(grdGrilla.RowSel, 1) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 2) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 3) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 4) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 6) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 7) = ""
        grdGrilla.TextMatrix(grdGrilla.RowSel, 8) = ""
        LimpiarTurno
        
        grdGrilla.row = grdGrilla.RowSel
        For J = 1 To grdGrilla.Cols - 1
            grdGrilla.Col = J
            grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR BLANCO
            grdGrilla.CellBackColor = &HC0FFC0       'ROSA
            grdGrilla.CellFontBold = True
        Next
    End If
End Sub

Private Sub cmdReport_Click()
    'If txtCodCliente.Text = "" Or GrillaAplicar.Rows = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    'lblEstado.Caption = "Buscando Recibo..."

    sql = "DELETE FROM TMP_TURNOS"
    DBConn.Execute sql
    i = 1
    
    For i = 1 To grdGrilla.Rows - 1
        sql = "INSERT INTO TMP_TURNOS "
        sql = sql & " (TMP_HORA,TMP_FECHA,TMP_DOCTOR,TMP_PACIENTE,TMP_TELEFONO,TMP_OSOCIAL,TMP_MOTIVO)"
        sql = sql & " VALUES ( "
        sql = sql & XS(grdGrilla.TextMatrix(i, 0)) & ","
        sql = sql & XDQ(MViewFecha.Value) & ","
        sql = sql & XS(cboDoctor.Text) & ","
        sql = sql & XS(grdGrilla.TextMatrix(i, 1)) & ","
        sql = sql & XS(grdGrilla.TextMatrix(i, 2)) & ","
        sql = sql & XS(grdGrilla.TextMatrix(i, 3)) & ","
        sql = sql & XS(grdGrilla.TextMatrix(i, 4)) & ")"
        DBConn.Execute sql
    Next
          

    Rep.WindowState = crptMaximized
    Rep.WindowBorderStyle = crptNoBorder
    Rep.Connect = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR

    Rep.WindowTitle = "Listado de Turnos del dia"
    Rep.ReportFileName = DirReport & "rptTurnosDiario.rpt"
    Rep.Action = 1
'    lblEstado.Caption = ""
    Screen.MousePointer = vbNormal
    Rep.SelectionFormula = ""
    
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Seguro que desea Salir", vbQuestion + vbYesNo, TIT_MSGBOX) = vbYes Then
        Set frmTurnos = Nothing
        'Set rec = Nothing
        'Set Rec1 = Nothing
        'Set Rec2 = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        cmdSalir_Click
    End If
End Sub

Private Sub Form_Load()
    Set rec = New ADODB.Recordset
    Set Rec1 = New ADODB.Recordset
    Set Rec2 = New ADODB.Recordset
    
    Call Centrar_pantalla(Me)
    
    MViewFecha.Value = Date
    'MsgBox WeekdayName(7, False)
    'lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(Weekday(Date) - 1, False) & " " & day(Date) & " de " & MonthName(Month(Date), False) & " de " & Year(Date)
    configurodia Date
    configurogrilla
    LlenarComboDoctor
    LlenarComboHoras
    BuscarTurnos Date, cboDoctor.ItemData(cboDoctor.ListIndex)
    ActivoGrid = 1
End Sub
Private Sub LimpiarGrilla()
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.TextMatrix(i, 1) = ""
        grdGrilla.TextMatrix(i, 2) = ""
        grdGrilla.TextMatrix(i, 3) = ""
        grdGrilla.TextMatrix(i, 4) = ""
        grdGrilla.row = i
        For J = 1 To grdGrilla.Cols - 1
            grdGrilla.Col = J
            grdGrilla.CellForeColor = &H80000008          'FUENTE COLOR BLANCO
            grdGrilla.CellBackColor = &HC0FFC0       'ROSA
            grdGrilla.CellFontBold = True
        Next
    Next
End Sub
Private Sub BuscarTurnos(day As Date, Doc As Integer)
    Dim sColor As String
    sql = "SELECT T.*,V.VEN_NOMBRE,C.CLI_RAZSOC,C.CLI_NRODOC,C.CLI_TELEFONO"
    sql = sql & " FROM TURNOS T, VENDEDOR V, CLIENTE C"
    sql = sql & " WHERE T.CLI_CODIGO = C.CLI_CODIGO"
    sql = sql & " AND T.VEN_CODIGO = V.VEN_CODIGO"
    sql = sql & " AND T.TUR_FECHA = " & XDQ(day)
    sql = sql & " AND T.VEN_CODIGO = " & Doc
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        Do While rec.EOF = False
            For i = 1 To grdGrilla.Rows - 1
                If Format(rec!TUR_HORAD, "hh:mm") = Left(Trim(grdGrilla.TextMatrix(i, 0)), 5) Then
                    grdGrilla.TextMatrix(i, 1) = rec!CLI_RAZSOC
                    grdGrilla.TextMatrix(i, 2) = ChkNull(rec!CLI_TELEFONO)
                    'SI ESTA VACIA BUSCAR LA OBRA SOCIAL
                    If Not IsNull(rec!TUR_OSOCIAL) Then
                        grdGrilla.TextMatrix(i, 3) = ChkNull(rec!TUR_OSOCIAL)
                    Else
                        grdGrilla.TextMatrix(i, 3) = BuscarOSocial(rec!CLI_CODIGO)
                    End If
                    grdGrilla.TextMatrix(i, 4) = ChkNull(rec!TUR_MOTIVO)
                    grdGrilla.TextMatrix(i, 5) = rec!VEN_CODIGO
                    grdGrilla.TextMatrix(i, 6) = rec!CLI_CODIGO
                    grdGrilla.TextMatrix(i, 7) = rec!TUR_ASISTIO
                    If IsNull(rec!CLI_NRODOC) Then
                        grdGrilla.TextMatrix(i, 8) = ChkNull(rec!CLI_CODIGO) 'SI NO TENGO EL DNI CARGADO, HAY QUE PONER EL CODIGO
                    Else
                        grdGrilla.TextMatrix(i, 8) = ChkNull(rec!CLI_NRODOC) '
                    End If
                    Select Case rec!TUR_ASISTIO
                    Case 0
                        sColor = &H80000008
                    Case 1
                        sColor = &HC000&
                    Case 2
                        sColor = &HFF&
                    End Select
                    grdGrilla.row = i
                    For J = 1 To grdGrilla.Cols - 1
                        grdGrilla.Col = J
                        grdGrilla.CellForeColor = sColor          'FUENTE COLOR NEGRO
                        grdGrilla.CellBackColor = &HC0C0FF      'ROSA
                        grdGrilla.CellFontBold = True
                    Next
                End If
            Next
            rec.MoveNext
        Loop
    End If
    rec.Close
End Sub

Private Sub LlenarComboDoctor()
    sql = "SELECT * FROM VENDEDOR"
    sql = sql & " WHERE PR_CODIGO > 1"
    sql = sql & " ORDER BY VEN_CODIGO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        'cboFactura1.AddItem "(Todas)"
        Do While rec.EOF = False
            cboDoctor.AddItem rec!VEN_NOMBRE
            cboDoctor.ItemData(cboDoctor.NewIndex) = rec!VEN_CODIGO
            rec.MoveNext
        Loop
        If User <> 99 Then
            Call BuscaCodigoProxItemData(XN(User), cboDoctor)
        Else
            cboDoctor.ListIndex = 0
        End If
    End If
    rec.Close
End Sub
Private Sub LlenarComboHoras()
    Dim cItems As Integer
    rec.Open "SELECT HS_DESDE,HS_HASTA FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        hDesde = Hour(rec!HS_DESDE)
        hHasta = Hour(rec!HS_HASTA)
    End If
    rec.Close
    cItems = (hHasta - hDesde) * 2 + 1
    i = 0
    For J = hDesde To hHasta
        cboDesde.AddItem Format(J, "00") & ":00"
        cboDesde.ItemData(cboDesde.NewIndex) = i
        cboDesde.AddItem Format(J, "00") & ":30"
        cboDesde.ItemData(cboDesde.NewIndex) = i + 0.5
        
        cbohasta.AddItem Format(J, "00") & ":00"
        cbohasta.ItemData(cbohasta.NewIndex) = i
        cbohasta.AddItem Format(J, "00") & ":30"
        cbohasta.ItemData(cbohasta.NewIndex) = i + 0.5
        
        i = i + 1
    Next
    cboDesde.ListIndex = -1
    cbohasta.ListIndex = -1
    
End Sub
Private Function configurogrilla()

    grdGrilla.FormatString = "^Horas|<Paciente|<Telefono|<Obra Social|<Motivo|>Doctor|>Cod Pac|>Asistio|DNI"
    grdGrilla.ColWidth(0) = 1400 'HORAS
    grdGrilla.ColWidth(1) = 2700 'PACIENTE
    grdGrilla.ColWidth(2) = 1500 'TELEFONO
    grdGrilla.ColWidth(3) = 2300 'O SOCIAL
    grdGrilla.ColWidth(4) = 2000 'MOTIVO
    grdGrilla.ColWidth(5) = 0 'DOCTOR
    grdGrilla.ColWidth(6) = 0 'Codigo Paciente
    grdGrilla.ColWidth(7) = 0 'Asistio
    grdGrilla.ColWidth(8) = 0 'DNI
    
    grdGrilla.Cols = 9
    grdGrilla.BorderStyle = flexBorderNone
    grdGrilla.row = 0
    For i = 0 To grdGrilla.Cols - 1
        grdGrilla.Col = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    ' Busco los horarios en parametros
    rec.Open "SELECT HS_DESDE,HS_HASTA FROM PARAMETROS", DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        hDesde = Hour(rec!HS_DESDE)
        hHasta = Hour(rec!HS_HASTA)
    End If
    rec.Close
    grdGrilla.Rows = (hHasta - hDesde) * 2 + 1
    
    For i = 1 To grdGrilla.Rows - 1
        grdGrilla.Col = 0
        grdGrilla.row = i
        grdGrilla.CellForeColor = &HFFFFFF 'FUENTE COLOR BLANCO
        grdGrilla.CellBackColor = &H808080    'GRIS OSCURO
        grdGrilla.CellFontBold = True
    Next
    
    J = hDesde
    For i = 1 To grdGrilla.Rows - 2
        grdGrilla.TextMatrix(i, 0) = Format(J, "00") & ":00 a " & Format(J, "00") & ":30 "
        grdGrilla.TextMatrix(i + 1, 0) = Format(J, "00") & ":30 a " & Format(J + 1, "00") & ":00 "
        i = i + 1
        J = J + 1
    Next
'   GRDGrilla.TextMatrix(I, 0) = Hour(I)
    
End Function

Private Sub grdGrilla_Click()
    If ActivoGrid = 1 Then
        If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
            txtBuscaCliente.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 8)
            'txtBuscaCliente_LostFocus
            txtCodigo.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 6)
            txtBuscarCliDescri.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
            txtTelefono.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 2)
            txtOSocial.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 3)
            txtMotivo.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 4)
            BuscaDescriProx Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cboDesde
            BuscaDescriProx Right(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cbohasta
        Else
            If txtBuscaCliente.Text <> "" Then
                MViewFecha.Value = Date
                txtBuscaCliente.Text = ""
                txtCodigo.Text = ""
                txtBuscarCliDescri.Text = ""
                txtTelefono.Text = ""
                txtOSocial.Text = ""
                txtMotivo.Text = ""
                cboDesde.ListIndex = -1
                cbohasta.ListIndex = -1
            End If
        End If
    End If
End Sub

Private Sub GRDGrilla_DblClick()
    ' tengo que marcar en la BD si asistio o no
    Dim nAsiste As Integer '0:Pen 1:SI y  2:NO
    If grdGrilla.TextMatrix(grdGrilla.RowSel, 1) <> "" Then
            'vMode = 9
            gPaciente = grdGrilla.TextMatrix(grdGrilla.RowSel, 6)
            ABMClientes.Show vbModal
            gPaciente = 0
'        txtBuscaCliente.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 8)
'            txtBuscaCliente_LostFocus
'            txtBuscarCliDescri.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 1)
'            txtTelefono.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 2)
'            txtOSocial.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 3)
'            txtMotivo.Text = grdGrilla.TextMatrix(grdGrilla.RowSel, 4)
'            BuscaDescriProx Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cboDesde
'            BuscaDescriProx Right(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 5), cbohasta
'
'        grdGrilla.row = grdGrilla.RowSel
'        For J = 1 To grdGrilla.Cols - 1
'            grdGrilla.Col = J
'            If grdGrilla.CellForeColor = &H80000008 Then
'                grdGrilla.CellForeColor = &HC000& ' Verde
'                nAsiste = 1
'            Else
'                If grdGrilla.CellForeColor = &HC000& Then
'                    grdGrilla.CellForeColor = &HFF& 'Rojo
'                    nAsiste = 2
'                Else
'                    grdGrilla.CellForeColor = &H80000008 ' Negro
'                    nAsiste = 0
'                End If
'            End If
'            FUENTE COLOR BLANCO
'            grdGrilla.CellBackColor = &HC0FFC0       'ROSA
'            grdGrilla.CellFontBold = True
'        Next
'    Actualizo la grilla
'    grdGrilla.TextMatrix(grdGrilla.RowSel, 5) = nAsiste
'    Actualizo BD
'    sql = "UPDATE TURNOS SET "
'    sql = sql & " TUR_ASISTIO =" & nAsiste
'    sql = sql & " WHERE "
'    sql = sql & " TUR_FECHA = " & XDQ(MViewFecha.Value)
'    sql = sql & " AND TUR_HORAD = #" & Left(Trim(grdGrilla.TextMatrix(grdGrilla.RowSel, 0)), 7) & "#"
'    sql = sql & " AND VEN_CODIGO = " & XN(grdGrilla.TextMatrix(grdGrilla.RowSel, 5))
'    DBConn.Execute sql
'    Else
''        If txtBuscaCliente.Text <> "" Then
''            MViewFecha.Value = Date
''            txtBuscaCliente.Text = ""
''            txtcodigo.Text = ""
''            txtBuscarCliDescri.Text = ""
''            txtMotivo.Text = ""
''            cboDesde.ListIndex = -1
''            cbohasta.ListIndex = -1
''        End If
    End If
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdQuitar_Click
    End If
End Sub

Private Sub MViewFecha_DateClick(ByVal DateClicked As Date)
    'lbldiaTurno.Caption = "Turnos del dia " & MViewFecha.Value
    'lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(Weekday(MViewFecha.Value) - 1, False) & " " & day(MViewFecha.Value) & " de " & MonthName(Month(MViewFecha.Value), False) & " de " & Year(MViewFecha.Value)
    configurodia MViewFecha.Value
    LimpiarGrilla
    BuscarTurnos MViewFecha.Value, cboDoctor.ItemData(cboDoctor.ListIndex)
End Sub
Private Sub configurodia(Fecha As Date)
    Dim DIA As Integer
    DIA = Weekday(Fecha, vbMonday)
    lbldiaTurno.Caption = "Turnos del dia " & WeekdayName(DIA, False) & " " & day(Fecha) & " de " & MonthName(Month(Fecha), False) & " de " & Year(Fecha)
End Sub
Private Function BuscarOSocial(CodCli As Integer) As String
Set Rec1 = New ADODB.Recordset
    sql = "SELECT O.OS_NOMBRE FROM OBRA_SOCIAL O, CLIENTE C"
    sql = sql & " WHERE C.OS_NUMERO = O.OS_NUMERO"
    sql = sql & " AND C.CLI_CODIGO = " & CodCli
    Rec1.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If Rec1.EOF = False Then
        BuscarOSocial = Rec1!OS_NOMBRE
    Else
        BuscarOSocial = ""
    End If
    Rec1.Close
End Function
Private Sub txtBuscaCliente_Change()
    If txtBuscaCliente.Text = "" Then
        txtBuscarCliDescri.Text = ""
        txtCodigo.Text = ""
        txtTelefono.Text = ""
        txtOSocial.Text = ""
    End If
    If Len(Trim(txtBuscaCliente.Text)) < 7 Then
        txtBuscaCliente.ToolTipText = "Numero de Paciente"
    Else
        txtBuscaCliente.ToolTipText = "DNI"
    End If
End Sub

Private Sub txtBuscaCliente_GotFocus()
    SelecTexto txtBuscaCliente
End Sub

Private Sub txtBuscaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
        ActivoGrid = 0
    End If
End Sub

Private Sub txtBuscaCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = CarNumeroEntero(KeyAscii)
End Sub
Private Sub txtBuscaCliente_LostFocus()
    If txtBuscaCliente.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            If Len(Trim(txtBuscaCliente.Text)) < 7 Then
                sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
            Else
                sql = sql & " CLI_NRODOC=" & XN(txtBuscaCliente)
            End If
             'sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
'        Else
'            sql = sql & " CLI_CODIGO LIKE '" & Trim(txtcodigo) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            'txtBuscaCliente.Text = rec!CLI_NRODOC
            txtBuscarCliDescri.Text = rec!CLI_RAZSOC
            txtCodigo.Text = rec!CLI_CODIGO
            txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
            txtOSocial.Text = BuscarOSocial(rec!CLI_CODIGO)
            'txtMotivo.SetFocus
            ActivoGrid = 1
        Else
            MsgBox "El Paciente no existe", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
        End If
        rec.Close
    End If
End Sub

Private Sub txtBuscarCliDescri_Change()
    If txtBuscarCliDescri.Text = "" Then
        txtBuscaCliente.Text = ""
        txtCodigo.Text = ""
        txtTelefono.Text = ""
        txtOSocial.Text = ""
    End If
End Sub

Private Sub txtBuscarCliDescri_GotFocus()
    SelecTexto txtBuscarCliDescri
End Sub

Private Sub txtBuscarCliDescri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        BuscarClientes "txtBuscaCliente", "CODIGO"
        ActivoGrid = 0
    End If
End Sub

Private Sub txtBuscarCliDescri_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
End Sub

Private Sub txtBuscarCliDescri_LostFocus()
    If txtBuscaCliente.Text = "" And txtBuscarCliDescri.Text <> "" Then
        Set rec = New ADODB.Recordset
        sql = "SELECT CLI_CODIGO, CLI_RAZSOC,CLI_NRODOC,CLI_TELEFONO"
        sql = sql & " FROM CLIENTE"
        sql = sql & " WHERE "
        If txtBuscaCliente.Text <> "" Then
            If Len(Trim(txtBuscaCliente.Text)) < 7 Then
                sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
            Else
                sql = sql & " CLI_NRODOC=" & XN(txtBuscaCliente)
            End If
            'sql = sql & " CLI_CODIGO=" & XN(txtBuscaCliente)
        Else
            sql = sql & " CLI_RAZSOC LIKE '" & Trim(txtBuscarCliDescri) & "%'"
        End If
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            If rec.RecordCount > 1 Then
                BuscarClientes "txtBuscaCliente", "CADENA", Trim(txtBuscarCliDescri.Text)
                If rec.State = 1 Then rec.Close
                txtBuscarCliDescri.SetFocus
            Else
                'txtBuscaCliente.Text = rec!CLI_DNI
                If Len(Trim(txtBuscaCliente.Text)) < 7 Then
                    txtBuscaCliente.Text = rec!CLI_CODIGO
                Else
                    txtBuscaCliente.Text = rec!CLI_NRODOC
                End If
                'txtBuscaCliente.Text = rec!CLI_NRODOC
                txtBuscarCliDescri.Text = rec!CLI_RAZSOC
                txtCodigo.Text = rec!CLI_CODIGO
                txtTelefono.Text = ChkNull(rec!CLI_TELEFONO)
            End If
            ActivoGrid = 0
        Else
            MsgBox "No se encontro el Paciente", vbExclamation, TIT_MSGBOX
            txtBuscaCliente.SetFocus
        End If
        If rec.State = 1 Then rec.Close
    End If
End Sub
Public Sub BuscarClientes(Txt As String, mQuien As String, Optional mCadena As String)
    Dim cSQL As String
    Dim hSQL As String
    Dim B As CBusqueda
    Dim i, posicion As Integer
    Dim cadena As String
    
    Set B = New CBusqueda
    With B
        cSQL = "SELECT CLI_RAZSOC, CLI_CODIGO,CLI_NRODOC"
        cSQL = cSQL & " FROM CLIENTE C"
        If mQuien = "CADENA" Then
            cSQL = cSQL & " WHERE CLI_RAZSOC LIKE '" & Trim(mCadena) & "%'"
        End If
        
        hSQL = "Nombre, Código, DNI"
        .sql = cSQL
        .Headers = hSQL
        .Field = "CLI_RAZSOC"
        campo1 = .Field
        .Field = "CLI_CODIGO"
        campo2 = .Field
        .Field = "CLI_NRODOC"
        campo3 = .Field
        
        .OrderBy = "CLI_RAZSOC"
        camponumerico = False
        .Titulo = "Busqueda de Clientes :"
        .MaxRecords = 1
        .Show

        ' utilizar la coleccion de datos devueltos
        If .ResultFields.Count > 0 Then
            If Txt = "txtcodCli" Then
                'txtcodCli.Text = .ResultFields(2)
                'txtCodCli_LostFocus
            Else
                If .ResultFields(3) = "" Then
                    txtBuscaCliente.Text = .ResultFields(2)
                Else
                    txtBuscaCliente.Text = .ResultFields(3)
                End If
                txtBuscaCliente_LostFocus
            End If
        End If
    End With
    
    Set B = Nothing
End Sub

Private Sub txtMotivo_GotFocus()
    SelecTexto txtMotivo
End Sub
