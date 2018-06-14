VERSION 5.00
Begin VB.Form FrmInicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Control de Acceso"
   ClientHeight    =   2655
   ClientLeft      =   3870
   ClientTop       =   2415
   ClientWidth     =   4440
   Icon            =   "FrmInicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2784.63
   ScaleMode       =   0  'User
   ScaleWidth      =   2197.8
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   4140
      Begin VB.TextBox TxtClave 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1770
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2115
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   0
         Top             =   555
         Width           =   2115
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   315
         Picture         =   "FrmInicio.frx":27A2
         Top             =   765
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ingrese su nombre de Usuario y Clave"
         Height          =   315
         Index           =   2
         Left            =   795
         TabIndex        =   8
         Top             =   255
         Width           =   2925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   1125
         TabIndex        =   6
         Top             =   1140
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   5
         Top             =   615
         Width           =   585
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      DisabledPicture =   "FrmInicio.frx":2AAC
      Height          =   750
      Left            =   3105
      Picture         =   "FrmInicio.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmInicio.frx":30C0
      Height          =   750
      Left            =   1890
      Picture         =   "FrmInicio.frx":33CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   2055
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CUANTAS_VECES As Integer
Dim rec As ADODB.Recordset
Dim sql As String

Public Sub Conexion()
    Dim sPathBase As String

    'sPathBase = "C:\Program Files\Microsoft Visual Studio\VB98\BIBLIO.MDB"
       
       
    Screen.MousePointer = vbHourglass
    CONECCION = False

    On Error GoTo ErrorIni
    LeoIni
    
    On Error GoTo ErrorTrans
    'ME CONECTO !
    Set DBConn = New ADODB.Connection
    'DBConn.ConnectionString = "ODBC;DATABASE=;UID=" & mNomUser & ";PWD=" & mPassword & ";DSN=" & DSN_DEF
    'DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SISESTILO"
    'DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    'DBConn.ConnectionTimeout = 0       'Default msado10.hlp => 15
    'DBConn.CommandTimeout = 0          'Default msado10.hlp => 30
    'DBConn.Open DBConn.ConnectionString, TxtUsuario, TxtClave
    'DBConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & SERVIDOR & ";"
    DBConn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & SERVIDOR
    DBConn.Open
'    ME CONECTO !
'    Set DBConn = New ADODB.Connection
'    DBConn.ConnectionString = "driver={SQL Server};server=" & SERVIDOR & ";DATABASE=" & BASEDATO
'    DBConn.ConnectionTimeout = 0       'Default msado10.hlp => 15
'    DBConn.CommandTimeout = 0          'Default msado10.hlp => 30
'    DBConn.Open DBConn.ConnectionString, TxtUsuario, TxtClave
        
    If DBConn.State = adStateOpen Then CONECCION = True
        
    PERMISOS mNomUser
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrorTrans:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo establecer la conección a la Base de Datos." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
        Exit Sub
ErrorIni:
        Screen.MousePointer = vbNormal
        MsgBox "No se pudo leer el archivo de configuración del sistema." & Chr(13) & Err.Description, vbExclamation, TIT_MSGBOX
End Sub
Public Sub PERMISOS(USUARIO As String)

    Dim sql As String
    Dim r As ADODB.Recordset
    Dim i As Integer

    Set r = New ADODB.Recordset

    On Error Resume Next

    If Trim(USUARIO) = "SA" Or Trim(USUARIO) = "A" Then
        'Este usuario tiene derecho a todo
        For i = 0 To Menu.Controls.Count - 1
            If TypeName(Menu.Controls(i)) = "Menu" Then
               Menu.Controls(i).Enabled = True
            End If
        Next
    Else
        For i = 0 To Menu.Controls.Count - 1
            If TypeName(Menu.Controls(i)) = "Menu" Then
               Menu.Controls(i).Enabled = False
            End If
        Next
    
        On Error GoTo 0
    
        sql = "SELECT * FROM PERMISOS WHERE " & _
        "USU_NOMBRE = '" & Trim(USUARIO) & "' AND " & _
        "PRM_SISTEMA = '" & Trim(App.Title) & "'"
        r.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If r.RecordCount > 0 Then
            r.MoveFirst
            Do While Not r.EOF
                For i = 0 To Menu.Controls.Count - 1
                    If TypeName(Menu.Controls(i)) = "Menu" Then
                        If UCase(Trim(Menu.Controls(i).Name)) = UCase(Trim(r!PRM_OPMENU)) Then
                            Menu.Controls(i).Enabled = True
                        End If
                    End If
                Next
                r.MoveNext
            Loop
        End If
        r.Close
    End If
End Sub


Private Sub cmdAceptar_Click()
    
    Set rec = New ADODB.Recordset
    mNomUser = Trim(TxtUsuario)
    
    Conexion
    
    If Not CONECCION Then
        If Err.Description <> "" Then
            MsgBox Err.Description
        End If
            
        CUANTAS_VECES = CUANTAS_VECES + 1
        If CUANTAS_VECES = 4 Then
            End
        End If
        TxtUsuario.SetFocus
        Exit Sub
    End If

    sql = "SELECT * FROM USUARIO WHERE " & _
          "USU_NOMBRE = '" & Trim(TxtUsuario) & "' AND " & _
           "USU_CLAVE = '" & Trim(TxtClave) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 1 Then
        sql = "La contraseña de usuario NO ES CORRECTA !" & Chr(13) & Chr(13)
        If CUANTAS_VECES = 3 Then
            sql = sql & "El sistema se cerrará."
        Else
            sql = sql & "Por favor intentelo nuevamente."
        End If
        MsgBox sql, vbCritical, "Error:"
        If CUANTAS_VECES = 3 Then
            'si ya pifió 3 veces salgo del Sistema
            cmdSalir_Click
        Else
            TxtClave.SelStart = 0
            TxtClave.SelLength = Len(TxtClave)
            TxtUsuario.SetFocus
            CUANTAS_VECES = CUANTAS_VECES + 1
        End If
    Else
        Label1(1).FontBold = True
        Label1(1).Caption = " Conectando ... "
        Label1(1).Refresh
        'muestro un figureti de coneccion
        mNomUser = Trim(TxtUsuario)
        mPassword = Trim(TxtClave)
        
        'BUSCO SUCURSALES---
            BuscoNroSucursal
        '-----------------
        Unload Me
        Set FrmInicio = Nothing
    End If
End Sub
Private Sub CmdAceptar_GotFocus()
    cmdAceptar.FontBold = True
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub CmdSalir_GotFocus()
    CmdSalir.FontBold = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then cmdSalir_Click
    If KeyAscii = 13 Then
        MySendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    CUANTAS_VECES = 1
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        cmdAceptar_Click
    End If
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    'If KeyAscii = vbKeyReturn Then TxtClave.SetFocus
End Sub
