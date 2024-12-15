VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Usuarios"
   ClientHeight    =   4755
   ClientLeft      =   2055
   ClientTop       =   1890
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   495
      Top             =   4905
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab Tabtb 
      Height          =   3840
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Usuarios"
      TabPicture(0)   =   "FrmUsuarios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CmdClave"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdNuevo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LstUsuarios"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdEliminar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tabborrar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Datos"
      TabPicture(1)   =   "FrmUsuarios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraClave"
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab tabborrar 
         Height          =   1950
         Left            =   240
         TabIndex        =   19
         Top             =   540
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3440
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame Frame1 
            Caption         =   " Borrar Usuario "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1680
            Left            =   135
            TabIndex        =   20
            Top             =   90
            Width           =   4155
            Begin VB.TextBox TxtBorrar 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Left            =   1710
               PasswordChar    =   "*"
               TabIndex        =   23
               Top             =   1035
               Width           =   1185
            End
            Begin VB.Label Label1 
               Caption         =   "Para poder eliminar un Usuario debe ingresar previamente la contraseña del mismo."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Index           =   6
               Left            =   135
               TabIndex        =   22
               Top             =   360
               Width           =   3900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "contraseña:"
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
               Index           =   5
               Left            =   495
               TabIndex        =   21
               Top             =   1080
               Width           =   1020
            End
         End
      End
      Begin VB.Frame FraClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3120
         Left            =   -74775
         TabIndex        =   13
         Top             =   450
         Width           =   6405
         Begin VB.TextBox txtDescrip 
            Height          =   330
            Left            =   1350
            MaxLength       =   20
            TabIndex        =   4
            Top             =   405
            Width           =   2895
         End
         Begin VB.TextBox TxtClaveConfirmar 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   2745
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   2295
            Width           =   1500
         End
         Begin VB.TextBox TxtClaveNueva 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   2745
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   1890
            Width           =   1500
         End
         Begin VB.TextBox TxtClave 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   2745
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1260
            Width           =   1500
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   5310
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   18
            Top             =   1755
            Width           =   480
         End
         Begin VB.TextBox TxtNombre 
            Height          =   330
            Left            =   2745
            MaxLength       =   20
            TabIndex        =   5
            Top             =   795
            Width           =   1500
         End
         Begin VB.CommandButton CmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   420
            Left            =   4860
            TabIndex        =   9
            Top             =   315
            Width           =   1320
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   420
            Left            =   4860
            TabIndex        =   10
            Top             =   765
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   0
            Left            =   405
            TabIndex        =   24
            Top             =   450
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Index           =   1
            Left            =   2100
            TabIndex        =   17
            Top             =   855
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ingrese la contraseña actual:"
            Height          =   195
            Index           =   2
            Left            =   630
            TabIndex        =   16
            Top             =   1305
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Confirme la contraseña:"
            Height          =   195
            Index           =   3
            Left            =   1035
            TabIndex        =   15
            Top             =   2340
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ingrese la nueva contraseña:"
            Height          =   195
            Index           =   4
            Left            =   630
            TabIndex        =   14
            Top             =   1935
            Width           =   2070
         End
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Borrar Usuario"
         Enabled         =   0   'False
         Height          =   420
         Left            =   4950
         TabIndex        =   2
         Top             =   1080
         Width           =   1680
      End
      Begin VB.ListBox LstUsuarios 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         Left            =   225
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   4515
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo Usuario"
         Height          =   420
         Left            =   4950
         TabIndex        =   1
         Top             =   540
         Width           =   1680
      End
      Begin VB.CommandButton CmdClave 
         Caption         =   "&Cambiar Contraseña"
         Enabled         =   0   'False
         Height          =   465
         Left            =   4950
         TabIndex        =   3
         Top             =   1620
         Width           =   1680
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   5925
      Picture         =   "FrmUsuarios.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3975
      Width           =   960
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As ADODB.Recordset
Dim sql As String

Private Sub cmdAceptar_Click()
    Dim DBConnAux As ADODB.Connection
    'Controlo que la clave sea correcta
    If TxtClave.Enabled Then
        If Not CONTROLAR_CLAVE Then Exit Sub
    End If
    'Controlo que la confirmacion de la contraseña sea correcta
    'Si la confirmacion y la nueva no son iguales no dejo grabar
    If Trim(TxtClaveNueva) <> Trim(TxtClaveConfirmar) Then
        Beep
        MsgBox "Las contraseñas ingresadas no coinciden !  " & _
        "La contraseña NO se ha actualizado.", vbCritical, "Error:"
        TxtClaveConfirmar.SetFocus
        Exit Sub
    End If
    
    On Error GoTo CLAVOSE
    Screen.MousePointer = vbHourglass
    Me.Refresh
    
    If txtNombre.Enabled Then
        'si el txtnombre esta habilitado es porque estoy cargando un nuevo usuario
        
        If Trim(txtNombre) = "" Or Trim(txtDescrip) = "" Then
            MsgBox "No ha ingresado el nombre del usuario !", vbExclamation, "Mensaje:"
            If txtNombre.Enabled Then txtNombre.SetFocus
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
       
        sql = "SELECT USU_NOMBRE FROM USUARIO WHERE USU_NOMBRE = '" & Trim(txtNombre) & "'"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.EOF = False Then
            MsgBox "El usuario ya existe !", vbCritical, "Error:"
            txtNombre.SetFocus
            Exit Sub
        End If
        
        'SI NOMBRE ESTA HABILITADO ESTOY CARGANDO UN USUARIO NUEVO
        DBConn.Execute "INSERT INTO USUARIO (USU_DESCRI, USU_NOMBRE, USU_CLAVE) VALUES " & _
        "('" & Trim(txtDescrip) & "','" & Trim(txtNombre) & "','" & Trim(TxtClaveNueva) & "')"
        MsgBox "El ususario ha sido ingresado !", vbInformation, "Mensaje:"
        CARGAR_USUARIOS
        CmdCancelar_Click
    Else
        DBConn.Execute "UPDATE USUARIO SET " & _
        "USU_DESCRI = '" & Trim(txtDescrip) & "', " & _
        "USU_CLAVE = '" & Trim(TxtClaveNueva) & "' WHERE " & _
        "USU_NOMBRE = '" & Trim(txtNombre) & "'"
        
        'sql = "sp_password " & XS(TxtClave) & ", " & XS(TxtClaveNueva)
        'DBConn.Execute sql
        
        MsgBox "La contraseña se ha actualizado correctamente !", vbInformation, "Mensaje:"
    End If
    
    Screen.MousePointer = vbNormal
    CmdCancelar_Click
    Exit Sub

CLAVOSE:
    If rec.State = 1 Then rec.Close
    Screen.MousePointer = vbNormal
    Mensaje 3
End Sub

Private Sub CmdCancelar_Click()
    TxtClave.Text = ""
    TxtClaveNueva.Text = ""
    TxtClaveConfirmar.Text = ""
    txtDescrip.Text = ""
    txtNombre.Text = ""
    Tabtb.TabEnabled(0) = True
    Tabtb.TabEnabled(1) = False
    Tabtb.Tab = 0
    LstUsuarios.SetFocus
End Sub

Private Sub CmdClave_Click()
    txtNombre.Enabled = False
    txtNombre = LstUsuarios.Text
    TxtClave.Enabled = True
    TxtClave.BackColor = vbWhite
    TxtClave.SetFocus
    sql = "SELECT USU_DESCRI FROM USUARIO WHERE USU_NOMBRE = '" & Trim(txtNombre) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
        txtDescrip.Text = ChkNull(rec!USU_DESCRI)
    End If
    rec.Close
    Tabtb.Tab = 1
End Sub

Private Sub cmdEliminar_Click()
    tabborrar.Top = 500
    tabborrar.Left = 1500
    tabborrar.Visible = True
    TxtBorrar.Text = ""
    CmdEliminar.Enabled = False
    CmdClave.Enabled = False
'    Menu.SB.SimpleText = "<ENTER> Aceptar - <ESC> Cancelar"
'    Menu.SB.Refresh
    TxtBorrar.SetFocus
End Sub

Private Sub CmdNuevo_Click()
    txtNombre.Enabled = True
    txtDescrip.Enabled = True
    txtDescrip.SetFocus
    TxtClave.Enabled = False
    TxtClave.BackColor = vbButtonFace
    Tabtb.Tab = 1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmUsuarios = Nothing
End Sub

Private Sub Form_Load()
'    Menu.SB.SimpleText = ""
    CARGAR_USUARIOS
    Tabtb.TabEnabled(1) = False
    Tabtb.Tab = 0
    Screen.MousePointer = vbNormal
End Sub

Private Sub LstUsuarios_Click()
    CmdEliminar.Enabled = True
    CmdClave.Enabled = True
End Sub

Private Sub LstUsuarios_GotFocus()
'    Menu.SB.SimpleText = " <Enter> Cambiar Contraseña - <Insert> Agregar nuevo Usuario - <Delete> Borrar Usuario"
'    Menu.SB.Refresh
End Sub

Private Sub LstUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdEliminar_Click
    ElseIf KeyCode = vbKeyInsert Then
        CmdNuevo_Click
    End If
End Sub

Private Sub LstUsuarios_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CmdClave_Click
End Sub

Private Sub TABTB_Click(PreviousTab As Integer)
    If Tabtb.Tab = 1 Then
        Tabtb.TabEnabled(0) = False
        Tabtb.TabEnabled(1) = True
        CmdEliminar.Enabled = False
        CmdClave.Enabled = False
        LstUsuarios.ListIndex = -1
    End If
End Sub

Private Sub TxtBorrar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        BORRAR_USUARIO
    ElseIf KeyAscii = vbKeyEscape Then
'        Menu.SB.SimpleText = ""
        tabborrar.Visible = False
        LstUsuarios.SetFocus
    End If
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TxtClaveNueva.SetFocus
End Sub

Private Sub TXTCLAVEConfirmar_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub

Private Sub TxtCLaveNueva_GotFocus()
    If Not txtNombre.Enabled Then
        'si no esta habilitado el nombre quiere decir que estoy cambiando
        'la contrasenia de un usuario existente y le pido la contrasenia
        'para asegurarme que no la cambie cualquiera !
        If Not CONTROLAR_CLAVE Then
            TxtClave.SelStart = 0
            TxtClave.SelLength = Len(TxtClave)
            TxtClave.SetFocus
        End If
    End If
End Sub

Private Sub TXTCLAVENUEVA_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TxtClaveConfirmar.SetFocus
End Sub

Private Sub BORRAR_USUARIO()
    sql = "SELECT USU_CLAVE FROM USUARIO WHERE USU_NOMBRE = '" & Trim(LstUsuarios.Text) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount = 1 Then
        If Trim(rec.Fields(0)) = Trim(TxtBorrar) Then
            'Si la contrasena coincide borro el usuario
            DBConn.Execute "DELETE FROM USUARIO WHERE USU_NOMBRE = '" & Trim(LstUsuarios.Text) & "'"
            MsgBox "El usuario ha sido eliminado.", vbInformation, "Mensaje:"
            LstUsuarios.RemoveItem (LstUsuarios.ListIndex)
'            Menu.SB.SimpleText = ""
            CmdEliminar.Enabled = False
            CmdClave.Enabled = False
            tabborrar.Visible = False
            If LstUsuarios.ListCount > 0 Then
                LstUsuarios.ListIndex = 0
            Else
                LstUsuarios.ListIndex = -1
            End If
            LstUsuarios.SetFocus
        Else
            Beep
            MsgBox "La contraseña no es correcta !  " & _
            "El usuario NO ha sido eliminado.", vbCritical, "Error:"
            TxtBorrar.SetFocus
        End If
    End If
    rec.Close
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = CarTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If TxtClave.Enabled Then
            TxtClave.SetFocus
        Else
            TxtClaveNueva.SetFocus
        End If
    End If
End Sub

Private Sub CARGAR_USUARIOS()
    Set rec = New ADODB.Recordset
    LstUsuarios.Clear
    rec.Open "SELECT * FROM USUARIO", DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
       Do While Not rec.EOF
          LstUsuarios.AddItem rec.Fields(0)
          rec.MoveNext
       Loop
    End If
    rec.Close
End Sub

Private Function CONTROLAR_CLAVE() As Boolean
    CONTROLAR_CLAVE = True
    sql = "select * from USUARIO WHERE " & _
    "USU_NOMBRE = '" & Trim(txtNombre) & "' AND " & _
    "USU_CLAVE = '" & Trim(TxtClave) & "'"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount <> 1 Then
        Beep
        MsgBox "La contraseña no es correcta !  " & _
        "No puede modificarla.", vbCritical, "Error:"
        CONTROLAR_CLAVE = False
    End If
    rec.Close
End Function
