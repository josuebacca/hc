VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPermisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Permisos"
   ClientHeight    =   5805
   ClientLeft      =   660
   ClientTop       =   420
   ClientWidth     =   6030
   Icon            =   "FrmPermisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4785
      Picture         =   "FrmPermisos.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Salir "
      Top             =   4950
      Width           =   975
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3810
      Picture         =   "FrmPermisos.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo"
      Top             =   4950
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2835
      Picture         =   "FrmPermisos.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar"
      Top             =   4950
      Width           =   975
   End
   Begin VB.CommandButton CmdDeselec 
      Caption         =   "&Deseleccionar todo"
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
      Left            =   2985
      TabIndex        =   4
      Top             =   4515
      Width           =   2790
   End
   Begin VB.CommandButton CmdSelec 
      Caption         =   "&Seleccionar todo"
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
      Left            =   210
      TabIndex        =   3
      Top             =   4515
      Width           =   2715
   End
   Begin VB.ComboBox CboUsuario 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   4740
   End
   Begin MSFlexGridLib.MSFlexGrid GrdMenuses 
      Height          =   3765
      Left            =   225
      TabIndex        =   1
      Top             =   675
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   6641
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "FrmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboUsuario_Click()

    Set rec = New ADODB.Recordset

    Screen.MousePointer = 11
    If Trim(CboUsuario) <> "" Then
        sql = "SELECT * FROM PERMISOS WHERE " & _
        "USU_NOMBRE = '" & Trim(CboUsuario) & "' AND REO_CODIGO = 1"
        rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
        If rec.RecordCount > 0 Then
            rec.MoveFirst
            Do While Not rec.EOF
                For A = 1 To GrdMenuses.Rows - 1
                    GrdMenuses.row = A
                    GrdMenuses.Col = 2
                    If Trim(GrdMenuses) = Trim(rec!PRM_OPMENU) Then
                        GrdMenuses.Col = 1
                        GrdMenuses = "SI"
                    End If
                Next
                rec.MoveNext
            Loop
        Else
            'Limpio_Grilla
        End If
        rec.Close
    End If
    Screen.MousePointer = 1
End Sub

Private Sub CmdDeselec_Click()
    For A = 1 To GrdMenuses.Rows - 1
        GrdMenuses.row = A
        GrdMenuses.Col = 1
        GrdMenuses = ""
    Next
End Sub
Private Sub CmdGrabar_Click()
    If Trim(CboUsuario) = "" Then
        MsgBox "No ha seleccionado el Usuario !", vbExclamation, TIT_MSGBOX
        CboUsuario.SetFocus
        Exit Sub
    End If
   
    Set rec = New ADODB.Recordset
   
    Screen.MousePointer = 11
    DBConn.BeginTrans

    DBConn.Execute "DELETE FROM PERMISOS WHERE " & _
    "USU_NOMBRE = '" & Trim(CboUsuario) & "' AND " & _
    "REO_CODIGO = 1 AND " & _
    "PRM_SISTEMA = '" & Trim(App.Title) & "'"
   
    For A = 1 To GrdMenuses.Rows - 1
        GrdMenuses.row = A
        GrdMenuses.Col = 1
            
        If Trim(GrdMenuses) <> "" Then
        
            GrdMenuses.Col = 2
            
            sql = "SELECT * FROM PERMISOS WHERE " & _
            "PRM_OPMENU = '" & Trim(GrdMenuses) & "' AND " & _
            "PRM_SISTEMA = '" & Trim(App.Title) & "' AND " & _
            "REO_CODIGO = 1 AND " & _
            "USU_NOMBRE = '" & Trim(CboUsuario) & "'"
            rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
            If rec.RecordCount = 0 Then
                sql = "INSERT INTO PERMISOS (REO_CODIGO,PRM_OPMENU,USU_NOMBRE,PRM_SISTEMA) VALUES (" & _
                "1,'" & Trim(GrdMenuses) & "','" & Trim(CboUsuario) & "','" & Trim(App.Title) & "') "
                DBConn.Execute sql
                
             End If
            rec.Close
        End If
    Next
    DBConn.CommitTrans
    
   
    
    Screen.MousePointer = 1
    Exit Sub
    
Exit Sub
    DBConn.RollbackTrans
    'Mensaje 1
End Sub

Private Sub CmdNuevo_Click()
    CmdDeselec_Click
    CboUsuario.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set FrmPermisos = Nothing
End Sub

Private Sub CmdSelec_Click()
    For A = 1 To GrdMenuses.Rows - 1
        GrdMenuses.row = A
        GrdMenuses.Col = 1
        GrdMenuses = "SI"
    Next
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    GrdMenuses.FormatString = "<Menu|^Permiso"
    GrdMenuses.ColWidth(0) = 4000
    GrdMenuses.ColWidth(1) = 1000
    GrdMenuses.ColWidth(2) = 0
    GrdMenuses.Rows = 1
    
    'Cargo los items del menu
    For i = 0 To Menu.Controls.Count - 1
        If TypeName(Menu.Controls(i)) = "Menu" Then
            If Trim(Menu.Controls(i).Caption) <> "-" Then
                GrdMenuses.AddItem Space(5 * Val(Menu.Controls(i).HelpContextID)) & Trim(LIMPIAR(Menu.Controls(i).Caption)) & Chr(9) & Chr(9) & Trim(LIMPIAR(Menu.Controls(i).Name))
            End If
        End If
    Next
    
    'CARGO LOS USUARIOS
    Set rec = New ADODB.Recordset
    sql = "SELECT * FROM USUARIO"
    rec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If rec.RecordCount > 0 Then
        rec.MoveFirst
        Do While Not rec.EOF
            CboUsuario.AddItem Trim(rec!USU_NOMBRE)
            rec.MoveNext
        Loop
    End If
    rec.Close
    
    Screen.MousePointer = 1
    
    CentrarW FrmPermisos
   
End Sub

Private Function LIMPIAR(texto As String) As String

    For A = 1 To Len(texto)
        If Mid(Trim(texto), A, 1) = "&" Then
            texto = Mid(Trim(texto), 1, A - 1) & Mid(Trim(texto), A + 1, Len(texto))
        End If
    Next
    
    LIMPIAR = texto
End Function

Private Sub GrdMenuses_GotFocus()
    GrdMenuses.Col = 0
    GrdMenuses.ColSel = 1
    GrdMenuses.HighLight = flexHighlightAlways
End Sub

Private Sub GrdMenuses_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        GrdMenuses.Col = 1
        If Trim(GrdMenuses) = "" Then
            GrdMenuses = "SI"
        Else
            GrdMenuses = ""
        End If
        GrdMenuses.Col = 0
        GrdMenuses.ColSel = 1
        
    End If
End Sub

Private Sub GrdMenuses_LostFocus()
    GrdMenuses.HighLight = flexHighlightNever
End Sub

