VERSION 5.00
Begin VB.Form Consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta"
   ClientHeight    =   3975
   ClientLeft      =   2280
   ClientTop       =   1905
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3975
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6375
      TabIndex        =   11
      Top             =   3495
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   225
      TabIndex        =   6
      Top             =   60
      Width           =   7620
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6075
         TabIndex        =   1
         Top             =   645
         Width           =   1455
      End
      Begin VB.CheckBox chkExacta 
         Caption         =   "Búsqueda exacta"
         Height          =   285
         Left            =   5940
         TabIndex        =   10
         Top             =   225
         Width           =   1545
      End
      Begin VB.ComboBox cboBusqueda 
         Height          =   315
         ItemData        =   "FrmConsulta.frx":0000
         Left            =   1965
         List            =   "FrmConsulta.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   255
         Width           =   1845
      End
      Begin VB.TextBox txtCond_Busqueda 
         Height          =   285
         Left            =   1965
         TabIndex        =   0
         Top             =   690
         Width           =   3960
      End
      Begin VB.Label Label2 
         Caption         =   "Búsqueda por:"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label lblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Búsqueda:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   765
         Width           =   1740
      End
   End
   Begin VB.ListBox lstDes_Cons 
      Height          =   1425
      Left            =   225
      TabIndex        =   2
      Top             =   1305
      Width           =   7620
   End
   Begin VB.TextBox txtDes_Cons 
      Height          =   285
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3090
      Width           =   7590
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4665
      TabIndex        =   3
      Top             =   3495
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Seleccionado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   5
      Top             =   2835
      Width           =   1935
   End
End
Attribute VB_Name = "Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim snpConsulta As ADODB.Recordset

Dim cSQL As String
Dim Ventana As Form
Dim Tabla As String

Dim CampoCod As String
Dim CampoDes As String

Dim PaiCodigo As String
Dim ProCodigo As String
Dim LocCodigo As String

Dim CrtlCodigo As Control
Dim CrtlDescrip As Control

Dim Desc_Busq As String
Dim Nom_Ventana As String

Dim ValPaiCodigo As Integer
Dim ValProCodigo As Integer
Dim ValLocCodigo As Integer


Public Function Parametros(auxVentana As Form, _
                           auxTabla As String, _
                           auxCampoCod As String, auxCampoDes As String, _
                           Optional auxPaiCodigo As String, Optional auxProCodigo As String, Optional auxLocCodigo As String, _
                           Optional auxCrtlCodigo As Control, Optional auxCrtlDescrip As Control, _
                           Optional auxNom_Ventana As Variant, _
                           Optional auxDesc_Busq As Variant, _
                           Optional Codigo1 As Control, Optional Codigo2 As Control, Optional Codigo3 As Control)
    
    Set Ventana = auxVentana 'Objeto ventana que llama a la ayuda
    
    Tabla = auxTabla         'Nombre de la tabla
    CampoCod = auxCampoCod   'Nombre del campo que tiene el codigo en la tabla
    CampoDes = auxCampoDes   'Nombre del campo que tiene la descripcion en la tabla
    PaiCodigo = auxPaiCodigo 'Nombre del campo que tiene el Código del País en la Tabla
    ProCodigo = auxProCodigo 'Nombre del campo que tiene el Código de la Provincia en la Tabla
    LocCodigo = auxLocCodigo 'Nombre del campo que tiene el Código de la Localidad en la Tabla
    
    'asigno los valores de los Códigos que necesito en el WHERE
    'If Codigo1 <> "" Then ValPaiCodigo = Codigo1
    'If Codigo2 <> "" Then ValProCodigo = Codigo2
    'If Codigo3 <> "" Then ValLocCodigo = Codigo3
    
    If IsMissing(auxDesc_Busq) Then
        Desc_Busq = "Condición de Búsqueda:"
    Else
        Desc_Busq = auxDesc_Busq 'Caption del label para el ingreso de datos
    End If
    
    Set CrtlCodigo = auxCrtlCodigo 'Objeto Control del form ventana al que se asigna el codigo
    Set CrtlDescrip = auxCrtlDescrip 'Objeto Control del form ventana al que se asigna la descripcion
    
    If IsMissing(auxNom_Ventana) Then
        Nom_Ventana = ""
    Else
        Nom_Ventana = auxNom_Ventana 'Caption de la ventana de ayuda
    End If
    
End Function

Private Sub cboBusqueda_Change()
 If cboBusqueda.Text = "Descripción" Then
     chkExacta.Value = 0 = True
    chkExacta.Visible = True
 ElseIf cboBusqueda.Text = "Código" Then
    chkExacta.Visible = False
 End If
End Sub

Private Sub cboBusqueda_Click()
 If cboBusqueda.Text = "Descripción" Then
    chkExacta.Value = 0
    chkExacta.Visible = True
 ElseIf cboBusqueda.Text = "Código" Then
    chkExacta.Visible = False
 End If
End Sub

Private Sub cmdAceptar_Click()
    Call ValidarIngreso
End Sub

Private Sub CmdBuscar_Click()

    Set rec = New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    If IsNull(txtCond_Busqueda.Text) Or IsEmpty(txtCond_Busqueda.Text) Or Len(txtCond_Busqueda.Text) = 0 Then
       txtCond_Busqueda.Text = ""
    Else
       txtCond_Busqueda.Text = UCase(txtCond_Busqueda)
    End If
    
    If cboBusqueda.Text = "Descripción" Then
    Select Case Nom_Ventana
    Case "Consulta de Clientes - Proveedores"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Items"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Tipos de Ingresos"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Tipos de Egresos"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de País"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Provincias por País"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & PaiCodigo & " = " & Trim(ValPaiCodigo)
         cSQL = cSQL & " And " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
     Case "Consulta de Localidades por Provincia"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & PaiCodigo & " = " & Trim(ValPaiCodigo)
         cSQL = cSQL & " And " & ProCodigo & " = " & Trim(ValProCodigo)
         cSQL = cSQL & " And " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Barrios por Localidad"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & PaiCodigo & " = " & Trim(ValPaiCodigo)
         cSQL = cSQL & " And " & ProCodigo & " = " & Trim(ValProCodigo)
         cSQL = cSQL & " And " & LocCodigo & " = " & Trim(ValLocCodigo)
         cSQL = cSQL & " And " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Tipos de Comitentes"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Tipos de Errores"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Tipos de Obleas"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Cuentas"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & " CONTA.DBO." & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    Case "Consulta de Rubros"
         cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
         cSQL = cSQL & Tabla & " WHERE " & CampoDes
        If chkExacta = 0 Then
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "%'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        Else
            cSQL = cSQL & " LIKE '" & txtCond_Busqueda
            cSQL = cSQL & "'"
            cSQL = cSQL & "ORDER BY " & CampoDes
        End If
    End Select
    ElseIf cboBusqueda.Text = "Código" Then
      If txtCond_Busqueda.Text = "" Then
         MsgBox "Ingrese el Código por el cual buscar.", 16, TIT_MSGBOX
         Screen.MousePointer = 1
         txtCond_Busqueda.SetFocus
         Exit Sub
      Else
        Select Case Nom_Ventana
        Case "Consulta de Clientes - Proveedores"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Provincias por País"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & PaiCodigo & " = " & Trim(ValPaiCodigo)
             cSQL = cSQL & " And " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
         Case "Consulta de Localidades por Provincia"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & PaiCodigo & " = " & Trim(ValPaiCodigo)
             cSQL = cSQL & " And " & ProCodigo & " = " & Trim(ValProCodigo)
             cSQL = cSQL & " And " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Barrios por Localidad"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & PaiCodigo & " = " & Trim(ValPaiCodigo)
             cSQL = cSQL & " And " & ProCodigo & " = " & Trim(ValProCodigo)
             cSQL = cSQL & " And " & LocCodigo & " = " & Trim(ValLocCodigo)
             cSQL = cSQL & " And " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Tipos de Comitentes"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Tipos de Errores"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Tipos de Obleas"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Cuentas"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & " CONTA.DBO." & Tabla & " WHERE " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        Case "Consulta de Rubros"
             cSQL = "SELECT " & CampoCod & ", " & CampoDes & " FROM "
             cSQL = cSQL & Tabla & " WHERE " & CampoCod
             cSQL = cSQL & " = " & txtCond_Busqueda
        End Select
      End If
   End If

    rec.Open cSQL, DBConn, adOpenStatic, adLockOptimistic
    If rec.EOF = False Then
      lstDes_Cons.Clear
      Do Until rec.EOF
          lstDes_Cons.AddItem rec(CampoCod) & "    " & rec(CampoDes)
          rec.MoveNext
      Loop
      If lstDes_Cons.ListCount > 0 Then
         lstDes_Cons.ListIndex = 0
         lstDes_Cons.Text = lstDes_Cons.Text
      Else
         txtDes_Cons = ""
      End If
      lstDes_Cons.SetFocus
    Else
      lstDes_Cons.Clear
      txtDes_Cons.Text = ""
      MsgBox "La Busqueda no tuvo éxito.", 16, TIT_MSGBOX
      txtCond_Busqueda.SetFocus
    End If
    rec.Close
    Screen.MousePointer = 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
    Set Consulta = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 And Screen.ActiveControl.Name <> "txtCond_Busqueda" Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call Centrar_pantalla(Me)
    cboBusqueda.Text = "Descripción"
    Me.Caption = Nom_Ventana
    lblDescripcion.Caption = Desc_Busq
    Screen.MousePointer = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub lstDes_Cons_Click()
    txtDes_Cons.Text = lstDes_Cons.Text
End Sub

Private Sub lstDes_Cons_DblClick()
    Call ValidarIngreso
End Sub

Private Sub lstDes_Cons_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call ValidarIngreso
    End If
End Sub

Private Sub txtCond_Busqueda_KeyPress(KeyAscii As Integer)
    If cboBusqueda.Text = "Descripción" Then
        KeyAscii = CarTexto(KeyAscii)
    Else
        KeyAscii = CarNumeroEntero(KeyAscii)
    End If
    
    If KeyAscii = vbKeyReturn Then
        CmdBuscar_Click
        lstDes_Cons.SetFocus
    End If
End Sub

Private Function ValidarIngreso()
Dim A As Integer
    If lstDes_Cons.ListIndex <> -1 Then
        'CrtlCodigo = lstDes_Cons.ItemData(lstDes_Cons.ListIndex)
        For A = 1 To Len(Trim(lstDes_Cons.Text))
            If Mid(Trim(lstDes_Cons.Text), A, 1) = " " Then
                CrtlCodigo = Trim(Mid(lstDes_Cons.Text, 1, A))
                CrtlDescrip = Trim(Mid(lstDes_Cons.Text, A + 1, 100))
                Exit For
            End If
        Next
    End If
    Unload Me
    Set Consulta = Nothing
End Function
Private Sub txtdes_cons_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub
