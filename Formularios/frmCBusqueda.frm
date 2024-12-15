VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar..."
   ClientHeight    =   4095
   ClientLeft      =   1185
   ClientTop       =   2505
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5400
      TabIndex        =   4
      Top             =   3660
      Width           =   1300
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "&Seleccionar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4020
      TabIndex        =   3
      Top             =   3660
      Width           =   1300
   End
   Begin VB.CommandButton cmdEje 
      Caption         =   "Iniciar &Busqueda"
      Height          =   360
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1635
   End
   Begin ComctlLib.ListView lvwLista 
      Height          =   2550
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   4498
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblTit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   180
   End
   Begin ComctlLib.ImageList ImgLstLista 
      Left            =   60
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   17
      ImageHeight     =   15
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCBusqueda.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros encontrados:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   780
      Width           =   1680
   End
End
Attribute VB_Name = "frmCBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancelar_Click()
    If lvwLista.ListItems.Count > 0 Then
        lvwLista.SelectedItem.Selected = False
    End If
    Unload Me
End Sub

Private Sub cmdEje_Click()
    Dim Rec As ADODB.Recordset
    Dim itmX As ListItem
    Dim ArrWidthColumn() As Integer
    Dim f As Field
    Dim Pos As Integer
    Dim CantCampos As Byte, TodosLosCampos As Byte
    Dim tempSQL As String
    
    If Trim(txtBuscar.Text) = "" Then
        'Exit Sub
        sField = campo1
        camponumerico = False
    Else
        If txtBuscar.Text = "" Then
            sField = campo2
        Else
            If lblTit.Caption = "Busqueda de Clientes :" Or lblTit.Caption = "Busqueda de Proveedores :" Then
                If Val(Mid(txtBuscar.Text, 1, 1)) > 0 Then
                    sField = campo2
                    camponumerico = False
                Else
                    sField = campo1
                    camponumerico = False
                End If
            Else
                sField = campo1
                camponumerico = False
            End If
        End If
    End If
    
   If camponumerico = True Then
        tempSQL = sSQL
        If InStr(UCase(sSQL), "WHERE") <> 0 Then ' ya tiene WHERE en la consulta
            tempSQL = tempSQL & " and " & sField & "=  " & Trim(txtBuscar.Text) & " order by " & sOrderBy
        Else
            tempSQL = tempSQL & " WHERE " & sField & " = " & Trim(txtBuscar.Text) & " order by " & sOrderBy
        End If
   Else
        tempSQL = sSQL
        If InStr(UCase(sSQL), "WHERE") <> 0 Then ' ya tiene WHERE en la consulta
            tempSQL = tempSQL & " and " & sField & " like '" & Trim(txtBuscar.Text) & "%' order by " & sOrderBy
        Else
            tempSQL = tempSQL & " WHERE " & sField & " like '" & Trim(txtBuscar.Text) & "%' order by " & sOrderBy
        End If
   End If
   
    CargarListView Me, lvwLista, tempSQL, , hSQL, ImgLstLista, iMaxRecords
    'aca toy
    If lvwLista.ListItems.Count > 0 Then
        lvwLista.SetFocus
        Set lvwLista.SelectedItem = lvwLista.ListItems(1)
        cmdSeleccionar.Enabled = True
        cmdSeleccionar.Default = True
    Else
        cmdSeleccionar.Enabled = False
        cmdEje.Default = True
    End If
End Sub


Private Sub cmdSeleccionar_Click()

Dim itmX As ListItem
Dim CantCol As Byte
Dim I As Integer

CantCol = frmCBusqueda.lvwLista.ColumnHeaders.Count - 1

Set itmX = frmCBusqueda.lvwLista.SelectedItem
ColSel.Add itmX.Text, frmCBusqueda.lvwLista.ColumnHeaders(1)
For I = 1 To CantCol
    ColSel.Add itmX.SubItems(I), frmCBusqueda.lvwLista.ColumnHeaders(I + 1)
Next I

Unload Me

End Sub

Public Sub CargarListView(ByRef pForm As Form, ByRef lvwSource As ListView, ByVal sql As String, Optional IDReg As String, Optional pHeaderSQL As String, Optional pImgList As ImageList, Optional pMaxRec As Long)

    'Carga un control de tipo ListView a partir de un string SQL dado como parametro
    'Parámetro: lvwSource es el control que será cargado
    '                       SQL es el string SQL a utilizar
    '                       [IDReg] nombre del campo identificador de registro opcional
    '                       [pHeaderSQL] lista headers para columnas
    '                       [pImgLst] lista de iconos
    '                       [pMaxRec] cantidad máxima de registros a cargar
    'Retorna: el control listview con datos y un formato de columnas

    Dim vRec  As ADODB.Recordset
    Dim itmX As ListItem
    Dim ArrWidthColumn() As Integer
    Dim f As Field
    Dim Pos As Integer
    Dim PosForm As Integer
    
    Dim CantCampos As Integer
    Dim DesdeCampo As Integer
    Dim HastaCampo As Integer
    
    Dim IndiceCampo As Integer
    Dim IndiceColumna As Integer
    
    Dim DatoMostrar As Variant
    Dim AlinColumna As Integer
    Dim IconoWidth As Integer
    Dim CantRegistros As Long
    
    Dim PosComa As Integer
    Dim TextHeader As String
    
    'control de parametros opcionales
'    If IsMissing(pMaxRec) Then
'        CantRegistros = 0
'    Else
'        CantRegistros = pMaxRec
'    End If
    
    If Not IsMissing(IDReg) And Trim(IDReg) <> "" Then
        'preparo el stringSQL agregando el campo IDReg al final de los campos seleccionados
        PosForm = InStr(1, sql, "FROM")
        sql = Left(sql, PosForm - 1) & ", " & IDReg & " " & Right(sql, Len(sql) - PosForm + 1)
    End If
        
    
    'llenar el recordset
    Set vRec = New ADODB.Recordset
    vRec.Open sql, DBConn, adOpenStatic, adLockOptimistic
    If vRec.EOF = False Then
        
        CantCampos = vRec.Fields.Count ' cantidad de campos en la tabla
        
        If Not IsMissing(IDReg) And Trim(IDReg) <> "" Then
            HastaCampo = CantCampos - 2
        Else
            HastaCampo = CantCampos - 1
        End If
        
        ' array que contiene los anchos de las columnas
        ReDim ArrWidthColumn(0 To HastaCampo)
        
        
        lvwSource.ListItems.Clear
        lvwSource.ColumnHeaders.Clear
        
        If Not IsMissing(pImgList) Then
            lvwSource.Icons = pImgList
            lvwSource.SmallIcons = pImgList
        End If
                
        'preparo headers de columnas dependiendo de los parametros ingresados
        If Not IsMissing(pHeaderSQL) And Trim(pHeaderSQL) <> "" Then
            
            'recorro el string de headers para columnas
            PosComa = 1
            For IndiceCampo = 0 To HastaCampo
            
                'selecciono el texto para la columna
                If InStr(PosComa, pHeaderSQL, ",") > 0 Then
                    TextHeader = Trim(Mid(pHeaderSQL, PosComa, InStr(PosComa, pHeaderSQL, ",") - PosComa))
                    PosComa = InStr(PosComa, pHeaderSQL, ",") + 1
                Else
                    TextHeader = Trim(Right(pHeaderSQL, Len(pHeaderSQL) - PosComa + 1))
                End If
            
                'preparo la alineación de columna
                Select Case vRec.Fields(IndiceCampo).Type
                    Case dbSqlNumeric, dbSqlInt, dbSqlSmallint
                        AlinColumna = lvwColumnRight
                    Case dbSqlDate, dbSqlChar, dbSqlVarchar
                        AlinColumna = lvwColumnLeft
                    Case Else
                        AlinColumna = lvwColumnLeft
                End Select
            
                If IndiceCampo = 0 Then
                    lvwSource.ColumnHeaders.Add , , TextHeader, lvwSource.Width / 5
                Else
                    lvwSource.ColumnHeaders.Add , , TextHeader, lvwSource.Width / 5, AlinColumna
                End If
                
                'guardo el ancho del texto del header
                ArrWidthColumn(IndiceCampo) = pForm.TextWidth(TextHeader)
                
            Next IndiceCampo
            
        Else
                
            ' creo una columna por campo encontrado en el recordset
            IndiceCampo = -1
            For Each f In vRec.Fields
                IndiceCampo = IndiceCampo + 1
                
                'preparo la alineación de columna
                Select Case f.Type
                    Case dbSqlNumeric, dbSqlInt, dbSqlSmallint
                        AlinColumna = lvwColumnRight
                    Case dbSqlDate, dbSqlChar, dbSqlVarchar
                        AlinColumna = lvwColumnLeft
                    Case Else
                        AlinColumna = lvwColumnLeft
                End Select
                
                If IndiceCampo <= HastaCampo Then
                    If IndiceCampo = 0 Then
                        lvwSource.ColumnHeaders.Add , , f.Name, lvwSource.Width / 5
                    Else
                        lvwSource.ColumnHeaders.Add , , f.Name, lvwSource.Width / 5, AlinColumna
                    End If
                    
                    'guardo el ancho del texto del header
                    ArrWidthColumn(IndiceCampo) = pForm.TextWidth(TextHeader)
                    
                End If
            Next f
            
        End If
                
        
        ' cargo los items y subitems de la lista
        While Not vRec.EOF
        
            For IndiceCampo = 0 To HastaCampo ' por cada campo
            
                'preparo formato del campo a mostrar
                If IsNull(vRec.Fields(IndiceCampo)) Then
                    DatoMostrar = ""
                Else
                    Select Case vRec.Fields(IndiceCampo).Type
                        Case dbSqlNumeric
                            DatoMostrar = Format(vRec.Fields(IndiceCampo), "0.00")
                        Case dbSqlDate
                            DatoMostrar = CDate(Format(vRec.Fields(IndiceCampo), "DD/MM/YYYY"))
                        Case dbSqlInt, dbSqlSmallint
                            DatoMostrar = vRec.Fields(IndiceCampo)
                        Case dbSqlChar, dbSqlVarchar
                            DatoMostrar = vRec.Fields(IndiceCampo)
                        Case Else
                            DatoMostrar = vRec.Fields(IndiceCampo)
                    End Select
                End If
            
                'muestro el campo formateado
                If IndiceCampo = 0 Then
                    'si es el primero, lo agrego como ITEM
                    If IsMissing(IDReg) Or Trim(IDReg) = "" Then
                        Set itmX = lvwSource.ListItems.Add(, , DatoMostrar, 1)
                    Else
                        Set itmX = lvwSource.ListItems.Add(, "'" & vRec.Fields(CantCampos - 1) & "'", DatoMostrar, 1)
                    End If
                    itmX.Icon = 1
                    itmX.SmallIcon = 1
                Else
                    'si no es el primero, lo agrego como SUBITEM
                    itmX.SubItems(IndiceCampo) = DatoMostrar
                End If
                
                ' calculo el ancho y mantengo guardado el ancho mayor para asignarlo luego como el de la columna
                If IndiceCampo = 0 Then
                    IconoWidth = 250
                Else
                    IconoWidth = 0
                End If
                
                If pForm.TextWidth(DatoMostrar) + IconoWidth > ArrWidthColumn(IndiceCampo) Then
                    ArrWidthColumn(IndiceCampo) = pForm.TextWidth(DatoMostrar) + IconoWidth
                End If
                
            Next IndiceCampo
            vRec.MoveNext
        Wend
            
        ' ajusto el tamaño de las columnas
        For IndiceColumna = 0 To HastaCampo
            lvwSource.ColumnHeaders(IndiceColumna + 1).Width = ArrWidthColumn(IndiceColumna)
        Next IndiceColumna
        
        vRec.Close
    End If
    
End Sub

Private Sub Form_Load()
    'variables de uso general
    Dim Rec As ADODB.Recordset
    Dim itmX As ListItem
    Dim ArrWidthColumn() As Integer
    Dim f As Field
    Dim Pos As Integer
    Dim CantCampos As Byte, TodosLosCampos As Byte
    
    lblTit.Caption = sTitulo ' establecer titulo del campo de busqueda
    CargarListView Me, lvwLista, sSQL, , hSQL, ImgLstLista, iMaxRecords
    
    If lvwLista.ListItems.Count > 0 Then
        Set lvwLista.SelectedItem = lvwLista.ListItems(1)
        cmdSeleccionar.Enabled = True
        cmdSeleccionar.Default = True
    Else
        cmdSeleccionar.Enabled = False
        cmdEje.Default = True
    End If
    
End Sub

Private Sub Form_Resize()

'If WindowState <> 1 Then ' si no esta minimizada
'    lvwLista.Width = frmCBusqueda.Width - 120
'    lvwLista.Height = frmCBusqueda.Height - 1900
'
'    cmdSeleccionar.Top = Me.Height - 800
'    cmdCancelar.Top = Me.Height - 800
'    cmdSeleccionar.Left = Me.Width - 2900
'    cmdCancelar.Left = Me.Width - 1500
'
'End If

End Sub

Private Sub lvwLista_Click()

    If lvwLista.ListItems.Count > 0 Then
        If lvwLista.SelectedItem.Selected = True Then
            cmdSeleccionar.Enabled = True
        Else
            cmdSeleccionar.Enabled = False
        End If
    Else
        cmdSeleccionar.Enabled = False
    End If

End Sub

Private Sub lvwLista_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    
    If lvwLista.SortOrder = lvwAscending Then
        lvwLista.SortOrder = lvwDescending
    Else
        lvwLista.SortOrder = lvwAscending
    End If
    
    lvwLista.SortKey = ColumnHeader.Index - 1
    ' Establece Sorted a True para ordenar la lista.
    lvwLista.Sorted = True

End Sub

Private Sub lvwLista_DblClick()
    If lvwLista.ListItems.Count > 0 Then 'And lvwLista.SelectedItem.Selected Then
        Call cmdSeleccionar_Click
    End If
End Sub

Private Sub lvwLista_GotFocus()
    cmdSeleccionar.Default = True
End Sub

Private Sub lvwLista_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If lvwLista.ListItems.Count > 0 Then 'And lvwLista.SelectedItem.Selected Then
            Call cmdSeleccionar_Click
        End If
    End If

End Sub

Private Sub txtBuscar_GotFocus()
    cmdEje.Default = True
End Sub

