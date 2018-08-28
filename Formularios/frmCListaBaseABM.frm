VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCListaBaseABM 
   Caption         =   "Cargando Lista ..."
   ClientHeight    =   4725
   ClientLeft      =   885
   ClientTop       =   1560
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCListaBaseABM.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   7980
   Begin ComctlLib.Toolbar tbarHerramientas 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      ImageList       =   "ImgLstBarra"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo registro"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Editar"
            Object.ToolTipText     =   "Editar registro"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Ver"
            Object.ToolTipText     =   "Detalle de registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refrescar"
            Object.ToolTipText     =   "Refrescar la lista de registros"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualiza"
            Object.ToolTipText     =   "Filtro Búsqueda"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir listado"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar ventana"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   7530
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar sBarEstado 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   4455
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   13555
            MinWidth        =   8819
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ComctlLib.ListView lstvLista 
      Height          =   4155
      Left            =   0
      TabIndex        =   1
      Top             =   390
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImgLstLista"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImgLstBarra 
      Left            =   7380
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":06F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":080A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstLista 
      Left            =   7365
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   17
      ImageHeight     =   15
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABM.frx":095C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCListaBaseABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'defino variables globales a nivel de form para cofigurar el uso de la ventana
Option Explicit

'Dim frmDatos As Form
Private vStringSQL As String
Private vFieldID As String
Private vMaxRecords As Long
Private vHeaderSQL As String
'Dim WidthWindow As Long
'Dim HeightWindow As Long
'Dim CenterWindow As Boolean
'Dim TextToolbar As Boolean

Private auxDll As CListaBaseABM


Public Function SetWindow(ByVal pStringSQL As String, pFieldID As String, pHeaderSQL As String, ByVal pDll As CListaBaseABM, Optional ByVal pMaxRec As Long)

    vStringSQL = pStringSQL
    vFieldID = pFieldID
    
    Set auxDll = pDll
    Set pDll = Nothing
    
    If IsMissing(pMaxRec) Then
        vMaxRecords = 0
    Else
        vMaxRecords = pMaxRec
    End If
    
    vHeaderSQL = pHeaderSQL
End Function



Private Sub Form_Activate()

    Set auxDllActiva = auxDll
    If lstvLista.ListItems.Count > 0 Then
        lstvLista.ListItems(1).Selected = True
        lstvLista.SetFocus
    End If
    
    
End Sub

Private Sub Form_Load()
                
    CargarListView Me, lstvLista, vStringSQL, vFieldID, vHeaderSQL, ImgLstLista
    'rptListado.Formulas(0) = "EMPRESA = '" & vDesEmpresa & "'"
    rptListado.Formulas(0) = ""
    If mOrigen = False And mAdentro = False Then
        Menu.stbPrincipal.Panels(1).Text = ""
    Else
    'If mOrigen = False And mAdentro = True Then
        Menu.stbPrincipal.Panels(1).Text = "<Insert> Nuevo  <ENTER> Edita  <DEL> Borra "
    'Else
    '    Menu.stbPrincipal.Panels(1).Text = ""
    End If
End Sub


Private Sub Form_Resize()
    'VER EL ERROR ACA
    If Me.WindowState <> 1 Then
        Me.Refresh
        lstvLista.Top = tbarHerramientas.Top + tbarHerramientas.Height
        lstvLista.Width = Me.Width - 120
        lstvLista.Height = Me.Height - (sBarEstado.Height + tbarHerramientas.Height + 400)
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    mAdentro = False
    Menu.stbPrincipal.Panels(1).Text = ""
    'Unload frmEligePeriodo
    Unload auxDll.FormDatos
    
End Sub

Private Sub lstvLista_Click()

    Menu.mnuContextABM(0).Enabled = False
    Menu.mnuContextABM(1).Enabled = False
    Menu.mnuContextABM(2).Enabled = False
    Menu.mnuContextABM(4).Enabled = False
    Menu.mnuContextABM(6).Enabled = False
    Menu.mnuContextABM(7).Enabled = False
    Menu.mnuContextABM(9).Enabled = False
    
End Sub

Private Sub lstvLista_ColumnClick(ByVal ColumnHeader As ColumnHeader)

    lstvLista.SortKey = ColumnHeader.Index - 1
    lstvLista.Sorted = True
    If lstvLista.SortOrder = lvwAscending Then
        lstvLista.SortOrder = lvwDescending
    Else
        lstvLista.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstvLista_DblClick()
    Call Menu.mnuContextABM_Click(1)
'    If mOrigen Then
'        Call Menu.mnuContextABM_Click(1)
'    Else
'        If mAdentro = False Then
'            frmEligePeriodo.txtLibro_Id.Text = ""
'            frmEligePeriodo.txtEmp_Id.Text = auxDllActiva.FormBase.lstvLista.SelectedItem.SubItems(3)
'            frmEligePeriodo.Show
'        Else
'            Call Menu.mnuContextABM_Click(1)
'        End If
'    End If
    
End Sub

Private Sub lstvLista_GotFocus()
    
    sBarEstado.Panels(1).Text = lstvLista.ListItems.Count & " Registro(s)"

End Sub


Private Sub lstvLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyInsert Then
        Call Menu.mnuContextABM_Click(0)
    End If
    If KeyCode = vbKeyF3 Then
        Call Menu.mnuContextABM_Click(1)
    End If
    If KeyCode = vbKeyDelete Then
        Call Menu.mnuContextABM_Click(2)
    End If
    If KeyCode = vbKeyEscape Then
        mAdentro = False
        Menu.stbPrincipal.Panels(1).Text = ""
        Menu.stbPrincipal.Panels(2).Text = ""
        Menu.stbPrincipal.Panels(3).Text = ""
        Unload Me
    End If
    If KeyCode = 13 Then
        If mOrigen = False Then
            If mAdentro = False Then
               lstvLista_DblClick
            Else
                Call Menu.mnuContextABM_Click(1)
            End If
        Else
            Call Menu.mnuContextABM_Click(1)
        End If
    End If
End Sub

Private Sub lstvLista_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'habilito botones
    If lstvLista.ListItems.Count > 0 Then
        If lstvLista.SelectedItem.Selected = True Then
            'habilito botones
            tbarHerramientas.Buttons(1).Enabled = True
            tbarHerramientas.Buttons(2).Enabled = True
            tbarHerramientas.Buttons(3).Enabled = True
            tbarHerramientas.Buttons(4).Enabled = True
            tbarHerramientas.Buttons(6).Enabled = True
            tbarHerramientas.Buttons(8).Enabled = True
            tbarHerramientas.Buttons(9).Enabled = True
            tbarHerramientas.Buttons(11).Enabled = True
        Else
            'habilito botones
            tbarHerramientas.Buttons(1).Enabled = True
            tbarHerramientas.Buttons(2).Enabled = False
            tbarHerramientas.Buttons(3).Enabled = False
            tbarHerramientas.Buttons(4).Enabled = False
            tbarHerramientas.Buttons(6).Enabled = True
            tbarHerramientas.Buttons(8).Enabled = True
            tbarHerramientas.Buttons(9).Enabled = True
            tbarHerramientas.Buttons(11).Enabled = True
        End If
    Else
        'habilito botones
        tbarHerramientas.Buttons(1).Enabled = True
        tbarHerramientas.Buttons(2).Enabled = False
        tbarHerramientas.Buttons(3).Enabled = False
        tbarHerramientas.Buttons(4).Enabled = False
        tbarHerramientas.Buttons(6).Enabled = False
        tbarHerramientas.Buttons(8).Enabled = False
        tbarHerramientas.Buttons(9).Enabled = False
        tbarHerramientas.Buttons(11).Enabled = True
    End If
    
    If auxDll.Report = "" Then
        Menu.mnuContextABM(7).Visible = False
    Else
        Menu.mnuContextABM(7).Visible = True
    End If
    
    'habilito menu contextual
    If Button = 2 Then
        If lstvLista.ListItems.Count > 0 Then
            If lstvLista.SelectedItem.Selected = True Then
                Menu.mnuContextABM(0).Enabled = False
                Menu.mnuContextABM(1).Enabled = True
                Menu.mnuContextABM(2).Enabled = True
                Menu.mnuContextABM(4).Enabled = False
                Menu.mnuContextABM(6).Enabled = False
                Menu.mnuContextABM(7).Enabled = False
                Menu.mnuContextABM(9).Enabled = False
                PopupMenu Menu.ContextBaseABM, , , , Menu.mnuContextABM(1)
            Else
                Menu.mnuContextABM(0).Enabled = True
                Menu.mnuContextABM(1).Enabled = False
                Menu.mnuContextABM(2).Enabled = False
                Menu.mnuContextABM(4).Enabled = True
                Menu.mnuContextABM(6).Enabled = True
                Menu.mnuContextABM(7).Enabled = True
                Menu.mnuContextABM(9).Enabled = False
                PopupMenu Menu.ContextBaseABM, , , , Menu.mnuContextABM(0)
            End If
        Else
            Menu.mnuContextABM(0).Enabled = True
            Menu.mnuContextABM(1).Enabled = False
            Menu.mnuContextABM(2).Enabled = False
            Menu.mnuContextABM(4).Enabled = False
            Menu.mnuContextABM(6).Enabled = False
            Menu.mnuContextABM(7).Enabled = False
            Menu.mnuContextABM(9).Enabled = False
            PopupMenu Menu.ContextBaseABM, , , , Menu.mnuContextABM(0)
        End If
    End If

End Sub


Private Sub tbarHerramientas_ButtonClick(ByVal Button As Button)

   Select Case Button.Key
        Case "Nuevo"
                 Call Menu.mnuContextABM_Click(0)
        Case "Editar"
                 Call Menu.mnuContextABM_Click(1)
        Case "Eliminar"
                 Call Menu.mnuContextABM_Click(2)
        Case "Ver"
                 Call Menu.mnuContextABM_Click(9)
        Case "Buscar"
                 Call Menu.mnuContextABM_Click(6)
        Case "Refrescar"
                 Call Menu.mnuContextABM_Click(4)
        Case "Imprimir"
                 Call Menu.mnuContextABM_Click(7)
        Case "Imprime"
                 Call Menu.mnuContextABM_Click(7)
        Case "Actualiza"
                 Call Menu.mnuContextABM_Click(5)
        Case "Cerrar"
            Unload Me
   End Select
   
End Sub


