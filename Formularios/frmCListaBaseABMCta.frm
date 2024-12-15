VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCListaBaseABMCta 
   Caption         =   "Cargando Lista ..."
   ClientHeight    =   4470
   ClientLeft      =   1185
   ClientTop       =   1905
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCListaBaseABMCta.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4470
   ScaleWidth      =   8115
   Begin ComctlLib.Toolbar tbarHerramientas 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
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
            Key             =   ""
            Object.ToolTipText     =   "Buscar registro"
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
   Begin ComctlLib.TreeView TVList 
      Height          =   3735
      Left            =   30
      TabIndex        =   2
      Top             =   420
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   564
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
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
   End
   Begin Crystal.CrystalReport rptListado 
      Left            =   7530
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin ComctlLib.StatusBar sBarEstado 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   13820
            MinWidth        =   8819
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":01E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":03C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstBarra 
      Left            =   7410
      Top             =   480
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
            Picture         =   "frmCListaBaseABMCta.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":06AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":07BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":09E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":0C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCListaBaseABMCta.frx":0D98
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCListaBaseABMCta"
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

Private auxDll As CListaBaseABMCta


Public Function SetWindow(ByVal pStringSQL As String, pFieldID As String, pHeaderSQL As String, ByVal pDll As CListaBaseABMCta, Optional ByVal pMaxRec As Long)

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

    Set auxDllActivaCta = auxDll
   
'    If lstvLista.ListItems.Count > 0 Then
'        lstvLista.ListItems(1).Selected = True
'        lstvLista.SetFocus
'    End If
    
    

End Sub

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    DoEvents
    CargaTree TVList
    'CargaArbol vEjercicio, vDesEjercicio, "", TVList
End Sub


Private Sub Form_Resize()
    
    If Me.WindowState <> 1 Then
        Me.Refresh
        TVList.Top = tbarHerramientas.Top + tbarHerramientas.Height + 30
        TVList.Width = Me.Width - 180
        TVList.Height = Me.Height - (sBarEstado.Height + tbarHerramientas.Height) - 460
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Unload auxDll.FormDatos
    
End Sub







Private Sub tbarHerramientas_ButtonClick(ByVal Button As Button)

   Select Case Button.Key
        Case "Nuevo"
                 Call frmPrincipal.mnuContextABMCta_Click(0)
        Case "Editar"
                 Call frmPrincipal.mnuContextABMCta_Click(1)
        Case "Eliminar"
                 Call frmPrincipal.mnuContextABMCta_Click(2)
        Case "Refrescar"
                 Call frmPrincipal.mnuContextABMCta_Click(4)
        Case "Cerrar"
            Unload Me
        Case "Imprimir"
            'rptListado.SelectionFormula = "{cuentas.EJE_ID} = " & vEjercicio
            Call frmPrincipal.mnuContextABMCta_Click(7)
    End Select
   
End Sub


Private Sub TVList_Click()

    frmPrincipal.mnuContextABMCta(0).Enabled = False
    frmPrincipal.mnuContextABMCta(1).Enabled = False
    frmPrincipal.mnuContextABMCta(2).Enabled = False
    frmPrincipal.mnuContextABMCta(3).Enabled = False
    frmPrincipal.mnuContextABMCta(4).Enabled = False
    
End Sub


Private Sub TVList_DblClick()
  
    If Not TVList.SelectedItem Is Nothing Then
        If TVList.SelectedItem.Key <> "#0" Then
            Call frmPrincipal.mnuContextABMCta_Click(1)
        End If
    End If
  
End Sub


Private Sub TVList_GotFocus()

    sBarEstado.Panels(1).Text = IIf(TVList.Nodes.Count < 1, 0, TVList.Nodes.Count - 1) & " Registro(s)"
    
End Sub



Private Sub TVList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sql As String
    Dim rec As ADODB.Recordset
    'habilito botones
    
    If TVList.HitTest(x, y) Is Nothing Then Exit Sub
    
    If Not TVList.SelectedItem Is Nothing Then
        If TVList.SelectedItem.Key <> "#0" Then
            tbarHerramientas.Buttons(1).Enabled = True
            tbarHerramientas.Buttons(2).Enabled = True
            tbarHerramientas.Buttons(3).Enabled = True
            tbarHerramientas.Buttons(4).Enabled = False
            tbarHerramientas.Buttons(6).Enabled = True
            tbarHerramientas.Buttons(8).Enabled = False
            tbarHerramientas.Buttons(9).Enabled = True
            tbarHerramientas.Buttons(11).Enabled = True
        Else
            'habilito botones
            tbarHerramientas.Buttons(1).Enabled = True
            tbarHerramientas.Buttons(2).Enabled = False
            tbarHerramientas.Buttons(3).Enabled = False
            tbarHerramientas.Buttons(4).Enabled = False
            tbarHerramientas.Buttons(6).Enabled = False
            tbarHerramientas.Buttons(8).Enabled = False
            tbarHerramientas.Buttons(9).Enabled = True
            tbarHerramientas.Buttons(11).Enabled = True
        End If
    End If
    'habilito menu contextual
    If Button = 2 Then
        If Not TVList.SelectedItem Is Nothing Then
            If TVList.SelectedItem.Key <> "#0" Then
                'Si es una cuenta imputable no habilito la opcion nuevo
                sql = "SELECT cta_imputable FROM cuentas WHERE cta_id = " & Right(TVList.SelectedItem.Key, Len(TVList.SelectedItem.Key) - 1)
                If DBConn.GetRecordset(rec, sql) = True Then
                    If rec.EOF = False Then
                        If (rec!cta_imputable = "N" Or IsNull(rec!cta_imputable)) Then
                            frmPrincipal.mnuContextABMCta(0).Enabled = True
                            frmPrincipal.mnuContextABMCta(1).Enabled = True
                            frmPrincipal.mnuContextABMCta(2).Enabled = True
                            frmPrincipal.mnuContextABMCta(4).Enabled = True
                            frmPrincipal.mnuContextABMCta(6).Enabled = True
                            rec.Close
                            PopupMenu frmPrincipal.ContextABMCta, , , , frmPrincipal.mnuContextABMCta(0)
                        Else
                            frmPrincipal.mnuContextABMCta(0).Enabled = False
                            frmPrincipal.mnuContextABMCta(1).Enabled = True
                            frmPrincipal.mnuContextABMCta(2).Enabled = True
                            frmPrincipal.mnuContextABMCta(4).Enabled = True
                            frmPrincipal.mnuContextABMCta(6).Enabled = True
                            rec.Close
                            PopupMenu frmPrincipal.ContextABMCta, , , , frmPrincipal.mnuContextABMCta(1)
                        End If
                    End If
                    
                End If
            Else
                frmPrincipal.mnuContextABMCta(0).Enabled = True
                frmPrincipal.mnuContextABMCta(1).Enabled = False
                frmPrincipal.mnuContextABMCta(2).Enabled = False
                frmPrincipal.mnuContextABMCta(4).Enabled = False
                frmPrincipal.mnuContextABMCta(6).Enabled = False
                PopupMenu frmPrincipal.ContextABMCta, , , , frmPrincipal.mnuContextABMCta(0)
            End If
        End If
    End If

End Sub

