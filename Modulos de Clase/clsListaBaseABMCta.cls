VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListaBaseABMCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Propiedades y metodos de la clase
Option Explicit

Private pColFields As Collection
Private pCaptionWindow As String
Private pFormDatos As Form
Private pFormBase As Form
Private pFormBuscar As Form
Private pStringSQL As String
Private pFieldsSQL As String
Private pHeaderSQL As String
Private pFieldID As String
Private pWidthWindow As Long
Private pHeightWindow As Long
Private pCenterWindow As Boolean
Private pTextToolbar As Boolean
Private pMaxRecords As Long
Private pReport As String


Public Property Let CenterWindow(ByVal NewCenterWindow As Boolean)

    pCenterWindow = NewCenterWindow
    
End Property

Public Property Let FieldsSQL(ByVal NewFieldsSQL As String)

    pFieldsSQL = NewFieldsSQL

End Property

Public Property Get FieldsSQL() As String

    FieldsSQL = pFieldsSQL

End Property

Public Sub Find()
    
    'Set pFormBuscar = New BuscarEnLista
    'pFormBuscar.SetWindow pFormBase, pStringSQL, pFieldID, pHeaderSQL, pFormBase.lstvLista
    'pFormBuscar.Left = (Screen.Width - pFormBuscar.Width) / 2
    'pFormBuscar.Top = (Screen.Height - pFormBuscar.Height) / 2
    'pFormBuscar.Show

End Sub

Public Property Set FormBase(ByVal NewFormBase As Form)

    Set pFormBase = NewFormBase
    
End Property

Public Property Get FormBase() As Form

    Set FormBase = pFormBase
    
End Property

Public Property Set FormDatos(ByVal NewFormDatos As Form)

    Set pFormDatos = NewFormDatos
    
End Property

Public Property Get FormDatos() As Form

    Set FormDatos = pFormDatos
    
End Property

Public Property Get HeaderSQL() As String

    HeaderSQL = pHeaderSQL
    
End Property

Public Property Let HeaderSQL(ByVal NewValue As String)

    pHeaderSQL = NewValue
    
End Property

Public Property Let Height(ByVal NewHeight As Long)

    pHeightWindow = NewHeight
    
End Property

Public Property Get Height() As Long

    Height = pHeightWindow
    
End Property

Public Property Let MaxRecords(ByVal NewValue As Long)

    pMaxRecords = NewValue
    
End Property

Public Property Let Report(ByVal NewValue As String)

    pReport = NewValue
    
End Property

Public Sub Show()

    Dim i As Integer

    Set pFormBase = New frmCListaBaseABMCta
    
    pFormBase.SetWindow pStringSQL, pFieldID, pHeaderSQL, auxDllActivaCta, pMaxRecords
    Set auxDllActiva = Nothing
    
    'establece el nombre a la ventana
    pFormBase.Caption = pCaptionWindow
    
    'texto de los botones del toolbar
    If pTextToolbar = True Then
        For i = 1 To pFormBase.tbarHerramientas.Buttons.Count
            pFormBase.tbarHerramientas.Buttons(i).Caption = pFormBase.tbarHerramientas.Buttons(i).Key
        Next i
    End If

    'tamaño de la ventana
    If pWidthWindow > 0 Then
        pFormBase.Width = pWidthWindow
    End If

    If pHeightWindow > 0 Then
        pFormBase.Height = pHeightWindow
    End If
    
    If pCenterWindow = True Then
        CentrarVentana pFormBase
    End If
    
    'inicializo el contrl de reporte
    If pReport <> "" Then
        pFormBase.tbarHerramientas.Buttons("Imprimir").Visible = True
        pFormBase.rptListado.ReportFileName = pReport
    Else
        pFormBase.tbarHerramientas.Buttons("Imprimir").Visible = False
    End If
    
    pFormBase.Show
    
End Sub

Public Property Get sql() As String

    sql = pStringSQL

End Property

Public Property Let sql(ByVal NewSQL As String)

pStringSQL = NewSQL

End Property









Public Property Get CollectionFields() As Collection

Set CollectionFields = pColFields

End Property

Public Property Set CollectionFields(ByVal NewCollectionFields As Collection)

'Set cResultFields = SignusEngine.ColSel

End Property

Public Property Get FieldID() As String

    FieldID = pFieldID

End Property

Public Property Let FieldID(ByVal NewFieldID As String)

    pFieldID = NewFieldID

End Property

Public Property Get Caption() As String

Caption = pCaptionWindow

End Property

Public Property Let Caption(ByVal NewCaption As String)

pCaptionWindow = NewCaption

End Property




Public Property Let TextToolbar(ByVal NewTextToolbar As Boolean)

    pTextToolbar = NewTextToolbar
    
End Property

Public Property Get TextToolbar() As Boolean

    TextToolbar = pTextToolbar
    
End Property

Public Property Let Width(ByVal NewWidth As Long)

    pWidthWindow = NewWidth
    
End Property

Public Property Get Width() As Long

    Width = pWidthWindow
    
End Property

Public Property Get CenterWindow() As Boolean

    CenterWindow = pCenterWindow
    
End Property

Private Sub Class_Initialize()

'establecer valores por defecto para las propiedades
pCaptionWindow = ""
pWidthWindow = 0
pHeightWindow = 0
pCenterWindow = False
pTextToolbar = False
pMaxRecords = 0
pReport = ""
pHeaderSQL = ""

End Sub




