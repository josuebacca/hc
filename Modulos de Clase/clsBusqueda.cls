VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit





Public Property Let Headers(ByVal NewValue As String)

    hSQL = NewValue
    
End Property

Public Property Get Headers() As String

    Headers = hSQL
    
End Property

Public Sub Show()

    frmCBusqueda.Show vbModal

End Sub

Public Property Get sql() As String
sql = sSQL
End Property

Public Property Let sql(ByVal sNewSQL As String)

sSQL = sNewSQL

End Property






Public Property Get Conn() As Connection

    Set Conn = dConn

End Property

Public Property Set Conn(ByVal sNewConn As Connection)

Set dConn = sNewConn

End Property


Public Property Get ResultFields() As Collection

Set ResultFields = ColSel

End Property

Public Property Set ResultFields(ByVal cNewResultFields As Collection)

'Set cResultFields = SignusEngine.ColSel

End Property

Public Property Get Field() As String

Field = sField

End Property

Public Property Let Field(ByVal sNewField As String)

sField = " " & sNewField & " "

End Property

Public Property Get Titulo() As String

Titulo = sTitulo

End Property

Public Property Let Titulo(ByVal sNewTitulo As String)

sTitulo = sNewTitulo

End Property

Public Property Get MaxRecords() As Long

MaxRecords = iMaxRecords

End Property

Public Property Let MaxRecords(ByVal iNewMaxRecords As Long)

iMaxRecords = iNewMaxRecords

End Property

Private Sub Class_Initialize()

'establecer valores por defecto para las propiedades
Me.MaxRecords = 20
Me.OrderBy = IIf(Trim(Me.OrderBy) = "", Me.Field, Me.OrderBy)
Set ColSel = New Collection

End Sub


Public Property Get OrderBy() As String

OrderBy = sOrderBy

End Property

Public Property Let OrderBy(ByVal sNewOrderBy As String)

sOrderBy = " " & sNewOrderBy & " "

End Property

