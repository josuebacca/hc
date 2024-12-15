Attribute VB_Name = "modDllBusqueda"
Option Explicit

Public ColSel As Collection
Public sSQL As String
Public hSQL As String
Public dConn As Connection
Public sField As String
Public sTitulo As String
Public iMaxRecords As Long
Public sOrderBy As String

'Campos para la búsqueda
Public campo1 As String
Public campo2 As String
Public campo3 As String
Public campo4 As String
Public camponumerico As Boolean
