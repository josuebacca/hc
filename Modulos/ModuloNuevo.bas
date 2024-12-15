Attribute VB_Name = "ModuloNuevo"
Option Explicit

Public Sub EliminarFilasDeGrilla(Grilla As MSFlexGrid)
'Borra totalmente una grilla sin dejar Rows
    Grilla.Rows = 1
    Do While Grilla.Rows > 1
        Grilla.RemoveItem 1
    Loop
End Sub

Public Sub LimpiarFilasDeGrilla(Grilla As MSFlexGrid, Optional Fila As Long)
   Dim i As Integer
   Dim j As Integer
'Borra datos de una o todas las filas
    Dim CantidadDeFilas As Integer
    Dim ActFila As Integer
    ActFila = Grilla.row
    If Fila <> 0 Then
        CantidadDeFilas = Fila + 1
    Else
        Fila = 1
        CantidadDeFilas = Grilla.Rows
    End If
    For i = Fila To CantidadDeFilas - 1
        Grilla.row = i
        For j = 0 To Grilla.Cols - 1
            Grilla.Col = j
            Grilla.Text = ""
        Next
    Next
    Grilla.row = ActFila
End Sub

