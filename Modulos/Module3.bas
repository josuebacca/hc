Attribute VB_Name = "Module3"

Public Sub CargaTree(tr As TreeView)
    
    Dim rec As ADODB.Recordset
    Dim sql As String
    Dim Nivel As Byte
    Dim i As Integer
    Dim li As Node
    Dim image As Integer
    
    tr.Visible = False
    tr.Nodes.Clear
    tr.Nodes.Add , , "#0", vDesEjercicio, 1
    
    sql = "SELECT cta_jerarquia, cta_descrip, cta_nivel, cta_id, cta_id_padre, cta_imputable, cta_codigo FROM cuentas WHERE eje_id=" & vEjercicio & " ORDER BY cta_jerarquia"
    If DBConn.GetRecordset(rec, sql) Then
        If Not rec.EOF Then
            i = 1
            While Not rec.EOF
                If rec("cta_imputable") = "N" Then
                    image = 2
                Else
                    image = 3
                End If
                If IsNull(rec("cta_id_padre")) Then
                    tr.Nodes.Add "#0", tvwChild, "#" & rec("cta_id"), " [" & rec!Cta_Jerarquia & "] - [" & rec!Cta_CODIGO & "] - " & UCase(Left(rec!cta_descrip, 1)) & LCase(Right(rec!cta_descrip, Len(rec!cta_descrip) - 1)), image
                Else
                    tr.Nodes.Add "#" & rec("cta_id_padre"), tvwChild, "#" & rec("cta_id"), " [" & rec!Cta_Jerarquia & "] - [" & rec!Cta_CODIGO & "] - " & UCase(Left(rec!cta_descrip, 1)) & LCase(Right(rec!cta_descrip, Len(rec!cta_descrip) - 1)), image
                End If
                rec.MoveNext
            Wend
        End If
        rec.Close
    End If
    
    For Each li In tr.Nodes
        If li.Children > 0 Then
            li.Expanded = True
            'li.Image = 2
        Else
            'li.Image = 3
        End If
        Exit For
    Next
    
    If tr.Nodes.Count > 0 Then
        Set tr.SelectedItem = tr.Nodes(1)
    End If
    
    tr.Visible = True
    
End Sub
Sub AcCtrl(CtrlName As Control)

    'activa un control para permitir la edición del mismo
    If TypeOf CtrlName Is TextBox Or _
       TypeOf CtrlName Is ComboBox Then
            
        CtrlName.Enabled = True
        CtrlName.Locked = False
        CtrlName.TabStop = True
    Else
        CtrlName.Enabled = True
        CtrlName.TabStop = True
    End If
    
    If TypeOf CtrlName Is CheckBox Then
        'CtrlName.BackColor = &H8000000F
    Else
        CtrlName.BackColor = QBColor(15) 'blanco
    End If
    
    'CtrlName.TabStop = True

End Sub

Sub AcCtrlx(CtrlName As Control)

    'activa un control para permitir la edición del mismo
    If TypeOf CtrlName Is TextBox Or _
       TypeOf CtrlName Is ComboBox Then
            
        CtrlName.Enabled = True
        CtrlName.Locked = False
        CtrlName.TabStop = True
    Else
        CtrlName.Enabled = True
        CtrlName.TabStop = True
    End If
    
    If TypeOf CtrlName Is CheckBox Then
        'CtrlName.BackColor = &H8000000F
    Else
        'CtrlName.BackColor = QBColor(15) 'blanco
    End If
    
    'CtrlName.TabStop = True
End Sub

Public Sub ManejoDeErrores(ByVal ErrorCode As Long)

    Select Case ErrorCode
        Case dbSqlDuplicateKey
            Beep
            MsgBox "Código de registro duplicado." & Chr(13) & _
                             "Ingrese un código no utilizado por otro registro " & Chr(13) & _
                             "o permita que el sistema le sugiera un nuevo código.", vbCritical + vbOKOnly, App.Title
        Case dbSqlPermission
            Beep
            MsgBox "Permiso denegado." & Chr(13) & _
                             "Comuniquese con el administrador de la base de datos " & Chr(13) & _
                             "para solicitar el permiso requerido.", vbCritical + vbOKOnly, App.Title
        Case Else
            Beep
            MsgBox "Error: " & ErrorCode & " - " & _
                             "Error intentando confirmar la transacción." & Chr(13) & _
                             "Reintente la operación.", vbCritical + vbOKOnly, App.Title
    End Select

End Sub

Sub DesacCtrl(CtrlName As Control)

    'desactiva un control para evitar la edición del mismo
    If TypeOf CtrlName Is TextBox Or _
       TypeOf CtrlName Is ComboBox Then
            
        If Trim(CtrlName.Text) = "" Then
            CtrlName.Enabled = False
            CtrlName.TabStop = False
        Else
            CtrlName.Locked = True
            CtrlName.TabStop = False
        End If
    Else
        'listbox, label
        CtrlName.Enabled = False
        CtrlName.TabStop = False
    End If
    
    CtrlName.BackColor = QBColor(7) 'gris
    'CtrlName.TabStop = False

End Sub

Sub DesacCtrlx(CtrlName As Control)

''    Desactiva un control para evitar la edición del mismo
''    If TypeOf CtrlName Is TextBox Or _
''       TypeOf CtrlName Is ComboBox Then
''
''        If Trim(CtrlName.Text) = "" Then
''            CtrlName.Enabled = False
''            CtrlName.TabStop = False
''        Else
''            CtrlName.Locked = True
''            CtrlName.TabStop = False
''        End If
''    Else
''        'listbox, label
''        CtrlName.Enabled = False
''        CtrlName.TabStop = False
''    End If
     
     CtrlName.Enabled = False
     CtrlName.TabStop = False

    'CtrlName.BackColor = QBColor(7) 'gris
    'CtrlName.TabStop = False
End Sub

Public Sub CargarListView(ByRef pForm As Form, ByRef lvwSource As ListView, ByVal sql As String, Optional IDReg As String, Optional pHeaderSQL As String, Optional pImgList As Variant, Optional pMaxRec As Long)

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
    
    Dim mRaz As ADODB.Recordset
    Set mRaz = New ADODB.Recordset
    
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
        
        CantCampos = vRec.Fields.Count       ' cantidad de campos en la tabla
        
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
                    If auxDllActiva.Caption = "Actualización de Asientos..." Then
                        ArrWidthColumn(IndiceCampo) = pForm.TextWidth(TextHeader - 20)
                    End If
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
                        'Case 6
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
                    
                    If Not IsMissing(pImgList) Then
                        If IsMissing(IDReg) Or Trim(IDReg) = "" Then
                            Set itmX = lvwSource.ListItems.Add(, , DatoMostrar, 1)
                        Else
                            Set itmX = lvwSource.ListItems.Add(, "'" & vRec.Fields(CantCampos - 1) & "'", DatoMostrar, 1)
                        End If
                        itmX.Icon = 1
                        itmX.SmallIcon = 1
                    Else
                        If IsMissing(IDReg) Or Trim(IDReg) = "" Then
                            Set itmX = lvwSource.ListItems.Add(, , DatoMostrar)
                        Else
                            Set itmX = lvwSource.ListItems.Add(, "'" & vRec.Fields(CantCampos - 1) & "'", DatoMostrar)
                        End If
                    End If
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

Public Function CentrarVentana(Ventana As Form)

    'centra el formulario (Ventana) dentro del formulario principal MDI
    Ventana.Left = (frmPrincipal.ScaleWidth - Ventana.Width) / 2
    Ventana.Top = (frmPrincipal.ScaleHeight - Ventana.Height) / 2
    
End Function
