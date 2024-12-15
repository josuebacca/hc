Attribute VB_Name = "ModuleNroLetra"
Function LeeNro(Numero As Double, lcad1 As Integer, lcad2 As Integer, tmoneda As String, cad1car As String, cad2car As String) As String
'----------------------------------------------------------
'DATOS QUE SE DEBEN SUMINISTRAR:
'-------------------------------
'Acepta como máximo 2 digitos en la parte decimal.
'Numero  = Maximo 999.999.999,99 Minimo 0.
'lcad1   = Longitud de la 1era cadena. El tamaño debe ser >= 6* y <=65536 (2^16 tamaño maximo de String) * Si no se utiliza tipo de moneda.
'lcad2   = Longitud de la 2da  cadena. El tamaño debe ser >= 0  y <=65531 (2^16 - 5 tamaño maximo de String - 5)
'tmoneda = Tipo de moneda. Debe ser "$" PESOS, "U$S" DOLARES ESTADOUNIDENSES ó "" Sin tipo de moneda.
'cad1car = Caracter de relleno para la 1era cadena. Ej "*" .Debe haber un caracter o espacio.
'cad2car = Caracter de relleno para la 2da  cadena. Ej " " .Debe haber un caracter o espacio.
'FUNCIONALIDAD :
'--------------
'Recibe un numero y devuelve 1 cadena.
'Si es necesario la funcion hace silabeo.
'Si se le indica completa la cadena con simbolos.
'Se debe mostrar de la siguiente forma:
'Ejemplo: cadnumerica=LeeNro(1234.27,20,60,"$","*","*")
'         cadnumerica="PESOS UN MIL DOS- **CIENTOS TREINTA Y CUATRO CON OCHENTA CENTAVOS.**************"
'         Mostrar 1ero Mid(cadnumerica,1,20)  -> PESOS UN MIL DOS- **
'         Mostrar 2do  Mid(cadnumerica,21,60) -> CIENTOS TREINTA Y CUATRO CON OCHENTA CENTAVOS.**************
'----------------------------------------------------------
Dim unidades(10) As String
'              123456789¦12345  12345  12345  12345  12345
unidades(1) = "UN             ¦ UN   ¦ -----¦ -----¦ -----"
unidades(2) = "DOS            ¦ DOS  ¦ -----¦ -----¦ -----"
unidades(3) = "TRES           ¦ TRES ¦ -----¦ -----¦ -----"
unidades(4) = "CUATRO         ¦ CUA  ¦ TRO  ¦ -----¦ -----"
unidades(5) = "CINCO          ¦ CIN  ¦ CO   ¦ -----¦ -----"
unidades(6) = "SEIS           ¦ SEIS ¦ -----¦ -----¦ -----"
unidades(7) = "SIETE          ¦ SIE  ¦ TE   ¦ -----¦ -----"
unidades(8) = "OCHO           ¦ O    ¦ CHO  ¦ -----¦ -----"
unidades(9) = "NUEVE          ¦ NUE  ¦ VE   ¦ -----¦ -----"
'----------------------------------------------------------
Dim dec10_19(10) As String
'              123456789¦12345  12345  12345  12345  12345
dec10_19(1) = "ONCE           ¦ ON   ¦ CE   ¦ -----¦ -----"
dec10_19(2) = "DOCE           ¦ DO   ¦ CE   ¦ -----¦ -----"
dec10_19(3) = "TRECE          ¦ TRE  ¦ CE   ¦ -----¦ -----"
dec10_19(4) = "CATORCE        ¦ CA   ¦ TOR  ¦ CE   ¦ -----"
dec10_19(5) = "QUINCE         ¦ QUIN ¦ CE   ¦ -----¦ -----"
dec10_19(6) = "DIECISEIS      ¦ DIE  ¦ CI   ¦ SEIS ¦ -----"
dec10_19(7) = "DIECISIETE     ¦ DIE  ¦ CI   ¦ SIE  ¦ TE   "
dec10_19(8) = "DIECIOCHO      ¦ DIE  ¦ CI   ¦ O    ¦ CHO  "
dec10_19(9) = "DIECINUEVE     ¦ DIE  ¦ CI   ¦ NUE  ¦ VE   "
'----------------------------------------------------------
Dim veinti(10) As String
'            123456789¦12345  12345  12345  12345  12345
veinti(1) = "VEINTIUN       ¦ VEIN ¦ TI   ¦ UN   ¦ -----"
veinti(2) = "VEINTIDOS      ¦ VEIN ¦ TI   ¦ DOS  ¦ -----"
veinti(3) = "VEINTITRES     ¦ VEIN ¦ TI   ¦ TRES ¦ -----"
veinti(4) = "VEINTICUATRO   ¦ VEIN ¦ TI   ¦ CUA  ¦ TRO  "
veinti(5) = "VEINTICINCO    ¦ VEIN ¦ TI   ¦ CIN  ¦ CO   "
veinti(6) = "VEINTISEIS     ¦ VEIN ¦ TI   ¦ SEIS ¦ -----"
veinti(7) = "VEINTISIETE    ¦ VEIN ¦ TI   ¦ SIE  ¦ TE   "
veinti(8) = "VEINTIOCHO     ¦ VEIN ¦ TI   ¦ O    ¦ CHO  "
veinti(9) = "VEINTINUEVE    ¦ VEIN ¦ TI   ¦ NUE  ¦ VE   "
'----------------------------------------------------------
Dim decenas(10) As String
'             123456789¦12345  12345  12345  12345  12345
decenas(1) = "DIEZ           ¦ DIEZ ¦ -----¦ -----¦ -----"
decenas(2) = "VEINTE         ¦ VEIN ¦ TE   ¦ -----¦ -----"
decenas(3) = "TREINTA        ¦ TREIN¦ TA   ¦ -----¦ -----"
decenas(4) = "CUARENTA       ¦ CUA  ¦ REN  ¦ TA   ¦ -----"
decenas(5) = "CINCUENTA      ¦ CIN  ¦ CUEN ¦ TA   ¦ -----"
decenas(6) = "SESENTA        ¦ SE   ¦ SEN  ¦ TA   ¦ -----"
decenas(7) = "SETENTA        ¦ SE   ¦ TEN  ¦ TA   ¦ -----"
decenas(8) = "OCHENTA        ¦ O    ¦ CHEN ¦ TA   ¦ -----"
decenas(9) = "NOVENTA        ¦ NO   ¦ VEN  ¦ TA   ¦ -----"
'----------------------------------------------------------
Dim centenas(10) As String
'              123456789¦12345  12345  12345  12345  12345
centenas(1) = "CIENTO         ¦ CIEN ¦ TO   ¦ -----¦ -----"
centenas(2) = "DOSCIENTOS     ¦ DOS  ¦ CIEN ¦ TOS  ¦ -----"
centenas(3) = "TRESCIENTOS    ¦ TRES ¦ CIEN ¦ TOS  ¦ -----"
centenas(4) = "CUATROCIENTOS  ¦ CUA  ¦ TRO  ¦ CIEN ¦ TOS  "
centenas(5) = "QUINIENTOS     ¦ QUI  ¦ NIEN ¦ TOS  ¦ -----"
centenas(6) = "SEISCIENTOS    ¦ SEIS ¦ CIEN ¦ TOS  ¦ -----"
centenas(7) = "SETECIENTOS    ¦ SE   ¦ TE   ¦ CIEN ¦ TOS  "
centenas(8) = "OCHOCIENTOS    ¦ O    ¦ CHO  ¦ CIEN ¦ TOS  "
centenas(9) = "NOVECIENTOS    ¦ NO   ¦ VE   ¦ CIEN ¦ TOS  "
'----------------------------------------------------------
Dim otras(10) As String
'           123456789¦12345  12345  12345  12345  12345
otras(1) = "UNO            ¦ UNO  ¦ -----¦ -----¦ -----"
otras(2) = "VEINTIUNO      ¦ VEIN ¦ TI   ¦ UNO  ¦ -----"
otras(3) = "CIEN           ¦ CIEN ¦ -----¦ -----¦ -----"
otras(4) = "MIL            ¦ MIL  ¦ -----¦ -----¦ -----"
otras(5) = "MILLON         ¦ MI   ¦ LLON ¦ -----¦ -----"
otras(6) = "MILLONES       ¦ MI   ¦ LLO  ¦ NES  ¦ -----"
otras(7) = "CENTAVOS       ¦ CEN  ¦ TA   ¦ VOS  ¦ -----"
otras(8) = "CENTAVO        ¦ CEN  ¦ TA   ¦ VO   ¦ -----"
otras(9) = "CON            ¦ CON  ¦ -----¦ -----¦ -----"
'----------------------------------------------------------
Dim cade As String, st As String
If Int(Numero) = 0 Then
    cade = " CERO "      ' Tomo entero x 0,12
Else
    cade = " "
End If
Dim NumeroInt As Double, CentaNro As Double

NumeroInt = Int(Numero)
CentaNro = Format(Numero - NumeroInt, "###,###,##0.00")
Centa = Trim(Str(CentaNro * 100))        ' Me quedo con los 2 dig. de centavos

Dim Trios(3) As Integer
For I = 1 To IIf(Centa > 0, 2, 1)
    If I = 1 Then
        nro = Right("000000000" + LTrim(Str(Int(Numero))), 9) ' Completo el Numero con ceros
    Else
        nro = Right("000000000" + LTrim(Str(Val(Centa))), 9) ' Completo el Numero con ceros
        cade = cade + "CON "
    End If
    Trios(1) = Val(Mid(nro, 1, 3)) ' Primer trio de dig.
    Trios(2) = Val(Mid(nro, 4, 3)) ' Segundo trio de dig.
    Trios(3) = Val(Mid(nro, 7, 3)) ' Tercer trio de dig.

    ciclo = 0                           ' Un ciclo x c/ trio
    Do While ciclo < 3
        ciclo = ciclo + 1
        pt = Int(Trios(ciclo) / 100)    ' Ej: 123 pt=1   Primer termino
        dt = Trios(ciclo) - (pt * 100)  ' 123-100 dt=23  decena del terrmino
        st = Int(dt / 10)               ' 23/10   st=2   segundo termino
        tt = dt - (st * 10)             ' 23-20   tt=3   tercer termino
        If Trios(ciclo) <> 0 Then
            If Trios(ciclo) = 100 Then     ' Caso Especial.
                cade = cade + "CIEN "
            End If
            If pt > 0 And Trios(ciclo) <> 100 Then  ' cien <> ciento
                cade = cade + Trim(Mid(centenas(pt), 1, 15)) + " "
            End If
            If dt > 10 And dt < 20 Then              ' Si es entre 11 y 19
                cade = cade + Trim(Mid(dec10_19(tt), 1, 15)) + " "
            Else                                     ' else es entre 0 y 10
                If dt > 20 And dt < 30 Then          'o es entre 20 y 99.
                    If dt = 21 And ciclo = 3 And I = 1 Then  ' Si es el tercer trio es UNO
                        cade = cade + "VEINTIUNO "
                    Else
                        cade = cade + Trim(Mid(veinti(tt), 1, 15)) + " "
                    End If
                Else
                    If dt >= 30 Or dt = 10 Or dt = 20 Then
                        cade = cade + Trim(Mid(decenas(st), 1, 15)) + " "
                    End If
                End If
                
                If tt > 0 And dt > 29 Then           ' los >29 son Ej: treinta y uno
                   cade = cade + "Y "
                End If
                
                If tt > 1 And (dt < 10 Or dt > 30) Then
                    cade = cade + Trim(Mid(unidades(tt), 1, 15)) + " " ' Busco la unidad sin problemas
                End If
                If tt = 1 And (dt < 10 Or dt > 30) Then
                    If ciclo = 3 Then                ' Si es el tercer trio es UNO
                        cade = cade + IIf(I = 1, "UNO ", Trim(Mid(unidades(tt), 1, 15)) + " ")
                    Else
                        cade = cade + Trim(Mid(unidades(tt), 1, 15)) + " "  ' else es UN. Ej: TREINTA
                    End If                            ' Y UN MIL TREINTA Y UNO.
                End If
            End If
            Select Case ciclo
            Case Is = 1
                If pt = 0 And st = 0 And tt = 1 Then ' solo si es 1 millon o un
                    cade = cade + "MILLON "          ' o un millon y algo mas
                Else
                    cade = cade + "MILLONES "
                End If
            Case Is = 2
                    cade = cade + "MIL "
            End Select
        End If
    Loop
Next
'La leyenda esta conformada falta agregar el punto final y el tipo de moneda.
Select Case tmoneda
Case Is = "$"
    cade = "PESOS" + cade + IIf(Centa > 0, IIf(Centa = 1, "CENTAVO", "CENTAVOS"), "")
Case Is = "U$S"
    cade = "DOLARES ESTADOUNIDENSES" + cade + IIf(Centa > 0, IIf(Centa = 1, "CENTAVO", "CENTAVOS"), "")
Case Is = ""
    cade = LTrim(cade) + IIf(Centa > 0, IIf(Centa = 1, "CENTAVO", "CENTAVOS"), "")
End Select
'Ahora divido la cadena en dos.
Dim compalabra As Integer, finpalabra As Integer
Dim cad1 As String, cad2 As String
If Len(cade) < lcad1 Then
    cad1 = Trim(cade) + "." + " " + String(lcad1 - Len(Trim(cade)), cad1car)
    cad2 = String(lcad2, cad2car)
Else
    MargenDerecho = lcad1
    Do While MargenDerecho >= 1
        If Mid(cade, MargenDerecho, 1) = " " Then
            Exit Do
        End If
        MargenDerecho = MargenDerecho - 1
    Loop
    compalabra = MargenDerecho + 1
    Do While MargenDerecho <= lcad1 + lcad2
        MargenDerecho = MargenDerecho + 1
        If Mid(cade, MargenDerecho, 1) = " " Then
            Exit Do
        End If
        palabra = palabra + Mid(cade, MargenDerecho, 1)
    Loop
    finpalabra = MargenDerecho
    For I = 1 To 9
        If Trim(Mid(unidades(I), 1, 15)) = palabra Then
            silabas unidades(I), cade, cad1, cad2, lcad1 - compalabra + 1, compalabra - 1, finpalabra, lcad1, lcad2, cad1car, cad2car
            Exit For
        End If
        If Trim(Mid(dec10_19(I), 1, 15)) = palabra Then
            silabas dec10_19(I), cade, cad1, cad2, lcad1 - compalabra + 1, compalabra - 1, finpalabra, lcad1, lcad2, cad1car, cad2car
            Exit For
        End If
        If Trim(Mid(veinti(I), 1, 15)) = palabra Then
            silabas veinti(I), cade, cad1, cad2, lcad1 - compalabra + 1, compalabra - 1, finpalabra, lcad1, lcad2, cad1car, cad2car
            Exit For
        End If
        If Trim(Mid(decenas(I), 1, 15)) = palabra Then
            silabas decenas(I), cade, cad1, cad2, lcad1 - compalabra + 1, compalabra - 1, finpalabra, lcad1, lcad2, cad1car, cad2car
            Exit For
        End If
        If Trim(Mid(centenas(I), 1, 15)) = palabra Then
            silabas centenas(I), cade, cad1, cad2, lcad1 - compalabra + 1, compalabra - 1, finpalabra, lcad1, lcad2, cad1car, cad2car
            Exit For
        End If
        If Trim(Mid(otras(I), 1, 15)) = palabra Then
            silabas otras(I), cade, cad1, cad2, lcad1 - compalabra + 1, compalabra - 1, finpalabra, lcad1, lcad2, cad1car, cad2car
            Exit For
        End If
        If I = 9 Then           ' No encontro la palabra a dividir
            cad1 = Mid(Trim(cade) + ".", 1, lcad1)
            cad2 = Trim(Mid(Trim(cade) + ".", lcad1 + 1, lcad2)) + String(lcad2 - Len(Trim(Mid(cade + ".", lcad1 + 1, lcad2))), cad2car)
        End If
    Next
    
End If
LeeNro = Trim(cad1 + cad2)
End Function

Private Sub silabas(palvector As String, cade As String, cad1 As String, cad2 As String, cancarac As Integer, compalabra As Integer, finpalabra As Integer, lcad1 As Integer, lcad2 As Integer, cad1car As String, cad2car As String)
    If Len(Trim(Mid(palvector, 18, 5))) + Len(Trim(Mid(palvector, 25, 5))) + Len(Trim(Mid(palvector, 32, 5))) < cancarac Then
        cad1 = Mid(cade, 1, compalabra) + Trim(Mid(palvector, 18, 5)) + Trim(Mid(palvector, 25, 5)) + Trim(Mid(palvector, 32, 5)) + "-"
        cad1 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad1")
        sil4 = IIf(Trim(Mid(palvector, 39, 5)) <> "-----", Trim(Mid(palvector, 39, 5)), "")
        cad2 = sil4 + Mid(cade, finpalabra, lcad2) + "."
        cad2 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad2")
        Exit Sub
    End If
    If Len(Trim(Mid(palvector, 18, 5))) + Len(Trim(Mid(palvector, 25, 5))) < cancarac Then
        cad1 = Mid(cade, 1, compalabra) + Trim(Mid(palvector, 18, 5)) + Trim(Mid(palvector, 25, 5)) + "-"
        cad1 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad1")
        sil3 = IIf(Trim(Mid(palvector, 32, 5)) <> "-----", Trim(Mid(palvector, 32, 5)), "")
        sil4 = IIf(Trim(Mid(palvector, 39, 5)) <> "-----", Trim(Mid(palvector, 39, 5)), "")
        cad2 = sil3 + sil4 + Mid(cade, finpalabra, lcad2) + "."
        cad2 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad2")
        Exit Sub
    End If
    If Len(Trim(Mid(palvector, 18, 5))) < cancarac Then
        cad1 = Mid(cade, 1, compalabra) + Trim(Mid(palvector, 18, 5)) + "-"
        cad1 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad1")
        sil2 = IIf(Trim(Mid(palvector, 25, 5)) <> "-----", Trim(Mid(palvector, 25, 5)), "")
        sil3 = IIf(Trim(Mid(palvector, 32, 5)) <> "-----", Trim(Mid(palvector, 32, 5)), "")
        sil4 = IIf(Trim(Mid(palvector, 39, 5)) <> "-----", Trim(Mid(palvector, 39, 5)), "")
        cad2 = sil2 + sil3 + sil4 + Mid(cade, finpalabra, lcad2) + "."
        cad2 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad2")
    Else
        cad1 = Mid(cade, 1, compalabra)
        cad1 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad1")
        cad2 = Trim(Mid(palvector, 1, 15)) + Mid(cade, finpalabra, lcad2) + "."
        cad2 = completa(cad1, cad2, lcad1, lcad2, cad1car, cad2car, "cad2")
    End If
End Sub

Function completa(cad1 As String, cad2 As String, lcad1 As Integer, lcad2 As Integer, cad1car As String, cad2car As String, nrocadena As String) As String
    If nrocadena = "cad1" Then
        If lcad1 - Len(cad1) > 1 Then
            cad1 = cad1 + " " + String(lcad1 - Len(cad1) - 1, cad1car)
        Else
            cad1 = cad1 + IIf(lcad1 - Len(cad1) = 1, " ", "")
        End If
        completa = cad1
    Else
        If lcad2 - Len(cad2) > 1 Then
            cad2 = cad2 + " " + String(lcad2 - Len(cad2) - 1, cad2car)
        Else
            cad2 = cad2 + IIf(lcad2 - Len(cad2) = 1, " ", "")
        End If
        completa = cad2
    End If
End Function

