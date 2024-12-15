Attribute VB_Name = "Module8"
Dim Des As String
Dim DesCent As String
Dim Num As Long
Dim Ban As Integer
Dim Centa As Currency
Dim Numero As Currency

Public Function LeeNumero(ByRef Numero_A_Leer)
    If Numero_A_Leer < 0 Or Numero_A_Leer >= 10000000 Then
        Des = ""
        Exit Function
    End If
    Numero = Numero_A_Leer
    Num = Int(Numero)
    Ban = 0
    Des = ""
    DesCent = ""
    Centa = Numero_A_Leer
    Call DescNum
    LeeNumero = Des
End Function

Public Function DescNum()
LongNum = Len(Trim(Str(Num)))
If Num > 0 Then
'  Centa = Val(Right(Mid(Centa, 13, 2), 2))
    Centa = Val(Right(Format(Centa, "##############.00"), 2))
    If Centa = 0 Then
        DesCent = " CON CERO CENTAVOS"
    Else
        Call Calc_Cent(Centa)
        DesCent = " CON " & DesCent & " CENTAVOS"
    End If
Else
'    Centa = Val(Right(Mid(Centa, 13, 2), 2))
    Centa = Val(Right(Centa, 2))
    If Centa = 0 Then
        DesCent = "CERO CENTAVOS"
    Else
        Call Calc_Cent(Centa)
        DesCent = DesCent & " CENTAVOS"
    End If
End If

If LongNum = 7 Then
    Select Case Val(Left(Trim(Str(Num)), 1))
        Case 1
            Des = "UN MILLON"
        Case 2
            Des = "DOS MILLONES"
        Case 3
            Des = "TRES MILLONES"
        Case 4
            Des = "CUATRO MILLONES"
        Case 5
            Des = "CINCO MILLONES"
        Case 6
            Des = "SEIS MILLONES"
        Case 7
            Des = "SIETE MILLONES"
        Case 8
            Des = "OCHO MILLONES"
        Case 9
            Des = "NUEVE MILLONES"
    End Select
    Num = Val(Right(Trim(Str(Num)), 6))
End If
LongNum = Len(Trim(Str(Num)))
If LongNum = 6 Then
    If Val(Left(Trim(Str(Num)), 1)) = 1 Then
        If Val(Left(Trim(Str(Num)), 3)) = 100 Then
            Des = Des & " CIEN MIL"
            Num = Val(Right(Trim(Str(Num)), 3))
        Else
            Call NumCiento(Val(Left(Trim(Str(Num)), 3)))
            Des = Des & " MIL"
            Num = Val(Right(Trim(Str(Num)), 3))
        End If
    Else
        Call NumCiento(Val(Left(Trim(Str(Num)), 3)))
        Des = Des & " MIL"
        Num = Val(Right(Trim(Str(Num)), 3))
    End If
End If
LongNum = Len(Trim(Str(Num)))
If LongNum = 5 Then
    Call CalcDec(Val(Left(Trim(Str(Num)), 2)))
    Des = Des & " MIL"
    Num = Val(Right(Trim(Str(Num)), 3))
End If
LongNum = Len(Trim(Str(Num)))
If LongNum = 4 Then
    NumAux = Val(Left(Trim(Str(Num)), 1))
    Select Case NumAux
        Case 1
            Des = Des & " MIL"
        Case 2
            Des = Des & " DOS MIL"
        Case 3
            Des = Des & " TRES MIL"
        Case 4
            Des = Des & " CUATRO MIL"
        Case 5
            Des = Des & " CINCO MIL"
        Case 6
            Des = Des & " SEIS MIL"
        Case 7
            Des = Des & " SIETE MIL"
        Case 8
            Des = Des & " OCHO MIL"
        Case 9
            Des = Des & " NUEVE MIL"
    End Select
    Num = Val(Right(Trim(Str(Num)), 3))
End If
LongNum = Len(Trim(Str(Num)))
If LongNum = 3 Then
    Call Num_Ciento(Int(Val(Mid(Trim(Str(Num)), 1, 3)))) ''''
    Des = Des & " PESOS"
End If
LongNum = Len(Trim(Str(Num)))
If LongNum = 2 Then
    Call Calc_Dec(Int(Val(Mid(Trim(Str(Num)), 1, 2)))) ''''
    Des = Des & " PESOS"
End If
LongNum = Len(Trim(Str(Num)))

If LongNum = 1 Then
    Call CalcUn(Int(Val(Mid(Trim(Str(Num)), 1, 1)))) ''''
    Des = Des & " PESOS"
End If
If Int(Numero) > 0 Then
    Des = Des + DesCent
Else
    Des = DesCent
End If
End Function




Function NumCiento(ByRef NM_U)
MN = Val(Left(Trim(Str(Num)), 1))
    Select Case MN
        Case 1
            If Num = 100 Then
                Des = Des & " CIEN"
            Else
                Des = Des & " CIENTO"
            End If
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 2
            Des = Des & " DOSCIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 3
            Des = Des & " TRESCIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 4
            Des = Des & " CUATROCIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 5
            Des = Des & " QUINIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 6
            Des = Des & " SEISCIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 7
            Des = Des & " SETECIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 8
            Des = Des & " OCHOCIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
        Case 9
            Des = Des & " NOVECIENTOS"
            Call CalcDec(Val(Right(Trim(Str(NM_U)), 2)))
    End Select
End Function



Function CalcDec(ByRef MM)
Select Case MM
    Case 10
        Des = Des & " DIEZ"
    Case 11
        Des = Des & " ONCE"
    Case 12
        Des = Des & " DOCE"
    Case 13
        Des = Des & " TRECE"
    Case 14
        Des = Des & " CATORCE"
    Case 15
        Des = Des & " QUINCE"
    Case 16
        Des = Des & " DIECISEIS"
    Case 17
        Des = Des & " DIECISIETE"
    Case 18
        Des = Des & " DIECIOCHO"
    Case 19
        Des = Des & " DIECINUEVE"
    Case 20
        Des = Des & " VEINTE"
    Case 21
        Des = Des & " VEINTIUN"
    Case 22
        Des = Des & " VEINTIDOS"
    Case 23
        Des = Des & " VEINTITRES"
    Case 24
        Des = Des & " VEINTICUATRO"
    Case 25
        Des = Des & " VEINTICINCO"
    Case 26
        Des = Des & " VEINTISEIS"
    Case 27
        Des = Des & " VEINTISIETE"
    Case 28
        Des = Des & " VEINTIOCHO"
    Case 29
        Des = Des & " VEINTINUEVE"
    Case 30
        Des = Des & " TREINTA"
    Case 40
        Des = Des & " CUARENTA"
    Case 50
        Des = Des & " CINCUENTA"
    Case 60
        Des = Des & " SESENTA"
    Case 70
        Des = Des & " SETENTA"
    Case 80
        Des = Des & " OCHENTA"
    Case 90
        Des = Des & " NOVENTA"
    Case Else
        If MM < 10 Then
            Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
        Else
            NumAux = Val(Left(Trim(Str(MM)), 1))
            Select Case NumAux
                Case 3
                    Des = Des & " TREINTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 4
                    Des = Des & " CUARENTA Y "
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 5
                    Des = Des & " CINCUENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 6
                    Des = Des & " SESENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 7
                    Des = Des & " SETENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 8
                    Des = Des & " OCHENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 9
                    Des = Des & " NOVENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
            End Select
        End If
    End Select
End Function


Function CalcUn(ByRef NumAux)
    Select Case NumAux
        Case 1
            Des = Des & " UN"
        Case 2
            Des = Des & " DOS"
        Case 3
            Des = Des & " TRES"
        Case 4
            Des = Des & " CUATRO"
        Case 5
            Des = Des & " CINCO"
        Case 6
            Des = Des & " SEIS"
        Case 7
            Des = Des & " SIETE"
        Case 8
            Des = Des & " OCHO"
        Case 9
            Des = Des & " NUEVE"
    End Select
End Function



Function Num_Ciento(ByRef NM_U)
MN = Val(Left(Trim(Str(Num)), 1))
Select Case MN
    Case 1
        If Right(NM_U, 1) = 0 Then
            Des = Des & " CIEN"
        Else
            Des = Des & " CIENTO"
        End If
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 2
        Des = Des & " DOSCIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 3
        Des = Des & " TRESCIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 4
        Des = Des & " CUATROCIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 5
        Des = Des & " QUINIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 6
        Des = Des & " SEISCIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 7
        Des = Des & " SETECIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 8
        Des = Des & " OCHOCIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    Case 9
        Des = Des & " NOVECIENTOS"
        Call Calc_Dec(Val(Right(Trim(Str(NM_U)), 2)))
    End Select
End Function



Function Calc_Dec(ByRef MM)
Select Case MM
    Case 10
        Des = Des & " DIEZ"
    Case 11
        Des = Des & " ONCE"
    Case 12
        Des = Des & " DOCE"
    Case 13
        Des = Des & " TRECE"
    Case 14
        Des = Des & " CATORCE"
    Case 15
        Des = Des & " QUINCE"
    Case 16
        Des = Des & " DIECISEIS"
    Case 17
        Des = Des & " DIECISIETE"
    Case 18
        Des = Des & " DIECIOCHO"
    Case 19
        Des = Des & " DIECINUEVE"
    Case 20
        Des = Des & " VEINTE"
    Case 21
        Des = Des & " VEINTIUN"
    Case 22
        Des = Des & " VEINTIDOS"
    Case 23
        Des = Des & " VEINTITRES"
    Case 24
        Des = Des & " VEINTICUATRO"
    Case 25
        Des = Des & " VEINTICINCO"
    Case 26
        Des = Des & " VEINTISEIS"
    Case 27
        Des = Des & " VEINTISIETE"
    Case 28
        Des = Des & " VEINTIOCHO"
    Case 29
        Des = Des & " VEINTINUEVE"
    Case 30
        Des = Des & " TREINTA"
    Case 40
        Des = Des & " CUARENTA"
    Case 50
        Des = Des & " CINCUENTA"
    Case 60
        Des = Des & " SESENTA"
    Case 70
        Des = Des & " SETENTA"
    Case 80
        Des = Des & " OCHENTA"
    Case 90
        Des = Des & " NOVENTA"
    Case Else
        If MM < 10 Then
            Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
        Else
            NumAux = Val(Left(Trim(Str(MM)), 1))
            Select Case NumAux
                Case 3
                    Des = Des & " TREINTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 4
                    Des = Des & " CUARENTA Y "
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 5
                    Des = Des & " CINCUENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 6
                    Des = Des & " SESENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 7
                    Des = Des & " SETENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 8
                    Des = Des & " OCHENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 9
                    Des = Des & " NOVENTA Y"
                    Call CalcUn(Val(Right(Trim(Str(MM)), 1)))
            End Select
        End If
    End Select
End Function


Function Calc_Un(ByRef NumAux)
    Select Case NumAux
        Case 1
            Des = Des & " UNO"
        Case 2
            Des = Des & " DOS"
        Case 3
            Des = Des & " TRES"
        Case 4
            Des = Des & " CUATRO"
        Case 5
            Des = Des & " CINCO"
        Case 6
            Des = Des & " SEIS"
        Case 7
            Des = Des & " SIETE"
        Case 8
            Des = Des & " OCHO"
        Case 9
            Des = Des & " NUEVE"
    End Select
End Function


Function Calc_Cent(ByRef MM)
Select Case MM
    Case 10
        DesCent = DesCent & " DIEZ"
    Case 11
        DesCent = DesCent & " ONCE"
    Case 12
        DesCent = DesCent & " DOCE"
    Case 13
        DesCent = DesCent & " TRECE"
    Case 14
        DesCent = DesCent & " CATORCE"
    Case 15
        DesCent = DesCent & " QUINCE"
    Case 16
        DesCent = DesCent & " DIECISEIS"
    Case 17
        DesCent = DesCent & " DIECISIETE"
    Case 18
        DesCent = DesCent & " DIECIOCHO"
    Case 19
        DesCent = DesCent & " DIECINUEVE"
    Case 20
        DesCent = DesCent & " VEINTE"
    Case 21
        DesCent = DesCent & " VEINTIUN"
    Case 22
        DesCent = DesCent & " VEINTIDOS"
    Case 23
        DesCent = DesCent & " VEINTITRES"
    Case 24
        DesCent = DesCent & " VEINTICUATRO"
    Case 25
        DesCent = DesCent & " VEINTICINCO"
    Case 26
        DesCent = DesCent & " VEINTISEIS"
    Case 27
        DesCent = DesCent & " VEINTISIETE"
    Case 28
        DesCent = DesCent & " VEINTIOCHO"
    Case 29
        DesCent = DesCent & " VEINTINUEVE"
    Case 30
        DesCent = DesCent & " TREINTA"
    Case 40
        DesCent = DesCent & " CUARENTA"
    Case 50
        DesCent = DesCent & " CINCUENTA"
    Case 60
        DesCent = DesCent & " SESENTA"
    Case 70
        DesCent = DesCent & " SETENTA"
    Case 80
        DesCent = DesCent & " OCHENTA"
    Case 90
        DesCent = DesCent & " NOVENTA"
    Case Else
        If MM < 10 Then
            Call Calc_UnCent(Val(Right(Trim(Str(MM)), 1)))
        Else
            NumAux = Val(Left(Trim(Str(MM)), 1))
            Select Case NumAux
                Case 3
                    DesCent = DesCent & " TREINTA Y"
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 4
                    DesCent = DesCent & " CUARENTA Y "
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 5
                    DesCent = DesCent & " CINCUENTA Y"
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 6
                    DesCent = DesCent & " SESENTA Y"
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 7
                    DesCent = DesCent & " SETENTA Y"
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 8
                    DesCent = DesCent & " OCHENTA Y"
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
                Case 9
                    DesCent = DesCent & " NOVENTA Y"
                    Call Cent_CalcUn(Val(Right(Trim(Str(MM)), 1)))
            End Select
        End If
    End Select
End Function


Function Cent_CalcUn(ByRef NumAux)
    Select Case NumAux
        Case 1
            DesCent = DesCent & " UN"
        Case 2
            DesCent = DesCent & " DOS"
        Case 3
            DesCent = DesCent & " TRES"
        Case 4
            DesCent = DesCent & " CUATRO"
        Case 5
            DesCent = DesCent & " CINCO"
        Case 6
            DesCent = DesCent & " SEIS"
        Case 7
            DesCent = DesCent & " SIETE"
        Case 8
            DesCent = DesCent & " OCHO"
        Case 9
            DesCent = DesCent & " NUEVE"
    End Select
End Function


Function Calc_UnCent(ByRef NumAux)
    Select Case NumAux
        Case 1
            DesCent = DesCent & " UN"
        Case 2
            DesCent = DesCent & " DOS"
        Case 3
            DesCent = DesCent & " TRES"
        Case 4
            DesCent = DesCent & " CUATRO"
        Case 5
            DesCent = DesCent & " CINCO"
        Case 6
            DesCent = DesCent & " SEIS"
        Case 7
            DesCent = DesCent & " SIETE"
        Case 8
            DesCent = DesCent & " OCHO"
        Case 9
            DesCent = DesCent & " NUEVE"
    End Select
End Function


