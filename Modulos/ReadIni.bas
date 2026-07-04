Attribute VB_Name = "ReadIni"
Option Explicit
Public DRIVE        As String   'Unidad donde está mapeado
Public SERVIDOR     As String   'Servidor al cual conectarse
Public BASEDATO     As String   'Base de datos a la que te queres conectar
Public USERID     As String   'User id de sesion de la bd
Public PASSWORD     As String   'Password del user
Public DirReport    As String
Public IMPRESORA    As String   'PARA SABER QUE IMPRESORA USA PARA LA FACTURA Y REMITO
Public DirBkp    As String
Public Ayuda As String
Public User As String
Public Doc As String
Public DIGOR_CORE_URL As String
Public DIGOR_PUBLIC_API_KEY As String
Public SERVIDOR_REPORTES As String
Public TURNOS_EXPORTADOS_DIR As String


Public Sub LeoIni()
Dim Pos     As Integer
Dim Largo   As Integer
Dim ValVar  As String
Dim NomVar  As String
Open "C:\WINDOWS\DIGOR.INI" For Input As #1
Do While Not EOF(1)
    Line Input #1, ValVar
    Largo = Len(ValVar)
    If Largo > 3 Then
        Pos = IIf(InStr(1, ValVar, "=") = 0, Largo, InStr(1, ValVar, "="))
        NomVar = UCase(Trim(Left(ValVar, Pos - 2)))
        ValVar = Trim(Right(ValVar, Largo - (Pos)))
        Select Case NomVar
           Case "SERVIDOR"
                SERVIDOR = ValVar
        
           'Case "BASEDATO_TESTING"  'TEST
            Case "BASEDATO"         'PROD
                BASEDATO = ValVar
          
           Case "DRIVE"
                DRIVE = ValVar
                
           Case "DIR_REPORT"
                DirReport = ValVar
           
           Case "IMPRESORA"
                IMPRESORA = ValVar
           
           Case "DIRBKP"
                DirBkp = ValVar
           Case "AYUDA"
                Ayuda = ValVar
           
           Case "USER"
                User = ValVar
           
           Case "DOCTOR"
                Doc = ValVar
           Case "USERID"
                    USERID = ValVar
        
           Case "PASSWORD"
                PASSWORD = ValVar
                
           'Case "DIGOR_CORE_URL_TEST"  'TESTING
           Case "DIGOR_CORE_URL"      'PROD
                DIGOR_CORE_URL = ValVar
            
           'Case "DIGOR_PUBLIC_API_KEY_TEST"    'TESTING
           Case "DIGOR_PUBLIC_API_KEY"        'PROD
              DIGOR_PUBLIC_API_KEY = ValVar
                
            Case "SERVIDOR_REPORTES"
                SERVIDOR_REPORTES = ValVar
                
            Case "TURNOS_EXPORTADOS_DIR"
                    TURNOS_EXPORTADOS_DIR = ValVar

            
        End Select
    End If
Loop
Close #1
End Sub
