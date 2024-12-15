Attribute VB_Name = "Module5"
   mNombreImpresora

   mNombreImpresora = Printer.DeviceName
    Select Case cboDestino.ListIndex
        Case 0
            Rep.Destination = crptToWindow
        Case 1
            Rep.Destination = crptToPrinter
        Case 2 'pdf
            Rep.Destination = crptToPrinter
            EstableceDefaultPrinter ("PDFCreator")
    End Select

    Rep.Action = 1


    Function

   EstableceDefaultPrinter (mNombreImpresora)

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const HWND_BROADCAST = &HFFFF
Private Const WM_WININICHANGE = &H1A


Private Declare Function GetProfileString Lib "kernel32" _
      Alias "GetProfileStringA" (ByVal lpAppName As String, _
      ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function WriteProfileString Lib "kernel32" _
      Alias "WriteProfileStringA" (ByVal lpszSection As String, _
      ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
               ByVal hWnd As Long, ByVal wMsg As Long, _
               ByVal wParam As Long, lparam As Any) As Long
             
Public Sub EstableceDefaultPrinter(ByVal strNombreImpresora As String)
    'Agradecimientos Serge Baranovsky sergeb@vbcity.com http://www.vbcity.com/
    Dim StrBuffer As String
    Dim lRet As Long
    Dim iDriver As Integer
    Dim iPort As Integer
    Dim DriverName As String
    Dim DriverPort As String
    Dim DeviceLine As String
    Dim PrinterPort As String

    StrBuffer = Space(1024)
    lRet = GetProfileString("PrinterPorts", strNombreImpresora, "", _
           StrBuffer, Len(StrBuffer))
   
   iDriver = InStr(StrBuffer, ",")
   If iDriver > 0 Then
      DriverName = Left(StrBuffer, iDriver - 1)
      iPort = InStr(iDriver + 1, StrBuffer, ",")
      If iPort > 0 Then
         PrinterPort = Mid(StrBuffer, iDriver + 1, iPort - iDriver - 1)
      End If
   End If

   DeviceLine = strNombreImpresora & "," & DriverName & "," & PrinterPort
   lRet = WriteProfileString("windows", "Device", DeviceLine)
   lRet = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub


'[10:20 AM, 4/21/2020] Marcos Martin: mNombreImpresora es una variable string
'[10:21 AM, 4/21/2020] Marcos Martin: lo que hacemos es ver que impresora tenia predeterminada con :mNombreImpresora = Printer.DeviceName
'[10:22 AM, 4/21/2020] Marcos Martin: si me pide pdf cambio esa predeterminada  por la pdfcreator
'[10:22 AM, 4/21/2020] Marcos Martin: y al final vuelvo a poner la que estaba



mNombreImpresora = Printer.DeviceName
    Rep.Destination = crptToPrinter
    Call EstableceDefaultPrinter("PDFCreator")

    strAutoSaveDirectory = App.Path & "\PDF\"
        strAutoSaveFileName = "ORDEN_TRABAJO_" & txtNroPresupueso.Text & "_" & txtPatente.Text & ".pdf"
    
    If Dir(strAutoSaveDirectory & strAutoSaveFileName) <> "" Then Kill (strAutoSaveDirectory & strAutoSaveFileName)
    
    Set objWSH = CreateObject("WScript.Shell")
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\UseAutoSave", 1, "REG_SZ"
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\UseAutoSaveDirectory", 1, "REG_SZ"
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutoSaveDirectory", strAutoSaveDirectory, "REG_SZ"
    objWSH.RegWrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutoSaveFileName", strAutoSaveFileName, "REG_SZ"
    Rep.Action = 1
