Attribute VB_Name = "Module1"
Global Const HelpFinder = &HB           'Mostrar el contenido de los libritos del CNT
Global Const cdlHelpContext = &H1       'Muestra Ayuda acerca de un tema determinado.


'Declaración del API para 32 bits.
Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hWnd As Long, ByVal lpHelpFile As String, _
    ByVal wCommand As Long, ByVal dwData As Long) As Long

'Global Base  As Database
'Global WKS As Workspace

'Global BaseFox As Database
'Global WKSFox As Workspace

Global mOrigen As Boolean
Global mAdentro As Boolean

Global Const StrCon = "ODBC;DSN=ODBC_libro"
Global Const HourGlass = 11
Global Const Normal = 0
Global Const AppName = "Sistema Liquidación de IVA Compras y Ventas"
Global Const Modo_Paso = dbFailOnError
Global Const Agencia = 1
Public txtCtrl As TextBox
Global MontoCuota As Currency
Global MontoAsegurado As Currency
Global FecCont As Date
Public wActivo
Public mLugar
Public TAREA
Public G_Nrodoc
Public G_Tipdoc
Public P_PWD As String
Public P_UID As String
Public frmNada As Form
Public frmDest As Form
Public ssSQL(4) As String
Global mDisco As String



'Public Sub Main()
'    Screen.MousePointer = vbHourglass
'    'Abrir.Show 1
'    'Unload Abrir
'    Menu.Show
'    Screen.MousePointer = vbNormal
'End Sub


