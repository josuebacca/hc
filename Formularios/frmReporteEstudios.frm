VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteEstudios 
   Caption         =   "Reporte de estudios"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Generar reporte de estudios entre fechas"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdIrCarpetaReporte 
         Height          =   435
         Left            =   5520
         Picture         =   "frmReporteEstudios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ir a carpeta del estudio"
         Top             =   1320
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   585
         Left            =   3720
         Picture         =   "frmReporteEstudios.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   1065
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   585
         Left            =   2400
         Picture         =   "frmReporteEstudios.frx":0CB4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker dtFechaDesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   151257089
         CurrentDate     =   40071
      End
      Begin MSComCtl2.DTPicker dtFechaHasta 
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   151257089
         CurrentDate     =   40071
      End
      Begin VB.Label lblExito 
         Caption         =   "Carpeta generada correctamente!"
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblError 
         Alignment       =   2  'Center
         Caption         =   "Ocurrió un error al generar la carpeta del reporte  comuníquese con el administrador"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Hasta:"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Fecha Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmReporteEstudios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim link As String
Dim message As String

Private Sub cmdCerrar_Click()
 Unload Me
End Sub

Private Sub cmdGenerar_Click()
    DTFechaDesde.Enabled = False
    DTFechaHasta.Enabled = False
    cmdGenerar.Enabled = False
    cmdCerrar.Enabled = False
    lblExito.Visible = False
    lblError.Visible = False
    cmdIrCarpetaReporte.Visible = False
    link = ""
    message = ""
    PostStudiesReport
End Sub

Private Sub LimpioForm()
    DTFechaDesde.Value = Date
    DTFechaHasta.Value = Date
    cmdGenerar.Enabled = True
    cmdCerrar.Enabled = True
    lblExito.Visible = False
    lblError.Visible = False
    cmdIrCarpetaReporte.Visible = False
    link = ""
    message = ""
End Sub

Private Sub cmdIrCarpetaReporte_Click()
    If link <> "" Then
        Shell "cmd /c start " & link, vbNormalFocus
    Else
        MsgBox "Error al redireccionar a la carpeta del reporte", vbExclamation, "Información"
    End If
End Sub

Private Sub Form_Load()
LimpioForm
End Sub
Private Sub getLinkFromReportJSON(JsonString As String)
    Dim jsonObject As Object
    Dim success As String
    Dim patientObject As Object
    Dim statusCode As String
    Dim ErrorMessage As String
    
    Set jsonObject = JsonConverter.ParseJson(JsonString)
    success = jsonObject("success")
    
    If success = "Verdadero" Then
        link = jsonObject("link")
        message = jsonObject("message")
    Else
        link = ""
        statusCode = jsonObject("statusCode")
        ErrorMessage = jsonObject("message")
        If statusCode <> "500" And ErrorMessage <> "" Then
            message = ErrorMessage
        End If
    End If
End Sub
Public Sub PostStudiesReport()

    Dim request As Object
    Dim responseText As String
    Dim linkDrive As String
    Dim jsonBodyToSend As String
    Dim endpoint As String
    Dim jsonBody As String
    
    endpoint = "/api/v1/studies-report"
    
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Construcción del JSON a enviar
    jsonBody = "{""dateFrom"": """ & Format(DTFechaDesde.Value, "yyyy-mm-dd") & """, " & _
               """dateTo"": """ & Format(DTFechaHasta.Value, "yyyy-mm-dd") & """}"

    
    request.Open "POST", DIGOR_CORE_URL & endpoint, False    'populates object fields
    request.setRequestHeader "Authorization", "Bearer " & DIGOR_PUBLIC_API_KEY
    request.setRequestHeader "Content-Type", "application/json"

    request.send jsonBody
    responseText = request.responseText
    
    'Obtengo el link de la carpeta del estudio y el mensaje de información
    getLinkFromReportJSON (responseText)
    
    If link = "" Then
        ' Hubo error
        lblError.Visible = True
        If message <> "" Then
            ' Hay un mensaje para el usuario, lo mostramos
            lblError.Caption = message
        Else
            'No hay un mensjae para el usuario, el error es técnico
            lblError.Caption = "Ha ocurrido un error, comuníquese con el administrador"
        End If
    End If
    
    If link <> "" Then
        lblExito.Visible = True
        cmdIrCarpetaReporte.Visible = True
        lblExito.Caption = message
    End If
    
    cmdCerrar.Enabled = True
    DTFechaDesde.Enabled = True
    DTFechaHasta.Enabled = True
    
    Set request = Nothing
End Sub

Private Sub dtFechaDesde_Click()
    cmdGenerar.Enabled = True
End Sub
Private Sub dtFechaDesde_Change()
    cmdGenerar.Enabled = True
End Sub
Private Sub dtFechaHasta_Click()
    cmdGenerar.Enabled = True
End Sub
Private Sub dtFechaHasta_Change()
    cmdGenerar.Enabled = True
End Sub
