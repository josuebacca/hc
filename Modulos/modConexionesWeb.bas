Attribute VB_Name = "modConexionesWeb"
Option Explicit

'Minutos que deben pasar como minimo entre un aviso y el siguiente
Private Const MINUTOS_ENTRE_AVISOS As Long = 30

'Fecha/hora del ultimo aviso mostrado (0 = todavia no se aviso)
Private ultimoAviso As Date

'Muestra el aviso de falla de conexion web, como maximo una vez cada MINUTOS_ENTRE_AVISOS
Public Sub AvisarFalloWeb(ByVal nroError As Long, ByVal descError As String, ByVal origen As String)
    Dim codigo As Long
    Dim msg As String

    'Si ya se aviso hace poco, no se vuelve a molestar al usuario
    If ultimoAviso <> 0 Then
        If DateDiff("n", ultimoAviso, Now) < MINUTOS_ENTRE_AVISOS Then Exit Sub
    End If
    ultimoAviso = Now

    'Errores de red de MSXML/WinHTTP (12000-12200): sin internet, DNS, timeout, etc.
    codigo = nroError And &HFFFF&
    If codigo >= 12000 And codigo <= 12200 Then
        msg = "No hay conexión a internet o no se pudo conectar con el servidor web."
    Else
        msg = "El servidor web no está respondiendo correctamente."
    End If

    msg = msg & vbCrLf & vbCrLf & _
          "Puede seguir trabajando normalmente." & vbCrLf & _
          "Los links y estudios de Google Drive pueden no estar actualizados."

    MsgBox msg, vbExclamation, TIT_MSGBOX
End Sub

'Se llama cuando una request termina bien.
'No resetea el aviso: solo deja un registro en la ventana Inmediato
Public Sub MarcarConexionWebOk()
    Debug.Print Now & " - Request web OK"
End Sub

